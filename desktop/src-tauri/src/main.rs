// Supervision is in host/; workbook policy remains here.
const COORDINATOR_ENTRY: &str = "desktop/out/main/index.js";
/// The prefix of this editor's environment variables (host/profile.rs).
const ENV_PREFIX: &str = "XLSX_EDITOR";
const CHANNEL_PREFIX: &str = "xlsx:";
fn configure_coordinator(command: &mut std::process::Command, root: &std::path::Path, resources: &std::path::Path, packaged: bool) {
    command.env("XLSX_EDITOR_NATIVE_BIN", if packaged { resources.join("native/workbook-engine") } else { root.join("native/bin/darwin-arm64/workbook-engine") });
}
mod dialogs;
mod host;
mod menu;
mod native;
mod workbook_print;
use host::{backend, locked, memory, profile};
use serde_json::{Value, json};
use std::sync::{
    Arc, Mutex,
    atomic::{AtomicBool, Ordering},
};
use tauri::{Manager, State, WebviewWindow};
use tauri_plugin_dialog::{DialogExt, MessageDialogButtons};
#[derive(Default)]
pub struct QuitState {
    approved: AtomicBool,
    pending: AtomicBool,
    dialog: AtomicBool,
    zoom: Mutex<f64>,
}
/// What this editor gives the coordinator supervision (host/backend.rs).
struct Editor;
impl backend::Shell for Editor {
    const MANIFEST_DIR: &'static str = env!("CARGO_MANIFEST_DIR");
    const COORDINATOR_ENTRY: &'static str = COORDINATOR_ENTRY;
    const ENV_PREFIX: &'static str = ENV_PREFIX;
    const CHANNEL_PREFIX: &'static str = CHANNEL_PREFIX;
    fn configure_coordinator(command: &mut std::process::Command, root: &std::path::Path, resources: &std::path::Path, packaged: bool) {
        configure_coordinator(command, root, resources, packaged)
    }
    fn dispatch_native(app: &tauri::AppHandle, request: &Value, reply: backend::NativeReply) {
        dispatch_native(app, request, reply)
    }
    fn quit_approved(app: &tauri::AppHandle) {
        app.state::<QuitState>().approved.store(true, Ordering::Release);
    }
    fn quit_settled(app: &tauri::AppHandle) {
        app.state::<QuitState>().pending.store(false, Ordering::Release);
    }
}
fn app_origin(url: &tauri::Url) -> bool {
    (url.scheme() == "tauri" && url.host_str() == Some("localhost"))
        || (url.scheme() == "http" && url.host_str() == Some("tauri.localhost"))
}
fn navigation_allowed(url: &tauri::Url) -> bool {
    // WKWebView applies this callback to iframe navigations too. The original
    // print preview uses a script-disabled srcdoc frame. Permit only its internal
    // about URLs; those URLs still cannot pass the privileged root-origin guard.
    app_origin(url) || matches!(url.as_str(), "about:blank" | "about:srcdoc")
}
#[cfg(test)]
mod navigation_tests {
    use super::*;
    #[test]
    fn preview_frames_do_not_gain_root_ipc_authority() {
        for value in ["about:blank", "about:srcdoc"] {
            let url = tauri::Url::parse(value).unwrap();
            assert!(navigation_allowed(&url));
            assert!(!app_origin(&url));
        }
        for value in [
            "https://example.invalid/",
            "file:///tmp/untrusted.html",
            "data:text/html,untrusted",
            "about:config",
            "tauri://untrusted/",
        ] {
            let url = tauri::Url::parse(value).unwrap();
            assert!(!navigation_allowed(&url));
            assert!(!app_origin(&url));
        }
        for value in ["tauri://localhost/", "http://tauri.localhost/"] {
            let url = tauri::Url::parse(value).unwrap();
            assert!(navigation_allowed(&url));
            assert!(app_origin(&url));
        }
    }
}
fn trusted(window: &WebviewWindow) -> Result<(), String> {
    if window.label() != "main" {
        return Err("permission_denied: Unknown standalone window".into());
    }
    let url = window.url().map_err(|e| e.to_string())?;
    if app_origin(&url) {
        Ok(())
    } else {
        Err("permission_denied: Untrusted standalone origin".into())
    }
}
#[tauri::command]
async fn standalone_initialize(
    window: WebviewWindow,
    backend: State<'_, Arc<backend::Backend>>,
    token: String,
) -> Result<Value, String> {
    trusted(&window)?;
    backend.attach(token.clone())?;
    let mut value = backend.invoke("standalone:initialize", vec![]).await?;
    backend.owns(&token)?;
    value["windowCount"] = window.app_handle().webview_windows().len().into();
    Ok(value)
}
#[tauri::command]
async fn standalone_invoke(
    window: WebviewWindow,
    backend: State<'_, Arc<backend::Backend>>,
    token: String,
    channel: String,
    args: Vec<Value>,
) -> Result<Value, String> {
    trusted(&window)?;
    backend.owns(&token)?;
    if !channel.starts_with("xlsx:") || channel.len() > 100 || args.len() > 8 {
        return Err("permission_denied: Unknown standalone operation".into());
    }
    let result = backend.invoke(&channel, args).await?;
    backend.owns(&token)?;
    Ok(result)
}
#[tauri::command]
fn standalone_event_receipt(
    window: WebviewWindow,
    backend: State<'_, Arc<backend::Backend>>,
    receipt: u64,
) -> Result<(), String> {
    trusted(&window)?;
    backend.acknowledge(receipt)
}
#[tauri::command]
fn standalone_quit(window: WebviewWindow) -> Result<(), String> {
    trusted(&window)?;
    window.app_handle().exit(0);
    Ok(())
}
#[tauri::command]
fn standalone_memory(window: WebviewWindow) -> Result<Value, String> {
    trusted(&window)?;
    serde_json::to_value(memory::sample()?).map_err(|e| e.to_string())
}
#[tauri::command]
fn standalone_status(
    window: WebviewWindow,
    backend: State<'_, Arc<backend::Backend>>,
) -> Result<Value, String> {
    trusted(&window)?;
    let state = window.app_handle().state::<QuitState>();
    Ok(
        json!({"windowCount":window.app_handle().webview_windows().len(),"available":backend.available(),"pendingQuit":state.pending.load(Ordering::Acquire),"dialog":state.dialog.load(Ordering::Acquire)}),
    )
}
#[tauri::command]
fn verification_report(window: WebviewWindow, report: Value) -> Result<(), String> {
    trusted(&window)?;
    if !cfg!(debug_assertions) || std::env::var("XLSX_EDITOR_GUI_TESTS").as_deref() != Ok("1") {
        return Err("permission_denied: Verification disabled".into());
    }
    let path =
        std::env::var("XLSX_EDITOR_GUI_REPORT").map_err(|_| "validation: Verification path missing")?;
    std::fs::write(
        path,
        serde_json::to_vec_pretty(&report).map_err(|e| e.to_string())?,
    )
    .map_err(|e| e.to_string())
}
fn dispatch_native(app: &tauri::AppHandle, request: &Value, reply: workbook_print::Reply) {
    match request["action"].as_str() {
        Some("standaloneMenu") => {
            reply(menu::install(app, &request["options"]["items"]).map(|_| Value::Null))
        }
        Some("standaloneReload") => reply(
            app.get_webview_window("main")
                .ok_or("unavailable: Standalone window missing".into())
                .and_then(|window| {
                    window
                        .reload()
                        .map(|_| Value::Null)
                        .map_err(|e| e.to_string())
                }),
        ),
        Some("messageDialog")
            if request["options"]["buttons"]
                .as_array()
                .is_some_and(|buttons| buttons.len() == 4) =>
        {
            dialogs::dispatch(app, &request["options"], reply)
        }
        _ => native::dispatch(app, request, reply),
    }
}
fn native_files(app: &tauri::AppHandle, files: Vec<String>) {
    if files.is_empty() {
        return;
    }
    if let Some(window) = app.get_webview_window("main") {
        let _ = window.show();
        let _ = window.set_focus();
    }
    let backend = app.state::<Arc<backend::Backend>>().inner().clone();
    tauri::async_runtime::spawn(async move {
        if let Err(error) = backend.invoke("standalone:files", vec![json!(files)]).await {
            eprintln!("{error}")
        }
    });
}
fn launch_files(args: impl IntoIterator<Item = String>, cwd: &std::path::Path) -> Vec<String> {
    let mut options = false;
    let mut files = Vec::new();
    for arg in args {
        if options {
            options = false;
            continue;
        }
        if arg == "--profile" || arg == "--restart-after" {
            options = true;
            continue;
        }
        if arg.starts_with('-') {
            continue;
        }
        let path = std::path::PathBuf::from(arg);
        let path = if path.is_absolute() {
            path
        } else {
            cwd.join(path)
        };
        if path
            .extension()
            .and_then(|value| value.to_str())
            .is_some_and(|value| {
                ["xlsx", "xlsm", "xls", "xlsb", "csv", "tsv"]
                    .contains(&value.to_ascii_lowercase().as_str())
            })
        {
            files.push(path.to_string_lossy().into_owned())
        }
    }
    files.truncate(16);
    files
}
fn main() {
    if std::env::args().any(|argument| argument == "--print-host-runtime") {
        println!("{{\"kind\":\"bun\"}}");
        return;
    }

    if let Err(error) = profile::wait_for_restart_parent() {
        eprintln!("{error}");
        return;
    }
    tauri::Builder::default()
  .plugin(tauri_plugin_single_instance::init(|app,args,cwd|{let files=launch_files(args.into_iter().skip(1),std::path::Path::new(&cwd));native_files(app,files);if let Some(window)=app.get_webview_window("main"){let _=window.show();let _=window.set_focus();}}))
  .plugin(tauri_plugin_dialog::init()).plugin(tauri_plugin_clipboard_manager::init()).plugin(tauri_plugin_opener::init())
  .on_menu_event(|app,event|menu::event(app,event.id().as_ref()))
  .setup(|app|{
   let selected=profile::selected(ENV_PREFIX,app.path().app_data_dir()?).map_err(std::io::Error::other)?;
   let profile=profile::Profile::acquire(selected).map_err(std::io::Error::other)?;let store=profile.data_store_id;
   app.manage(profile);app.manage(workbook_print::PrintJobs::default());app.manage(QuitState{zoom:Mutex::new(1.0),..Default::default()});
   let window=app.config().app.windows.first().ok_or("Standalone window configuration missing")?;
   tauri::WebviewWindowBuilder::from_config(app,window)?.data_store_identifier(store).on_navigation(navigation_allowed).build()?;
   let backend=backend::Backend::start::<Editor>(app.handle(),&app.state::<profile::Profile>()).map_err(std::io::Error::other)?;app.manage(backend);
   native_files(app.handle(),launch_files(std::env::args().skip(1),&std::env::current_dir()?));Ok(())
  })
  .on_page_load(|webview,payload|{if cfg!(debug_assertions)&&std::env::var("XLSX_EDITOR_GUI_TESTS").as_deref()==Ok("1")&&payload.event()==tauri::webview::PageLoadEvent::Finished{if let Ok(path)=std::env::var("XLSX_EDITOR_GUI_SCRIPT"){if let Ok(script)=std::fs::read_to_string(path){let _=webview.eval(script);}}}})
  .invoke_handler(tauri::generate_handler![standalone_initialize,standalone_invoke,standalone_event_receipt,standalone_quit,standalone_memory,standalone_status,verification_report])
  .on_window_event(|window,event|{if let tauri::WindowEvent::CloseRequested{api,..}=event{api.prevent_close();window.app_handle().exit(0);}})
  .build(tauri::generate_context!()).expect("Standalone Tauri shell failed")
  .run(|app,event|{
   if matches!(event,tauri::RunEvent::Exit){app.state::<Arc<backend::Backend>>().stop();app.state::<profile::Profile>().release();if cfg!(debug_assertions)&&std::env::var("XLSX_EDITOR_GUI_TESTS").as_deref()==Ok("1"){if let Ok(path)=std::env::var("XLSX_EDITOR_GUI_EXIT_REPORT"){let _=std::fs::write(path,json!({"approvedQuit":app.state::<QuitState>().approved.load(Ordering::Acquire)}).to_string());}}}
   if let tauri::RunEvent::ExitRequested{ref api,..}=event{
    let state=app.state::<QuitState>();if state.approved.load(Ordering::Acquire){return}api.prevent_exit();if state.pending.swap(true,Ordering::AcqRel){return}
    let backend=app.state::<Arc<backend::Backend>>().inner().clone();let handle=app.clone();
    if backend.available(){tauri::async_runtime::spawn(async move{if let Err(error)=backend.invoke("standalone:close",vec![]).await{handle.state::<QuitState>().pending.store(false,Ordering::Release);handle.dialog().message(format!("Recovery could not finish: {error}")).show(|_|{});}});}
    else{let handle=app.clone();app.dialog().message("The spreadsheet service stopped. Saved files and accepted recovery data are retained.").buttons(MessageDialogButtons::OkCancelCustom("Quit editor".into(),"Keep open".into())).show(move|accepted|{let state=handle.state::<QuitState>();if accepted{state.approved.store(true,Ordering::Release);handle.exit(0)}else{state.pending.store(false,Ordering::Release)}});}
   }
   if let tauri::RunEvent::Opened{ref urls}=event{native_files(app,urls.iter().filter_map(|url|url.to_file_path().ok()).map(|path|path.to_string_lossy().into_owned()).collect());}
   if let tauri::RunEvent::Reopen{..}=event{if let Some(window)=app.get_webview_window("main"){let _=window.show();let _=window.set_focus();}}
  });
}
