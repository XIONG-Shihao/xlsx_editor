use serde_json::{Value, json};
use tauri::{AppHandle, Manager};
use tauri_plugin_clipboard_manager::ClipboardExt;
use tauri_plugin_dialog::{DialogExt, MessageDialogButtons, MessageDialogResult};
use tauri_plugin_opener::OpenerExt;

fn text<'a>(options: &'a Value, name: &str) -> Result<&'a str, String> {
    options
        .get(name)
        .and_then(Value::as_str)
        .ok_or_else(|| format!("validation: Missing {name}"))
}

fn clipboard_text(
    result: tauri_plugin_clipboard_manager::Result<String>,
) -> Result<String, String> {
    match result {
        Ok(text) => Ok(text),
        // The plugin erases arboard's variant to a string. Compare against the
        // same locked arboard variant, not a broad substring or every failure.
        // Electron's readText returns empty text for an empty/non-text clipboard.
        Err(tauri_plugin_clipboard_manager::Error::Clipboard(detail))
            if detail == arboard::Error::ContentNotAvailable.to_string() =>
        {
            Ok(String::new())
        }
        Err(error) => Err(error.to_string()),
    }
}
pub fn action(app: &AppHandle, request: &Value) -> Result<Value, String> {
    let action = request
        .get("action")
        .and_then(Value::as_str)
        .ok_or("validation: Missing native action")?;
    let options = &request["options"];
    let window = app
        .get_webview_window("main")
        .ok_or("unavailable: Main window is unavailable")?;
    match action {
        "processMemory" => {
            serde_json::to_value(crate::memory::sample()?).map_err(|e| e.to_string())
        }
        "reveal" => {
            if !window.is_visible().map_err(|e| e.to_string())?
                || window.is_minimized().map_err(|e| e.to_string())?
            {
                window.unminimize().map_err(|e| e.to_string())?;
                window.show().map_err(|e| e.to_string())?;
                window.set_focus().map_err(|e| e.to_string())?;
            }
            Ok(Value::Null)
        }
        "show" => {
            window.show().map_err(|e| e.to_string())?;
            Ok(Value::Null)
        }
        "focus" => {
            window.set_focus().map_err(|e| e.to_string())?;
            Ok(Value::Null)
        }
        "hide" => {
            window.hide().map_err(|e| e.to_string())?;
            Ok(Value::Null)
        }
        "clipboardRead" => clipboard_text(app.clipboard().read_text()).map(Value::String),
        "clipboardWrite" => {
            app.clipboard()
                .write_text(text(options, "text")?)
                .map_err(|e| e.to_string())?;
            Ok(Value::Null)
        }
        "openExternal" => {
            let url = text(options, "url")?;
            let parsed = tauri::Url::parse(url).map_err(|e| e.to_string())?;
            if !matches!(parsed.scheme(), "http" | "https") || parsed.host_str().is_none() {
                return Err("permission_denied: External URL is not HTTP(S)".into());
            }
            app.opener()
                .open_url(url, None::<&str>)
                .map_err(|e| e.to_string())?;
            Ok(Value::Null)
        }
        "openPath" => {
            app.opener()
                .open_path(text(options, "path")?, None::<&str>)
                .map_err(|e| e.to_string())?;
            Ok(Value::String(String::new()))
        }
        "revealPath" => {
            app.opener()
                .reveal_item_in_dir(text(options, "path")?)
                .map_err(|e| e.to_string())?;
            Ok(Value::Null)
        }
        "openDialog" => {
            let mut dialog = app.dialog().file().set_parent(&window);
            if let Some(title) = options.get("title").and_then(Value::as_str) {
                dialog = dialog.set_title(title);
            }
            if let Some(filters) = options.get("filters").and_then(Value::as_array) {
                for filter in filters {
                    let extensions: Vec<&str> = filter["extensions"]
                        .as_array()
                        .map(|v| v.iter().filter_map(Value::as_str).collect())
                        .unwrap_or_default();
                    dialog =
                        dialog.add_filter(filter["name"].as_str().unwrap_or("Files"), &extensions);
                }
            }
            let properties = options.get("properties").and_then(Value::as_array);
            let has = |name| {
                properties.is_some_and(|items| items.iter().any(|item| item.as_str() == Some(name)))
            };
            let selected = if has("openDirectory") {
                if has("multiSelections") {
                    dialog.blocking_pick_folders()
                } else {
                    dialog.blocking_pick_folder().map(|file| vec![file])
                }
            } else if has("multiSelections") {
                dialog.blocking_pick_files()
            } else {
                dialog.blocking_pick_file().map(|file| vec![file])
            };
            let paths: Vec<String> = selected
                .unwrap_or_default()
                .into_iter()
                .map(|file| {
                    file.into_path()
                        .map(|path| path.to_string_lossy().into_owned())
                })
                .collect::<Result<_, _>>()
                .map_err(|e| e.to_string())?;
            Ok(json!({"canceled":paths.is_empty(),"filePaths":paths}))
        }
        "saveDialog" => {
            let mut dialog = app.dialog().file().set_parent(&window);
            if let Some(title) = options.get("title").and_then(Value::as_str) {
                dialog = dialog.set_title(title);
            }
            if let Some(default) = options.get("defaultPath").and_then(Value::as_str) {
                let path = std::path::Path::new(default);
                if let Some(name) = path.file_name().and_then(|name| name.to_str()) {
                    dialog = dialog.set_file_name(name);
                }
                if let Some(directory) = path.parent().filter(|path| path.is_dir()) {
                    dialog = dialog.set_directory(directory);
                }
            }
            let file = dialog
                .blocking_save_file()
                .map(|file| file.into_path())
                .transpose()
                .map_err(|e| e.to_string())?;
            Ok(
                json!({"canceled":file.is_none(),"filePath":file.map(|path|path.to_string_lossy().into_owned())}),
            )
        }
        "messageDialog" => {
            let message = format!(
                "{}\n{}",
                options["message"].as_str().unwrap_or(""),
                options["detail"].as_str().unwrap_or("")
            );
            let labels: Vec<String> = options
                .get("buttons")
                .and_then(Value::as_array)
                .map(|items| {
                    items
                        .iter()
                        .filter_map(Value::as_str)
                        .map(str::to_owned)
                        .collect()
                })
                .unwrap_or_else(|| vec!["OK".into()]);
            let buttons = match labels.as_slice() {
                [one] => MessageDialogButtons::OkCustom(one.clone()),
                [one, two] => MessageDialogButtons::OkCancelCustom(one.clone(), two.clone()),
                [one, two, three] => {
                    MessageDialogButtons::YesNoCancelCustom(one.clone(), two.clone(), three.clone())
                }
                _ => {
                    return Err(
                        "validation: Native message dialogs support one to three buttons".into(),
                    );
                }
            };
            let result = app
                .dialog()
                .message(message)
                .buttons(buttons)
                .blocking_show_with_result();
            let index = match result {
                MessageDialogResult::Custom(label) => labels
                    .iter()
                    .position(|item| item == &label)
                    .unwrap_or(labels.len() - 1),
                MessageDialogResult::No => 1,
                MessageDialogResult::Cancel => options
                    .get("cancelId")
                    .and_then(Value::as_u64)
                    .map(|value| value as usize)
                    .unwrap_or(labels.len() - 1),
                _ => 0,
            };
            Ok(json!({"response":index}))
        }
        _ => Err(format!(
            "unavailable: Native action has not been migrated: {action}"
        )),
    }
}

pub fn dispatch(app: &AppHandle, request: &Value, reply: crate::workbook_print::Reply) {
    let action = request["action"].as_str().unwrap_or("");
    if action.starts_with("workbookPrint") {
        crate::workbook_print::handle(app, action, &request["options"], reply)
    } else {
        reply(self::action(app, request))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn absent_clipboard_text_is_empty_but_other_errors_remain_errors() {
        assert_eq!(
            clipboard_text(Ok("中文\t2\n3\t4".into())).unwrap(),
            "中文\t2\n3\t4"
        );
        assert_eq!(
            clipboard_text(Err(arboard::Error::ContentNotAvailable.into())).unwrap(),
            ""
        );
        for detail in [
            "Clipboard occupied",
            "Clipboard permission denied",
            "Unsupported clipboard backend",
        ] {
            assert_eq!(
                clipboard_text(Err(tauri_plugin_clipboard_manager::Error::Clipboard(
                    detail.into()
                )))
                .unwrap_err(),
                detail
            );
        }
    }
}
