//! One private native print view. Only the supervised coordinator can send
//! these actions; renderer print requests remain opaque workbook-owner leases.
use crate::locked::Locked;
use serde::{Deserialize, Serialize};
use serde_json::{Value, json};
use std::{
    path::PathBuf,
    sync::{Arc, Mutex},
};
use tauri::{AppHandle, Manager};

pub type Reply = Box<dyn FnOnce(Result<Value, String>) + Send>;

fn print_result(success: bool, printer: bool, page_count: usize) -> Result<Value, String> {
    if success {
        Ok(json!({"submitted":printer,"rendered":!printer,"pageCount":page_count}))
    } else {
        Err("unavailable: Native print operation failed after dispatch".into())
    }
}
fn chooser_result(response: isize, page_count: usize) -> Result<Option<Value>, String> {
    match response {
        0 => Ok(Some(json!({"cancelled":true,"pageCount":page_count}))),
        1 => Ok(None),
        _ => Err("unavailable: Native print chooser returned an unknown result".into()),
    }
}

#[derive(Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct Input {
    id: String,
    root: String,
    sha256: String,
    inspection: String,
    page_width_pt: f64,
    page_height_pt: f64,
    page_count: usize,
}
impl Input {
    fn validate(&self) -> Result<(), String> {
        if uuid::Uuid::parse_str(&self.id)
            .map(|id| id.to_string())
            .ok()
            .as_deref()
            != Some(&self.id)
            || self.sha256.len() != 64
            || !self.sha256.bytes().all(|b| b.is_ascii_hexdigit())
            || self.inspection.is_empty()
            || self.inspection.len() > 65536
            || !(1..=100).contains(&self.page_count)
            || [self.page_width_pt, self.page_height_pt]
                .iter()
                .any(|n| !n.is_finite() || !(100.0..=1500.0).contains(n))
        {
            return Err("validation: Invalid bounded workbook print input".into());
        }
        Ok(())
    }
    fn source(&self) -> Result<(PathBuf, String), String> {
        use sha2::{Digest, Sha256};
        use std::{
            fs,
            io::Read,
            os::unix::fs::{MetadataExt, OpenOptionsExt},
        };
        self.validate()?;
        let root = PathBuf::from(&self.root);
        let metadata = fs::symlink_metadata(&root).map_err(|e| e.to_string())?;
        if !metadata.is_dir()
            || metadata.file_type().is_symlink()
            || metadata.uid() != unsafe { libc::geteuid() }
            || metadata.mode() & 0o077 != 0
            || !root
                .file_name()
                .and_then(|n| n.to_str())
                .is_some_and(|name| name.starts_with(&format!("platform-tauri-print-{}-", self.id)))
        {
            return Err("permission_denied: Print staging is not private owned storage".into());
        }
        let root = root.canonicalize().map_err(|e| e.to_string())?;
        if root.parent()
            != Some(
                std::env::temp_dir()
                    .canonicalize()
                    .map_err(|e| e.to_string())?
                    .as_path(),
            )
        {
            return Err("permission_denied: Print staging is outside temporary storage".into());
        }
        let file = fs::OpenOptions::new()
            .read(true)
            .custom_flags(libc::O_NOFOLLOW)
            .open(root.join("document.html"))
            .map_err(|e| e.to_string())?;
        let meta = file.metadata().map_err(|e| e.to_string())?;
        if !meta.is_file()
            || meta.uid() != unsafe { libc::geteuid() }
            || meta.mode() & 0o077 != 0
            || meta.nlink() != 1
            || meta.len() > 16 * 1024 * 1024
        {
            return Err("permission_denied: Unsafe print document file".into());
        }
        let mut bytes = Vec::new();
        file.take(16 * 1024 * 1024 + 1)
            .read_to_end(&mut bytes)
            .map_err(|e| e.to_string())?;
        if bytes.len() > 16 * 1024 * 1024 || format!("{:x}", Sha256::digest(&bytes)) != self.sha256
        {
            return Err("conflict: Accepted print document changed".into());
        }
        let html =
            String::from_utf8(bytes).map_err(|_| "validation: Print document is not UTF-8")?;
        if !html.to_ascii_lowercase().starts_with("<!doctype html>") || !html.contains("<head>") {
            return Err("validation: Print document is not self-contained HTML".into());
        }
        Ok((root, html))
    }
}

struct Job {
    input: Input,
    root: PathBuf,
    inner: Mutex<JobInner>,
}
struct JobInner {
    closed: bool,
    ready: bool,
    inspecting: bool,
    dispatched: bool,
    completed: bool,
    printer: bool,
    prepare_reply: Option<Reply>,
    dispatch_reply: Option<Reply>,
}
#[derive(Default)]
pub struct PrintJobs(Mutex<Option<Arc<Job>>>);
impl Job {
    fn prepared(&self, result: Result<Value, String>) {
        let reply = {
            let mut state = self.inner.locked();
            if state.closed {
                return;
            }
            state.ready = result.is_ok();
            state.prepare_reply.take()
        };
        if let Some(reply) = reply {
            reply(result)
        }
    }
    fn finished(&self, result: Result<Value, String>) {
        let reply = {
            let mut state = self.inner.locked();
            state.completed = true;
            state.dispatch_reply.take()
        };
        if let Some(reply) = reply {
            reply(result)
        }
    }
}

pub fn handle(app: &AppHandle, action: &str, options: &Value, reply: Reply) {
    #[cfg(target_os = "macos")]
    {
        mac::handle(app, action, options, reply);
    }
    #[cfg(not(target_os = "macos"))]
    {
        let _ = (app, action, options);
        reply(Err(
            "unavailable: Native workbook printing requires the platform print adapter".into(),
        ));
    }
}

#[cfg(target_os = "macos")]
mod mac {
    use super::*;
    use block2::RcBlock;
    use objc2::{
        DefinedClass, MainThreadMarker, MainThreadOnly, define_class, msg_send,
        rc::Retained,
        runtime::{AnyObject, ProtocolObject},
        sel,
    };
    use objc2_app_kit::{
        NSPrintHeaderAndFooter, NSPrintInfo, NSPrintJobSavingURL, NSPrintOperation, NSPrintSaveJob,
        NSWindow,
    };
    use objc2_foundation::{
        NSCopying, NSError, NSNumber, NSObject, NSObjectProtocol, NSPoint, NSRect, NSSize,
        NSString, NSURL,
    };
    use objc2_web_kit::{
        WKContentWorld, WKNavigation, WKNavigationAction, WKNavigationActionPolicy,
        WKNavigationDelegate, WKWebView, WKWebViewConfiguration, WKWebsiteDataStore,
    };
    use std::{cell::RefCell, ffi::c_void};

    struct DelegateState {
        job: Arc<Job>,
        app: AppHandle,
    }
    define_class!(
        #[unsafe(super(NSObject))]
        #[thread_kind = MainThreadOnly]
        #[ivars = DelegateState]
        #[name = "PlatformPrivateWorkbookPrintDelegate"]
        struct PrintDelegate;
        unsafe impl NSObjectProtocol for PrintDelegate {}
        unsafe impl WKNavigationDelegate for PrintDelegate {
            #[unsafe(method(webView:decidePolicyForNavigationAction:decisionHandler:))]
            unsafe fn navigation(&self,_view:&WKWebView,action:&WKNavigationAction,decision:&block2::DynBlock<dyn Fn(WKNavigationActionPolicy)>) {
                let allowed=unsafe { action.request().URL().and_then(|url|url.absoluteString()).is_some_and(|url|url.to_string()=="about:blank") };
                decision.call((if allowed {WKNavigationActionPolicy::Allow}else{WKNavigationActionPolicy::Cancel},));
            }
            #[unsafe(method(webView:didFinishNavigation:))]
            unsafe fn loaded(&self,view:&WKWebView,_navigation:Option<&WKNavigation>) {
                let job=self.ivars().job.clone();
                {let mut state=job.inner.locked();if state.closed || state.inspecting{return}state.inspecting=true;}
                let script=NSString::from_str(&format!("await document.fonts.ready; await Promise.all(Array.from(document.images,image=>image.decode())); const overflow=await {}; return JSON.stringify({{pages:document.querySelectorAll('[data-xlsx-print-page]').length,images:Array.from(document.images).every(image=>image.complete&&image.naturalWidth>0),overflow}});",job.input.inspection));
                let captured=job.clone();
                let callback=RcBlock::new(move|value:*mut AnyObject,error:*mut NSError|{
                    let result=(||{
                        if let Some(error)=unsafe{error.as_ref()}{return Err(format!("validation: Print readiness failed: {}",error.localizedDescription()))}
                        let value=unsafe{value.as_ref()}.and_then(|value|value.downcast_ref::<NSString>()).ok_or("validation: Missing print readiness")?;
                        let result:Value=serde_json::from_str(&value.to_string()).map_err(|e|e.to_string())?;
                        if result["pages"]!=captured.input.page_count || result["images"]!=true || result["overflow"]["pages"]!=captured.input.page_count || !result["overflow"]["failures"].as_array().is_some_and(|errors|errors.is_empty()) {
                            return Err("validation: Print layout is incomplete or text overflows; adjust the accepted page layout".into());
                        }
                        Ok(json!({"prepared":true,"pageCount":captured.input.page_count}))
                    })();captured.prepared(result);
                });
                unsafe{view.callAsyncJavaScript_arguments_inFrame_inContentWorld_completionHandler(&script,None,None,&WKContentWorld::pageWorld(MainThreadMarker::new().expect("WebKit main callback")),Some(&callback));}
            }
            #[unsafe(method(webView:didFailNavigation:withError:))]
            unsafe fn failed(&self,_view:&WKWebView,_navigation:Option<&WKNavigation>,error:&NSError){self.ivars().job.prepared(Err(format!("unavailable: Print view failed: {}",error.localizedDescription())));}
            #[unsafe(method(webView:didFailProvisionalNavigation:withError:))]
            unsafe fn failed_initial(&self,_view:&WKWebView,_navigation:Option<&WKNavigation>,error:&NSError){self.ivars().job.prepared(Err(format!("unavailable: Print view could not load: {}",error.localizedDescription())));}
        }
        impl PrintDelegate {
            #[unsafe(method(printOperation:didRun:contextInfo:))]
            fn printed(&self,_operation:&NSPrintOperation,success:bool,_context:*mut c_void){
                let job=&self.ivars().job;
                let printer=job.inner.locked().printer;
                job.finished(print_result(success,printer,job.input.page_count));
            }
            #[unsafe(method(printPanel:didEnd:contextInfo:))]
            fn chosen(&self,_panel:&objc2_app_kit::NSPrintPanel,response:isize,_context:*mut c_void){
                let job=&self.ivars().job;
                match chooser_result(response,job.input.page_count){Ok(Some(cancelled))=>{job.finished(Ok(cancelled));return},Err(error)=>{job.finished(Err(error));return},Ok(None)=>{}}
                let result=(||{
                    let window=self.ivars().app.get_webview_window("main").ok_or("unavailable: Main print parent is missing")?;
                    let pointer=window.ns_window().map_err(|e|e.to_string())?;
                    let parent=unsafe{&*(pointer as *const NSWindow)};
                    let operation=VIEW.with(|slot|slot.borrow().as_ref().filter(|view|view.id==job.input.id).and_then(|view|view.operation.clone())).ok_or("conflict: Native print view retired")?;
                    unsafe{operation.runOperationModalForWindow_delegate_didRunSelector_contextInfo(parent,Some(self),Some(sel!(printOperation:didRun:contextInfo:)),std::ptr::null_mut());}
                    Ok(())
                })();if let Err(error)=result{job.finished(Err(error));}
            }
        }
    );
    impl PrintDelegate {
        fn new(mtm: MainThreadMarker, job: Arc<Job>, app: AppHandle) -> Retained<Self> {
            let object = Self::alloc(mtm).set_ivars(DelegateState { job, app });
            unsafe { msg_send![super(object), init] }
        }
    }
    struct View {
        id: String,
        view: Retained<WKWebView>,
        delegate: Retained<PrintDelegate>,
        operation: Option<Retained<NSPrintOperation>>,
    }
    thread_local! { static VIEW:RefCell<Option<View>>=const{RefCell::new(None)}; }

    fn get(app: &AppHandle, options: &Value) -> Result<Arc<Job>, String> {
        let id = options["id"]
            .as_str()
            .ok_or("validation: Print identity required")?;
        app.state::<PrintJobs>()
            .0
            .locked()
            .as_ref()
            .filter(|job| job.input.id == id)
            .cloned()
            .ok_or("conflict: Native print lease is missing or retired".into())
    }
    pub(super) fn handle(app: &AppHandle, action: &str, options: &Value, reply: Reply) {
        match action {
            "workbookPrintPrepare" => {
                let result = (|| {
                    let input: Input = serde_json::from_value(options.clone())
                        .map_err(|e| format!("validation: {e}"))?;
                    let (root, html) = input.source()?;
                    let registry = app.state::<PrintJobs>();
                    let mut slot = registry.0.locked();
                    if slot.is_some() {
                        return Err("busy: Another native print view is active".into());
                    }
                    let job = Arc::new(Job {
                        input,
                        root,
                        inner: Mutex::new(JobInner {
                            closed: false,
                            ready: false,
                            inspecting: false,
                            dispatched: false,
                            completed: false,
                            printer: false,
                            prepare_reply: None,
                            dispatch_reply: None,
                        }),
                    });
                    *slot = Some(job.clone());
                    Ok((job, html))
                })();
                let (job, html) = match result {
                    Ok(value) => value,
                    Err(error) => {
                        reply(Err(error));
                        return;
                    }
                };
                job.inner.locked().prepare_reply = Some(reply);
                let captured = job.clone();
                let app_handle = app.clone();
                if let Err(error)=app.run_on_main_thread(move||{
                    if captured.inner.locked().closed{return}
                    let mtm=MainThreadMarker::new().expect("Tauri main thread");
                    unsafe{
                        let config=WKWebViewConfiguration::new(mtm);config.setWebsiteDataStore(&WKWebsiteDataStore::nonPersistentDataStore(mtm));
                        let view=WKWebView::initWithFrame_configuration(WKWebView::alloc(mtm),NSRect::new(NSPoint::new(0.0,0.0),NSSize::new(captured.input.page_width_pt*96.0/72.0,captured.input.page_height_pt*96.0/72.0)),&config);
                        let delegate=PrintDelegate::new(mtm,captured.clone(),app_handle);view.setNavigationDelegate(Some(ProtocolObject::from_ref(&*delegate)));
                        let policy="default-src 'none'; script-src 'none'; img-src data:; style-src 'unsafe-inline'; font-src 'none'; connect-src 'none'; frame-src 'none'; object-src 'none'; base-uri 'none'; form-action 'none'";
                        let html=html.replacen("<head>",&format!("<head><meta http-equiv=\"Content-Security-Policy\" content=\"{policy}\">"),1);
                        view.loadHTMLString_baseURL(&NSString::from_str(&html),None);
                        VIEW.with(|slot|*slot.borrow_mut()=Some(View{id:captured.input.id.clone(),view,delegate,operation:None}));
                    }
                }){job.prepared(Err(error.to_string()));}
            }
            "workbookPrintPdf" | "workbookPrintDispatch" => {
                let job = match get(app, options) {
                    Ok(job) => job,
                    Err(error) => {
                        reply(Err(error));
                        return;
                    }
                };
                {
                    let mut state = job.inner.locked();
                    if state.closed || !state.ready || state.dispatched {
                        reply(Err(
                            "conflict: Native print lease is not ready or was already dispatched"
                                .into(),
                        ));
                        return;
                    }
                    state.dispatched = true;
                    state.printer = action == "workbookPrintDispatch";
                    state.dispatch_reply = Some(reply);
                }
                let captured = job.clone();
                let handle = app.clone();
                if let Err(error) = app.run_on_main_thread(move || {
                    let result = (|| {
                        let window = handle
                            .get_webview_window("main")
                            .ok_or("unavailable: Main print parent is missing")?;
                        let pointer = window.ns_window().map_err(|e| e.to_string())?;
                        let parent = unsafe { &*(pointer as *const NSWindow) };
                        let (operation,delegate,printer)=VIEW.with(|slot| {
                            let mut slot = slot.borrow_mut();
                            let view = slot
                                .as_mut()
                                .filter(|view| view.id == captured.input.id)
                                .ok_or("conflict: Native print view retired")?;
                            let printer = captured.inner.locked().printer;
                            unsafe {
                                let info = NSPrintInfo::sharedPrintInfo().copy();
                                info.setPaperSize(NSSize::new(
                                    captured.input.page_width_pt,
                                    captured.input.page_height_pt,
                                ));
                                info.setLeftMargin(0.0);
                                info.setRightMargin(0.0);
                                info.setTopMargin(0.0);
                                info.setBottomMargin(0.0);
                                info.setHorizontallyCentered(false);
                                info.setVerticallyCentered(false);
                                info.dictionary().setObject_forKey(
                                    &NSNumber::new_bool(false),
                                    ProtocolObject::from_ref(NSPrintHeaderAndFooter),
                                );
                                if !printer {
                                    info.setJobDisposition(NSPrintSaveJob);
                                    info.dictionary().setObject_forKey(
                                        &NSURL::fileURLWithPath(&NSString::from_str(
                                            &captured.root.join("pages.pdf").to_string_lossy(),
                                        )),
                                        ProtocolObject::from_ref(NSPrintJobSavingURL),
                                    );
                                }
                                let operation = view.view.printOperationWithPrintInfo(&info);
                                // Keep WebKit's asynchronous print worker. Forcing
                                // synchronous AppKit rendering stalls WKWebView.
                                operation.setShowsPrintPanel(false);
                                operation.setShowsProgressPanel(false);
                                view.operation = Some(operation.clone());
                                Ok::<_,String>((operation,view.delegate.clone(),printer))
                            }
                        })?;
                        if printer {
                            // Receive chooser Cancel separately: the operation's
                            // success flag combines cancellation and driver errors,
                            // and PrintCore resets its status before that callback.
                            #[allow(deprecated)]
                            unsafe { operation.printPanel().beginSheetWithPrintInfo_modalForWindow_delegate_didEndSelector_contextInfo(
                                &operation.printInfo(),parent,Some(&delegate),Some(sel!(printPanel:didEnd:contextInfo:)),std::ptr::null_mut()); }
                        } else {
                            unsafe { operation.runOperationModalForWindow_delegate_didRunSelector_contextInfo(
                                parent,Some(&delegate),Some(sel!(printOperation:didRun:contextInfo:)),std::ptr::null_mut()); }
                        }
                        Ok::<_,String>(())
                    })();
                    if let Err(error) = result {
                        captured.finished(Err(error));
                    }
                }) {
                    job.finished(Err(error.to_string()));
                }
            }
            "workbookPrintInspect" => reply(get(app, options).and_then(|job| inspect(&job))),
            "workbookPrintClose" => {
                let job = match get(app, options) {
                    Ok(job) => job,
                    Err(error) => {
                        reply(Err(error));
                        return;
                    }
                };
                {
                    let mut state = job.inner.locked();
                    if state.dispatched && !state.completed {
                        reply(Err(
                            "busy: Native print operation is still owned by the print system"
                                .into(),
                        ));
                        return;
                    }
                    state.closed = true;
                }
                job.prepared(Err("cancelled: Print preparation cancelled".into()));
                // closed suppresses late callbacks, but the original prepare
                // caller still needs its terminal reply before deleting staging.
                if let Some(pending) = job.inner.locked().prepare_reply.take() {
                    pending(Err("cancelled: Print preparation cancelled".into()))
                }
                let id = job.input.id.clone();
                let handle = app.clone();
                // A queued disposal is not a released view. Acknowledge only
                // after WebKit stops using the staging files on its main thread.
                let response = Arc::new(Mutex::new(Some(reply)));
                let captured = response.clone();
                if let Err(error) = app.run_on_main_thread(move || {
                    VIEW.with(|slot| {
                        let mut slot = slot.borrow_mut();
                        if slot.as_ref().is_some_and(|view| view.id == id) {
                            if let Some(view) = slot.take() {
                                unsafe {
                                    view.view.setNavigationDelegate(None);
                                    view.view.stopLoading();
                                }
                            }
                        }
                    });
                    let registry = handle.state::<PrintJobs>();
                    let mut slot = registry.0.locked();
                    if slot.as_ref().is_some_and(|job| job.input.id == id) {
                        *slot = None
                    }
                    drop(slot);
                    if let Some(reply) = captured.locked().take() {
                        reply(Ok(json!({"closed":true})))
                    }
                }) {
                    if let Some(reply) = response.locked().take() {
                        reply(Err(error.to_string()))
                    }
                }
            }
            _ => reply(Err("validation: Unknown native print action".into())),
        }
    }
    fn inspect(job: &Job) -> Result<Value, String> {
        use objc2_core_foundation::{CFString, CFURL, CFURLPathStyle};
        use objc2_core_graphics::{CGPDFBox, CGPDFDocument, CGPDFPage};
        use std::{fs, os::unix::fs::MetadataExt};
        let state = job.inner.locked();
        if state.closed || state.printer || !state.completed {
            return Err("conflict: Native PDF is not ready".into());
        }
        drop(state);
        let path = job.root.join("pages.pdf");
        let metadata = fs::symlink_metadata(&path).map_err(|e| e.to_string())?;
        if !metadata.is_file()
            || metadata.file_type().is_symlink()
            || metadata.len() < 8
            || metadata.len() > 16 * 1024 * 1024
            || metadata.nlink() != 1
        {
            return Err("validation: Invalid bounded native PDF".into());
        }
        let url = CFURL::with_file_system_path(
            None,
            Some(&CFString::from_str(&path.to_string_lossy())),
            CFURLPathStyle::CFURLPOSIXPathStyle,
            false,
        )
        .ok_or("validation: PDF URL unavailable")?;
        let document =
            CGPDFDocument::with_url(Some(&url)).ok_or("validation: Native output is not a PDF")?;
        if CGPDFDocument::number_of_pages(Some(&document)) != job.input.page_count {
            return Err("validation: Native PDF page count differs from accepted layout".into());
        }
        for index in 1..=job.input.page_count {
            let page = CGPDFDocument::page(Some(&document), index)
                .ok_or("validation: Native PDF page missing")?;
            let rectangle = CGPDFPage::box_rect(Some(&page), CGPDFBox::MediaBox);
            if (rectangle.size.width - job.input.page_width_pt).abs() > 0.5
                || (rectangle.size.height - job.input.page_height_pt).abs() > 0.5
            {
                return Err(
                    "validation: Native PDF paper size differs from accepted layout".into(),
                );
            }
        }
        Ok(
            json!({"pageCount":job.input.page_count,"pageWidthPt":job.input.page_width_pt,"pageHeightPt":job.input.page_height_pt,"bytes":metadata.len()}),
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn printer_cancel_is_distinct_from_native_failures() {
        assert_eq!(chooser_result(0, 2).unwrap().unwrap()["cancelled"], true);
        assert!(chooser_result(1, 2).unwrap().is_none());
        assert!(chooser_result(-1, 2).is_err());
        assert!(
            print_result(false, true, 2)
                .unwrap_err()
                .starts_with("unavailable:")
        );
        assert!(print_result(false, false, 2).is_err());
        assert_eq!(print_result(true, false, 2).unwrap()["rendered"], true);
    }
    #[test]
    fn print_bounds_and_identity_are_checked_before_reading() {
        let mut input = Input {
            id: uuid::Uuid::new_v4().to_string(),
            root: "/tmp/not-owned".into(),
            sha256: "0".repeat(64),
            inspection: "({pages:1,failures:[]})".into(),
            page_width_pt: 595.0,
            page_height_pt: 842.0,
            page_count: 1,
        };
        assert!(input.validate().is_ok());
        input.page_count = 101;
        assert!(input.validate().is_err());
        input.page_count = 1;
        input.page_width_pt = f64::NAN;
        assert!(input.validate().is_err());
        input.page_width_pt = 595.0;
        input.id = "../escape".into();
        assert!(input.validate().is_err());
    }
    #[test]
    fn print_source_requires_private_regular_unchanged_files() {
        use sha2::{Digest, Sha256};
        use std::{
            fs,
            os::unix::fs::{PermissionsExt, symlink},
        };
        let id = uuid::Uuid::new_v4().to_string();
        let root = std::env::temp_dir().join(format!("platform-tauri-print-{id}-test"));
        fs::create_dir(&root).unwrap();
        struct Cleanup(PathBuf);
        impl Drop for Cleanup {
            fn drop(&mut self) {
                let _ = fs::remove_dir_all(&self.0);
            }
        }
        let _cleanup = Cleanup(root.clone());
        fs::set_permissions(&root, fs::Permissions::from_mode(0o700)).unwrap();
        let html = "<!doctype html><html><head></head><body>中文</body></html>";
        let file = root.join("document.html");
        fs::write(&file, html).unwrap();
        fs::set_permissions(&file, fs::Permissions::from_mode(0o600)).unwrap();
        let input = Input {
            id,
            root: root.to_string_lossy().into(),
            sha256: format!("{:x}", Sha256::digest(html.as_bytes())),
            inspection: "({pages:1,failures:[]})".into(),
            page_width_pt: 595.0,
            page_height_pt: 842.0,
            page_count: 1,
        };
        assert_eq!(input.source().unwrap().1, html);
        fs::set_permissions(&file, fs::Permissions::from_mode(0o644)).unwrap();
        assert!(
            input
                .source()
                .unwrap_err()
                .starts_with("permission_denied:")
        );
        fs::set_permissions(&file, fs::Permissions::from_mode(0o600)).unwrap();
        fs::write(&file, format!("{html}changed")).unwrap();
        assert!(input.source().unwrap_err().starts_with("conflict:"));
        fs::write(&file, html).unwrap();
        fs::hard_link(&file, root.join("linked.html")).unwrap();
        assert!(
            input
                .source()
                .unwrap_err()
                .starts_with("permission_denied:")
        );
        fs::remove_file(root.join("linked.html")).unwrap();
        fs::rename(&file, root.join("original.html")).unwrap();
        symlink("original.html", &file).unwrap();
        assert!(input.source().is_err());
        fs::remove_file(&file).unwrap();
        fs::rename(root.join("original.html"), &file).unwrap();
        fs::set_permissions(&root, fs::Permissions::from_mode(0o755)).unwrap();
        assert!(
            input
                .source()
                .unwrap_err()
                .starts_with("permission_denied:")
        );
    }
}
