//! Four-way Save/Cancel/Discard/Keep decisions preserve the standalone contract.
use serde_json::{Value, json};
use std::sync::{Arc, Mutex, atomic::Ordering};
use tauri::{AppHandle, Manager};
pub fn dispatch(app: &AppHandle, options: &Value, reply: crate::workbook_print::Reply) {
    let labels = options["buttons"]
        .as_array()
        .filter(|labels| labels.len() == 4)
        .and_then(|labels| {
            labels
                .iter()
                .map(|label| {
                    label
                        .as_str()
                        .filter(|label| !label.is_empty() && label.len() <= 128)
                        .map(str::to_owned)
                })
                .collect::<Option<Vec<_>>>()
        });
    let Some(labels) = labels else {
        reply(Err(
            "validation: Invalid four-way standalone decision".into()
        ));
        return;
    };
    let message = options["message"].as_str().unwrap_or("").to_owned();
    let detail = options["detail"].as_str().unwrap_or("").to_owned();
    let default = options["defaultId"].as_u64().unwrap_or(0) as usize;
    let cancel = options["cancelId"].as_u64().unwrap_or(1) as usize;
    if message.len() + detail.len() > 16384 || default >= 4 || cancel >= 4 {
        reply(Err("validation: Invalid standalone dialog bounds".into()));
        return;
    }
    if app
        .state::<crate::QuitState>()
        .dialog
        .swap(true, Ordering::AcqRel)
    {
        reply(Err("busy: A standalone decision is already open".into()));
        return;
    }
    let response = Arc::new(Mutex::new(Some(reply)));
    let failed = response.clone();
    let handle = app.clone();
    if let Err(error) = app.run_on_main_thread(move || {
        let result = (|| -> Result<(), String> {
            use block2::RcBlock;
            use objc2::{MainThreadMarker, rc::Retained};
            use objc2_app_kit::{NSAlert, NSWindow};
            use objc2_foundation::NSString;
            let mtm = MainThreadMarker::new()
                .ok_or("unavailable: Native decision requires the main thread")?;
            let window = handle
                .get_webview_window("main")
                .ok_or("unavailable: Standalone window missing")?;
            let pointer = window.ns_window().map_err(|e| e.to_string())?;
            let parent = unsafe { Retained::<NSWindow>::retain(pointer.cast()) }
                .ok_or("unavailable: Native parent missing")?;
            let alert = NSAlert::new(mtm);
            alert.setMessageText(&NSString::from_str(&message));
            alert.setInformativeText(&NSString::from_str(&detail));
            for (index, label) in labels.iter().enumerate() {
                let button = alert.addButtonWithTitle(&NSString::from_str(label));
                button.setKeyEquivalent(&NSString::from_str(if index == cancel {
                    "\u{1b}"
                } else if index == default {
                    "\r"
                } else {
                    ""
                }));
            }
            let response = response.clone();
            let handle = handle.clone();
            let retained = alert.clone();
            let callback = RcBlock::new(move |value: isize| {
                let _keep_alive = &retained;
                handle
                    .state::<crate::QuitState>()
                    .dialog
                    .store(false, Ordering::Release);
                if let Some(reply) = response.lock().unwrap().take() {
                    let selected = value - 1000;
                    reply(if (0..4).contains(&selected) {
                        Ok(json!({"response":selected}))
                    } else {
                        Err("unavailable: Native decision returned an unknown result".into())
                    });
                }
            });
            alert.beginSheetModalForWindow_completionHandler(&parent, Some(&callback));
            Ok(())
        })();
        if let Err(error) = result {
            handle
                .state::<crate::QuitState>()
                .dialog
                .store(false, Ordering::Release);
            if let Some(reply) = response.lock().unwrap().take() {
                reply(Err(error));
            }
        }
    }) {
        app.state::<crate::QuitState>()
            .dialog
            .store(false, Ordering::Release);
        if let Some(reply) = failed.lock().unwrap().take() {
            reply(Err(error.to_string()));
        }
    }
}
