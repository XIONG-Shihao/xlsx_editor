use serde_json::Value;
use tauri::{
    AppHandle, Manager,
    menu::{Menu, MenuItem, MenuItemKind, PredefinedMenuItem, Submenu},
};
pub const QUIT: &str = "standalone.quit";
pub fn install(app: &AppHandle, items: &Value) -> Result<(), String> {
    let entries = items
        .as_array()
        .filter(|items| items.len() <= 8)
        .ok_or("validation: Invalid standalone menus")?;
    let menu = Menu::new(app).map_err(|e| e.to_string())?;
    for entry in entries {
        match entry["role"].as_str() {
            Some("appMenu") => {
                let default = Menu::default(app).map_err(|e| e.to_string())?;
                let entries = default.items().map_err(|e| e.to_string())?;
                let Some(MenuItemKind::Submenu(application)) = entries.first() else {
                    return Err("unavailable: Native application menu missing".into());
                };
                let children = application.items().map_err(|e| e.to_string())?;
                let Some(MenuItemKind::Predefined(quit)) = children.last() else {
                    return Err("unavailable: Native Quit menu missing".into());
                };
                let text = quit.text().map_err(|e| e.to_string())?;
                application
                    .remove_at(children.len() - 1)
                    .map_err(|e| e.to_string())?;
                application
                    .append(
                        &MenuItem::with_id(app, QUIT, text, true, Some("CmdOrCtrl+Q"))
                            .map_err(|e| e.to_string())?,
                    )
                    .map_err(|e| e.to_string())?;
                menu.append(application).map_err(|e| e.to_string())?;
            }
            Some("viewMenu") => {
                let view = Submenu::new(app, "View", true).map_err(|e| e.to_string())?;
                for (id, label, key) in [
                    ("standalone.reload", "Reload view", "CmdOrCtrl+R"),
                    ("standalone.zoom.reset", "Actual size", "CmdOrCtrl+0"),
                    ("standalone.zoom.in", "Zoom in", "CmdOrCtrl+Plus"),
                    ("standalone.zoom.out", "Zoom out", "CmdOrCtrl+-"),
                    ("standalone.fullscreen", "Toggle full screen", "Ctrl+Cmd+F"),
                ] {
                    view.append(
                        &MenuItem::with_id(app, id, label, true, Some(key))
                            .map_err(|e| e.to_string())?,
                    )
                    .map_err(|e| e.to_string())?;
                }
                menu.append(&view).map_err(|e| e.to_string())?;
            }
            Some("windowMenu") => {
                let window = Submenu::new(app, "Window", true).map_err(|e| e.to_string())?;
                window
                    .append(&PredefinedMenuItem::minimize(app, None).map_err(|e| e.to_string())?)
                    .map_err(|e| e.to_string())?;
                window
                    .append(
                        &PredefinedMenuItem::maximize(app, Some("Zoom"))
                            .map_err(|e| e.to_string())?,
                    )
                    .map_err(|e| e.to_string())?;
                window
                    .append(
                        &MenuItem::with_id(
                            app,
                            "standalone.front",
                            "Bring editor to front",
                            true,
                            None::<&str>,
                        )
                        .map_err(|e| e.to_string())?,
                    )
                    .map_err(|e| e.to_string())?;
                menu.append(&window).map_err(|e| e.to_string())?;
            }
            None => {
                let label = entry["label"]
                    .as_str()
                    .filter(|text| text.len() <= 128)
                    .ok_or("validation: Invalid standalone menu label")?;
                let submenu = Submenu::new(app, label, true).map_err(|e| e.to_string())?;
                for item in entry["submenu"]
                    .as_array()
                    .filter(|items| items.len() <= 32)
                    .ok_or("validation: Invalid standalone submenu")?
                {
                    if item["type"] == "separator" {
                        submenu
                            .append(&PredefinedMenuItem::separator(app).map_err(|e| e.to_string())?)
                            .map_err(|e| e.to_string())?;
                        continue;
                    }
                    let role = match item["role"].as_str() {
                        Some("cut") => Some(PredefinedMenuItem::cut(app, None)),
                        Some("copy") => Some(PredefinedMenuItem::copy(app, None)),
                        Some("paste") => Some(PredefinedMenuItem::paste(app, None)),
                        Some("selectAll") => Some(PredefinedMenuItem::select_all(app, None)),
                        Some(_) => {
                            return Err("validation: Unsupported standalone menu role".into());
                        }
                        None => None,
                    };
                    if let Some(item) = role {
                        submenu
                            .append(&item.map_err(|e| e.to_string())?)
                            .map_err(|e| e.to_string())?;
                        continue;
                    }
                    let id = item["id"]
                        .as_str()
                        .filter(|id| id.starts_with("xlsx-menu.") && id.len() <= 80)
                        .ok_or("validation: Invalid standalone menu identity")?;
                    let label = item["label"]
                        .as_str()
                        .filter(|text| text.len() <= 128)
                        .ok_or("validation: Invalid standalone menu item")?;
                    submenu
                        .append(
                            &MenuItem::with_id(app, id, label, true, item["accelerator"].as_str())
                                .map_err(|e| e.to_string())?,
                        )
                        .map_err(|e| e.to_string())?;
                }
                menu.append(&submenu).map_err(|e| e.to_string())?;
            }
            Some(_) => return Err("validation: Unsupported standalone menu".into()),
        }
    }
    app.set_menu(menu).map_err(|e| e.to_string())?;
    Ok(())
}
pub fn event(app: &AppHandle, id: &str) {
    if id == QUIT {
        app.exit(0);
        return;
    }
    if id.starts_with("xlsx-menu.") {
        let backend = app
            .state::<std::sync::Arc<crate::backend::Backend>>()
            .inner()
            .clone();
        let id = id.to_owned();
        tauri::async_runtime::spawn(async move {
            if let Err(error) = backend.invoke("standalone:menu", vec![id.into()]).await {
                eprintln!("{error}")
            }
        });
        return;
    }
    let Some(window) = app.get_webview_window("main") else {
        return;
    };
    let result = match id {
        "standalone.reload" => window.reload(),
        "standalone.front" => window.show().and_then(|_| window.set_focus()),
        "standalone.fullscreen" => window
            .is_fullscreen()
            .and_then(|value| window.set_fullscreen(!value)),
        "standalone.zoom.reset" | "standalone.zoom.in" | "standalone.zoom.out" => {
            let state = app.state::<crate::QuitState>();
            let mut zoom = state.zoom.lock().unwrap();
            *zoom = if id.ends_with("reset") {
                1.0
            } else {
                (*zoom * if id.ends_with("in") { 1.2 } else { 1.0 / 1.2 }).clamp(0.5, 3.0)
            };
            window.set_zoom(*zoom)
        }
        _ => return,
    };
    if let Err(error) = result {
        eprintln!("{error}")
    }
}
