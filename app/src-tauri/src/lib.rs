mod document_session;

#[cfg(target_os = "android")]
fn android_task_method(window: tauri::WebviewWindow, method: &'static str, id: String, name: String) -> Result<(), String> {
    use jni::objects::{JObject, JValue};
    window.with_webview(move |webview| {
        webview.jni_handle().exec(move |env, activity, _| {
            if let (Ok(id), Ok(name)) = (env.new_string(id), env.new_string(name)) {
                let id = JObject::from(id);
                let name = JObject::from(name);
                let _ = env.call_method(activity, method, "(Ljava/lang/String;Ljava/lang/String;)V", &[JValue::Object(&id), JValue::Object(&name)]);
            }
        });
    }).map_err(|error| error.to_string())
}

#[tauri::command]
fn register_task(window: tauri::WebviewWindow, id: String, name: String) -> Result<(), String> {
    #[cfg(target_os = "android")]
    { android_task_method(window, "registerDocxTask", id, name) }
    #[cfg(not(target_os = "android"))]
    { let _ = (window, id, name); Ok(()) }
}

#[tauri::command]
fn focus_task(window: tauri::WebviewWindow, id: String) -> Result<(), String> {
    #[cfg(target_os = "android")]
    { android_task_method(window, "focusDocxTask", id, String::new()) }
    #[cfg(not(target_os = "android"))]
    { let _ = (window, id); Ok(()) }
}

#[tauri::command]
fn close_task(window: tauri::WebviewWindow) -> Result<(), String> {
    #[cfg(target_os = "android")]
    {
        let id = window.label().strip_prefix("doc-").ok_or("Not a document task")?.to_string();
        android_task_method(window, "closeDocxTask", id, String::new())
    }
    #[cfg(not(target_os = "android"))]
    { let _ = window; Ok(()) }
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .manage(document_session::DocumentSessions::default())
        .invoke_handler(tauri::generate_handler![register_task, focus_task, close_task,
            document_session::get_document_session, document_session::join_document_session,
            document_session::leave_document_session,
            document_session::publish_document_session, document_session::acquire_document_session,
            document_session::release_document_session, document_session::update_document_source_fingerprint,
            document_session::reset_document_session, document_session::confirm_document_handoff])
        .plugin(tauri_plugin_clipboard_manager::init())
        .plugin(tauri_plugin_http::init())
        .plugin(tauri_plugin_keyring_store::init())
        .plugin(tauri_plugin_fs::init())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
