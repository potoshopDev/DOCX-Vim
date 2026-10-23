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

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![register_task, focus_task])
        .plugin(tauri_plugin_fs::init())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
