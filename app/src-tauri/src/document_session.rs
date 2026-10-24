use serde::Serialize;
use serde_json::Value;
use std::{
    collections::{HashMap, HashSet},
    sync::Mutex,
    time::Duration,
};
use tauri::{AppHandle, Emitter, Manager, State};
use tokio::sync::oneshot;

#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SessionSnapshot {
    revision: u64,
    editor_label: String,
    draft: Value,
    unconfirmed_from: Option<String>,
}

struct Session {
    snapshot: SessionSnapshot,
    participants: HashSet<String>,
    pending: Option<(String, oneshot::Sender<SessionSnapshot>)>,
}

#[derive(Default)]
pub struct DocumentSessions(Mutex<HashMap<String, Session>>);

#[tauri::command]
pub fn get_document_session(
    id: String,
    app: AppHandle,
    sessions: State<'_, DocumentSessions>,
) -> Result<Option<SessionSnapshot>, String> {
    let mut sessions = sessions.0.lock().map_err(|error| error.to_string())?;
    let Some(session) = sessions.get_mut(&id) else {
        return Ok(None);
    };
    let mut confirmed = false;
    if let Some(previous) = &session.snapshot.unconfirmed_from {
        if app.get_webview_window(previous).is_none() {
            session.snapshot.unconfirmed_from = None;
            session.snapshot.revision += 1;
            confirmed = true;
        }
    }
    let snapshot = session.snapshot.clone();
    drop(sessions);
    if confirmed {
        let _ = app.emit("docx-vim-handoff-confirmed", &snapshot);
    }
    Ok(Some(snapshot))
}

#[tauri::command]
pub fn join_document_session(
    id: String,
    draft: Value,
    window: tauri::WebviewWindow,
    sessions: State<'_, DocumentSessions>,
) -> Result<SessionSnapshot, String> {
    let mut sessions = sessions.0.lock().map_err(|error| error.to_string())?;
    let session = sessions.entry(id).or_insert_with(|| Session {
        snapshot: SessionSnapshot {
            revision: 0,
            editor_label: window.label().to_string(),
            draft,
            unconfirmed_from: None,
        },
        participants: HashSet::new(),
        pending: None,
    });
    session.participants.insert(window.label().to_string());
    Ok(session.snapshot.clone())
}

#[tauri::command]
pub fn leave_document_session(
    id: String,
    window: tauri::WebviewWindow,
    app: AppHandle,
    sessions: State<'_, DocumentSessions>,
) -> Result<(), String> {
    let mut sessions = sessions.0.lock().map_err(|error| error.to_string())?;
    let Some(session) = sessions.get_mut(&id) else {
        return Ok(());
    };
    session.participants.remove(window.label());
    if session.participants.is_empty() {
        sessions.remove(&id);
        return Ok(());
    }
    if session.snapshot.editor_label == window.label() {
        session.snapshot.editor_label.clear();
    }
    if session
        .pending
        .as_ref()
        .is_some_and(|(target, _)| target == window.label())
    {
        session.pending = None;
    }
    if session.snapshot.unconfirmed_from.as_deref() == Some(window.label()) {
        session.snapshot.unconfirmed_from = None;
        session.snapshot.revision += 1;
        let snapshot = session.snapshot.clone();
        drop(sessions);
        let _ = app.emit("docx-vim-handoff-confirmed", &snapshot);
    }
    Ok(())
}

#[tauri::command]
pub fn publish_document_session(
    id: String,
    draft: Value,
    window: tauri::WebviewWindow,
    app: AppHandle,
    sessions: State<'_, DocumentSessions>,
) -> Result<SessionSnapshot, String> {
    let snapshot = {
        let mut sessions = sessions.0.lock().map_err(|error| error.to_string())?;
        let session = sessions.get_mut(&id).ok_or("Document session missing")?;
        if session.snapshot.editor_label != window.label() {
            return Err("Edit lease belongs to another window".into());
        }
        let mut draft = draft;
        // The save owner may have updated the source fingerprint after this editor's last snapshot.
        if draft.get("sourcePath") == session.snapshot.draft.get("sourcePath")
            && draft.get("sourceFingerprint").and_then(Value::as_str) != Some("new")
        {
            if let Some(revision) = session.snapshot.draft.get("sourceFingerprint").cloned() {
                if let Some(object) = draft.as_object_mut() {
                    object.insert("sourceFingerprint".into(), revision);
                }
            }
        }
        session.snapshot.revision += 1;
        session.snapshot.draft = draft;
        session.snapshot.clone()
    };
    let _ = app.emit("docx-vim-session", &snapshot);
    Ok(snapshot)
}

#[tauri::command]
pub async fn acquire_document_session(
    id: String,
    window: tauri::WebviewWindow,
    app: AppHandle,
    sessions: State<'_, DocumentSessions>,
) -> Result<SessionSnapshot, String> {
    let label = window.label().to_string();
    let (receiver, previous) = {
        let mut sessions = sessions.0.lock().map_err(|error| error.to_string())?;
        let session = sessions.get_mut(&id).ok_or("Document session missing")?;
        if session.snapshot.editor_label == label {
            return Ok(session.snapshot.clone());
        }
        if app
            .get_webview_window(&session.snapshot.editor_label)
            .is_none()
        {
            session.snapshot.editor_label = label.clone();
            session.pending = None;
            return Ok(session.snapshot.clone());
        }
        if session.pending.is_some() {
            return Err("Edit transfer already in progress".into());
        }
        let (sender, receiver) = oneshot::channel();
        session.pending = Some((label.clone(), sender));
        (receiver, session.snapshot.editor_label.clone())
    };
    let _ = app.emit(
        "docx-vim-yield",
        serde_json::json!({"id": id, "editorLabel": previous}),
    );
    let wait = if cfg!(target_os = "android") {
        Duration::from_secs(3)
    } else {
        Duration::from_secs(10)
    };
    match tokio::time::timeout(wait, receiver).await {
        Ok(Ok(snapshot)) => Ok(snapshot),
        _ => {
            let mut sessions = sessions.0.lock().map_err(|error| error.to_string())?;
            if let Some(session) = sessions.get_mut(&id) {
                if session
                    .pending
                    .as_ref()
                    .is_some_and(|(target, _)| target == window.label())
                {
                    session.pending = None;
                    if cfg!(target_os = "android") {
                        session.snapshot.unconfirmed_from =
                            Some(session.snapshot.editor_label.clone());
                        session.snapshot.editor_label = label;
                        let snapshot = session.snapshot.clone();
                        drop(sessions);
                        let _ = app.emit("docx-vim-lease", &snapshot);
                        return Ok(snapshot);
                    }
                }
            }
            Err("Edit transfer timed out".into())
        }
    }
}

#[tauri::command]
pub fn release_document_session(
    id: String,
    window: tauri::WebviewWindow,
    app: AppHandle,
    sessions: State<'_, DocumentSessions>,
) -> Result<(), String> {
    let mut sessions = sessions.0.lock().map_err(|error| error.to_string())?;
    let session = sessions.get_mut(&id).ok_or("Document session missing")?;
    if session.snapshot.editor_label != window.label() {
        return Ok(());
    }
    if let Some((target, sender)) = session.pending.take() {
        session.snapshot.editor_label = target;
        let snapshot = session.snapshot.clone();
        let _ = sender.send(snapshot.clone());
        let _ = app.emit("docx-vim-lease", snapshot);
    }
    Ok(())
}

#[tauri::command]
pub fn confirm_document_handoff(
    id: String,
    window: tauri::WebviewWindow,
    app: AppHandle,
    sessions: State<'_, DocumentSessions>,
) -> Result<(), String> {
    let mut sessions = sessions.0.lock().map_err(|error| error.to_string())?;
    let session = sessions.get_mut(&id).ok_or("Document session missing")?;
    if session.snapshot.unconfirmed_from.as_deref() == Some(window.label()) {
        session.snapshot.unconfirmed_from = None;
        session.snapshot.revision += 1;
        let snapshot = session.snapshot.clone();
        drop(sessions);
        let _ = app.emit("docx-vim-handoff-confirmed", &snapshot);
    }
    Ok(())
}

#[tauri::command]
pub fn update_document_source_fingerprint(
    id: String,
    fingerprint: String,
    app: AppHandle,
    sessions: State<'_, DocumentSessions>,
) -> Result<(), String> {
    let snapshot = {
        let mut sessions = sessions.0.lock().map_err(|error| error.to_string())?;
        let session = sessions.get_mut(&id).ok_or("Document session missing")?;
        if let Some(draft) = session.snapshot.draft.as_object_mut() {
            draft.insert("sourceFingerprint".into(), Value::String(fingerprint));
        }
        session.snapshot.clone()
    };
    let _ = app.emit("docx-vim-fingerprint", snapshot);
    Ok(())
}

#[tauri::command]
pub fn reset_document_session(
    id: String,
    draft: Value,
    window: tauri::WebviewWindow,
    sessions: State<'_, DocumentSessions>,
) -> Result<(), String> {
    let mut sessions = sessions.0.lock().map_err(|error| error.to_string())?;
    let session = sessions.get_mut(&id).ok_or("Document session missing")?;
    session.pending = None;
    session.snapshot.revision += 1;
    session.snapshot.editor_label = window.label().to_string();
    session.snapshot.draft = draft;
    session.snapshot.unconfirmed_from = None;
    Ok(())
}
