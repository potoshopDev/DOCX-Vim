package ru.docxvim.app

import android.os.Bundle
import android.app.Activity
import android.app.ActivityManager
import android.content.Intent
import androidx.activity.enableEdgeToEdge

open class MainActivity : TauriActivity() {
  override fun onCreate(savedInstanceState: Bundle?) {
    enableEdgeToEdge()
    super.onCreate(savedInstanceState)
    if (this::class.java == MainActivity::class.java) registerDocxTask("main", "DOCX Vim")
  }

  fun registerDocxTask(documentId: String, name: String) {
    getSharedPreferences("docx_tasks", MODE_PRIVATE).edit().putInt(documentId, taskId).apply()
    setTaskDescription(ActivityManager.TaskDescription(name))
  }

  fun focusDocxTask(documentId: String, _name: String) {
    val savedTaskId = getSharedPreferences("docx_tasks", MODE_PRIVATE).getInt(documentId, -1)
    if (savedTaskId < 0) return
    val manager = getSystemService(ACTIVITY_SERVICE) as ActivityManager
    manager.appTasks.firstOrNull { it.taskInfo?.taskId == savedTaskId }?.moveToFront()
  }

  @Deprecated("Required to retain access to documents selected by the system picker")
  override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
    if (resultCode == Activity.RESULT_OK && data?.data != null) {
      val flags = data.flags and (Intent.FLAG_GRANT_READ_URI_PERMISSION or Intent.FLAG_GRANT_WRITE_URI_PERMISSION)
      if (flags != 0) {
        try {
          contentResolver.takePersistableUriPermission(data.data!!, flags)
        } catch (_: SecurityException) {
          // ACTION_GET_CONTENT can grant access only for the current session.
        }
      }
    }
    super.onActivityResult(requestCode, resultCode, data)
  }
}
