package ru.docxvim.app

import android.os.Bundle
import android.app.Activity
import android.content.Intent
import androidx.activity.enableEdgeToEdge

class MainActivity : TauriActivity() {
  override fun onCreate(savedInstanceState: Bundle?) {
    enableEdgeToEdge()
    super.onCreate(savedInstanceState)
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
