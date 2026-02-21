package com.zys.excel2image

import android.content.ContentValues
import android.content.Context
import android.net.Uri
import android.os.Build
import android.os.Environment
import android.provider.MediaStore
import java.io.File
import java.io.OutputStream

object DocumentSaver {
    fun savePdfToDownloads(context: Context, displayName: String, write: (OutputStream) -> Unit): Uri {
        return if (Build.VERSION.SDK_INT >= 29) {
            saveViaMediaStore29Plus(context, displayName, write)
        } else {
            saveViaMediaStoreLegacy(context, displayName, write)
        }
    }

    private fun saveViaMediaStore29Plus(
        context: Context,
        displayName: String,
        write: (OutputStream) -> Unit,
    ): Uri {
        val resolver = context.contentResolver
        val values = ContentValues().apply {
            put(MediaStore.MediaColumns.DISPLAY_NAME, displayName)
            put(MediaStore.MediaColumns.MIME_TYPE, "application/pdf")
            put(
                MediaStore.MediaColumns.RELATIVE_PATH,
                Environment.DIRECTORY_DOWNLOADS + File.separator + "Excel2Image",
            )
            put(MediaStore.MediaColumns.IS_PENDING, 1)
        }

        val uri = resolver.insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, values)
            ?: error("MediaStore insert failed")

        resolver.openOutputStream(uri).useOrThrow { out ->
            write(out)
        }

        values.clear()
        values.put(MediaStore.MediaColumns.IS_PENDING, 0)
        resolver.update(uri, values, null, null)

        return uri
    }

    @Suppress("DEPRECATION")
    private fun saveViaMediaStoreLegacy(
        context: Context,
        displayName: String,
        write: (OutputStream) -> Unit,
    ): Uri {
        val downloads = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS)
        val dir = File(downloads, "Excel2Image").apply { mkdirs() }
        val targetFile = File(dir, displayName)

        val resolver = context.contentResolver
        val values = ContentValues().apply {
            put(MediaStore.Files.FileColumns.DISPLAY_NAME, displayName)
            put(MediaStore.Files.FileColumns.MIME_TYPE, "application/pdf")
            // Deprecated but still the most reliable way to control the output folder pre-Android 10.
            put(MediaStore.Files.FileColumns.DATA, targetFile.absolutePath)
        }

        val uri = resolver.insert(MediaStore.Files.getContentUri("external"), values)
            ?: error("MediaStore insert failed")

        resolver.openOutputStream(uri).useOrThrow { out ->
            write(out)
        }

        return uri
    }
}

