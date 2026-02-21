package com.zys.excel2image

import android.content.ContentValues
import android.content.Context
import android.graphics.Bitmap
import android.net.Uri
import android.os.Build
import android.os.Environment
import android.provider.MediaStore
import java.io.File

object ImageSaver {
    fun savePngToPictures(context: Context, displayName: String, bitmap: Bitmap): Uri {
        return if (Build.VERSION.SDK_INT >= 29) saveViaMediaStore29Plus(context, displayName, bitmap)
        else saveViaMediaStoreLegacy(context, displayName, bitmap)
    }

    private fun saveViaMediaStore29Plus(context: Context, displayName: String, bitmap: Bitmap): Uri {
        val resolver = context.contentResolver
        val values = ContentValues().apply {
            put(MediaStore.Images.Media.DISPLAY_NAME, displayName)
            put(MediaStore.Images.Media.MIME_TYPE, "image/png")
            put(
                MediaStore.Images.Media.RELATIVE_PATH,
                Environment.DIRECTORY_PICTURES + File.separator + "Excel2Image",
            )
            put(MediaStore.Images.Media.IS_PENDING, 1)
        }

        val uri = resolver.insert(MediaStore.Images.Media.EXTERNAL_CONTENT_URI, values)
            ?: error("MediaStore insert failed")

        resolver.openOutputStream(uri).useOrThrow { out ->
            if (!bitmap.compress(Bitmap.CompressFormat.PNG, 100, out)) {
                error("Bitmap compress failed")
            }
        }

        values.clear()
        values.put(MediaStore.Images.Media.IS_PENDING, 0)
        resolver.update(uri, values, null, null)

        return uri
    }

    @Suppress("DEPRECATION")
    private fun saveViaMediaStoreLegacy(context: Context, displayName: String, bitmap: Bitmap): Uri {
        val pictures = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_PICTURES)
        val dir = File(pictures, "Excel2Image").apply { mkdirs() }
        val targetFile = File(dir, displayName)

        val resolver = context.contentResolver
        val values = ContentValues().apply {
            put(MediaStore.Images.Media.DISPLAY_NAME, displayName)
            put(MediaStore.Images.Media.MIME_TYPE, "image/png")
            // Deprecated but still the most reliable way to control the output folder pre-Android 10.
            put(MediaStore.Images.Media.DATA, targetFile.absolutePath)
        }

        val uri = resolver.insert(MediaStore.Images.Media.EXTERNAL_CONTENT_URI, values)
            ?: error("MediaStore insert failed")

        resolver.openOutputStream(uri).useOrThrow { out ->
            if (!bitmap.compress(Bitmap.CompressFormat.PNG, 100, out)) {
                error("Bitmap compress failed")
            }
        }

        return uri
    }
}
