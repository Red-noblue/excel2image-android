package com.zys.excel2image

import android.content.Context
import android.net.Uri
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext

object ExcelExporter {
    suspend fun exportSheetAsPng(
        context: Context,
        workbook: org.apache.poi.ss.usermodel.Workbook,
        sheetIndex: Int,
        baseName: String,
        options: RenderOptions,
    ): List<Uri> = withContext(Dispatchers.Default) {
        val sheetName = workbook.getSheetAt(sheetIndex).sheetName

        val render = ExcelBitmapRenderer.renderSheet(workbook, sheetIndex, options)
        val out = ArrayList<Uri>(render.bitmaps.size)

        for ((i, bmp) in render.bitmaps.withIndex()) {
            val displayName = buildString {
                append(baseName)
                append("_")
                append(sanitizeForFileName(sheetName))
                if (render.bitmaps.size > 1) {
                    append("_p")
                    append((i + 1).toString().padStart(2, '0'))
                }
                append(".png")
            }
            val uri = ImageSaver.savePngToPictures(context, displayName, bmp)
            out += uri
            bmp.recycle()
        }

        out
    }

    private fun sanitizeForFileName(name: String): String {
        return name.replace(Regex("""[\\/:*?"<>|]"""), "_").trim().ifBlank { "sheet" }
    }
}

