package com.zys.excel2image

import android.content.Context
import android.net.Uri
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext

data class PdfExportResult(
    val uri: Uri,
    val writeResult: PdfWriteResult,
)

object ExcelPdfExporter {
    suspend fun exportSheetAsPdf(
        context: Context,
        workbook: org.apache.poi.ss.usermodel.Workbook,
        sheetIndex: Int,
        baseName: String,
        options: RenderOptions,
    ): PdfExportResult = withContext(Dispatchers.Default) {
        val sheetName = workbook.getSheetAt(sheetIndex).sheetName
        val displayName = buildString {
            append(baseName)
            append("_")
            append(sanitizeForFileName(sheetName))
            append(".pdf")
        }

        var writeResult: PdfWriteResult? = null
        val uri = DocumentSaver.savePdfToDownloads(context, displayName) { out ->
            writeResult = ExcelBitmapRenderer.writeSheetPdf(
                workbook = workbook,
                sheetIndex = sheetIndex,
                options = options,
                out = out,
            )
        }

        PdfExportResult(uri = uri, writeResult = writeResult ?: error("PDF write result missing"))
    }

    private fun sanitizeForFileName(name: String): String {
        return name.replace(Regex("""[\\/:*?"<>|]"""), "_").trim().ifBlank { "sheet" }
    }
}

