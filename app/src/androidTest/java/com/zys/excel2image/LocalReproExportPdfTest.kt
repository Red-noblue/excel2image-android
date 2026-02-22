package com.zys.excel2image

import android.net.Uri
import android.util.Log
import androidx.test.ext.junit.runners.AndroidJUnit4
import androidx.test.platform.app.InstrumentationRegistry
import org.junit.Assert.assertTrue
import org.junit.Assume.assumeTrue
import org.junit.Test
import org.junit.runner.RunWith
import java.io.File
import java.io.FileOutputStream

/**
 * Local-only repro helper.
 *
 * Usage:
 * 1) Push an xlsx into target external files dir as "repro.xlsx"
 * 2) Run: ./gradlew :app:connectedDebugAndroidTest
 * 3) Pull the generated "repro-out.pdf" from the same directory.
 *
 * The xlsx itself is NOT committed (it can contain sensitive data).
 */
@RunWith(AndroidJUnit4::class)
class LocalReproExportPdfTest {
    @Test
    fun exportPdf_fromExternalFilesDir() {
        val ctx = InstrumentationRegistry.getInstrumentation().targetContext
        val dir = ctx.getExternalFilesDir(null) ?: error("External files dir is null")

        val xlsx = File(dir, "repro.xlsx")
        assumeTrue("Missing repro file: ${xlsx.absolutePath}", xlsx.exists())

        val outPdf = File(dir, "repro-out.pdf")
        if (outPdf.exists()) outPdf.delete()

        val loaded = ExcelLoader.load(ctx.contentResolver, Uri.fromFile(xlsx))
        try {
            val options = RenderOptions(
                scale = 1.4f,
                maxBitmapDimension = 16_000,
                maxTotalPixels = Long.MAX_VALUE,
                uniformFontPerColumn = true,
                trimMaxCells = 250_000,
                columnWidthMaxCells = 250_000,
                columnFontMaxCells = 250_000,
                autoFitMaxCells = 250_000,
                maxAutoRowHeightPx = 2000,
                minFontPt = 8,
                maxFontPt = 20,
            )

            FileOutputStream(outPdf).use { os ->
                val res = ExcelBitmapRenderer.writeSheetPdf(
                    workbook = loaded.workbook,
                    sheetIndex = 0,
                    options = options,
                    out = os,
                )
                Log.i(
                    "LocalRepro",
                    "PDF exported: pages=${res.pageCount}, split=${res.wasSplit}, warnings=${res.warnings.joinToString("; ")}",
                )
            }
        } finally {
            runCatching { loaded.workbook.close() }
        }

        assertTrue("PDF not generated: ${outPdf.absolutePath}", outPdf.exists() && outPdf.length() > 0L)
    }
}
