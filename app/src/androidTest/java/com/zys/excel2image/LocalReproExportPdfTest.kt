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
 * 2) Run the instrumentation test (e.g. via scripts/local_repro_export_pdf.sh)
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
            val args = InstrumentationRegistry.getArguments()
            val sheetIndexArg = args.getString("sheetIndex")?.trim().orEmpty()
            val sheetNameArg = args.getString("sheetName")?.trim().orEmpty()

            val sheetIndex =
                when {
                    sheetNameArg.isNotEmpty() -> {
                        var idx = -1
                        for (i in 0 until loaded.workbook.numberOfSheets) {
                            if (loaded.workbook.getSheetAt(i).sheetName == sheetNameArg) {
                                idx = i
                                break
                            }
                        }
                        if (idx >= 0) idx else 0
                    }

                    sheetIndexArg.isNotEmpty() -> sheetIndexArg.toIntOrNull()?.coerceAtLeast(0) ?: 0
                    else -> 0
                }

            Log.i(
                "LocalRepro",
                "Using sheetIndex=$sheetIndex sheetName='${runCatching { loaded.workbook.getSheetAt(sheetIndex).sheetName }.getOrNull()}'",
            )

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

            val dbg = ExcelBitmapRenderer.debugComputeColumnWidths(
                workbook = loaded.workbook,
                sheetIndex = sheetIndex,
                options = options,
            )
            Log.i(
                "LocalRepro",
                "Debug colWidths: sheet='${dbg.sheetName}', cols=${dbg.colWidthsPx.size}, scale=${dbg.scale}, widthPx=${dbg.scaledWidthPx}",
            )
            dbg.colWidthsPx.forEachIndexed { i, w ->
                val header = dbg.headers.getOrNull(i).orEmpty()
                Log.i("LocalRepro", "col=${i + 1} header='$header' widthPxBase=$w")
            }

            val wrapIssues = ExcelBitmapRenderer.debugFindWrapIssues(
                workbook = loaded.workbook,
                sheetIndex = sheetIndex,
                options = options,
                minLineCount = 4,
                maxIssues = 50,
            )
            if (wrapIssues.isEmpty()) {
                Log.i("LocalRepro", "Wrap issues: none (>=4 lines).")
            } else {
                Log.w("LocalRepro", "Wrap issues (>=4 lines): count=${wrapIssues.size}")
                wrapIssues.forEach { iss ->
                    Log.w(
                        "LocalRepro",
                        "wrap row=${iss.row} col=${iss.col} header='${iss.header}' lines=${iss.lineCount} text='${iss.textPreview}'",
                    )
                }
            }

            FileOutputStream(outPdf).use { os ->
                val res = ExcelBitmapRenderer.writeSheetPdf(
                    workbook = loaded.workbook,
                    sheetIndex = sheetIndex,
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
