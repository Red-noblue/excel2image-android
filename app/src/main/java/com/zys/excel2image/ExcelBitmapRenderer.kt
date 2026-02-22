package com.zys.excel2image

import android.graphics.Bitmap
import android.graphics.Canvas
import android.graphics.Color
import android.graphics.Paint
import android.graphics.RectF
import android.graphics.Typeface
import android.graphics.pdf.PdfDocument
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellRangeAddress
import java.io.OutputStream
import kotlin.math.ceil
import kotlin.math.max
import kotlin.math.min
import kotlin.math.roundToInt

data class RenderOptions(
    val scale: Float,
    val maxBitmapDimension: Int = 16_000,
    val maxTotalPixels: Long = 100_000_000L,
    // "Share clear (tight)" helpers.
    val trimBlankEdges: Boolean = true,
    val autoFitRowHeights: Boolean = true,
    val adaptiveColumnWidths: Boolean = true,
    val autoWrapOverflowText: Boolean = true,
    // Typography: keep each column's font size consistent to avoid "some cells suddenly tiny/huge".
    val uniformFontPerColumn: Boolean = false,
    // Safety limits to avoid heavy scanning on huge sheets.
    val trimMaxCells: Int = 120_000,
    val columnWidthMaxCells: Int = 120_000,
    val columnFontMaxCells: Int = 120_000,
    val autoFitMaxCells: Int = 120_000,
    // For column width estimation we skip extremely long texts to avoid one outlier making the column huge.
    // Keep this fairly high because many Chinese "project/task names" are 30-60 chars and should still be sampled.
    val columnWidthSampleMaxTextLength: Int = 80,
    val minColumnWidthPx: Int = 12,
    // "Empty" means: no visible text in this column within the used range.
    // We still keep a small width (instead of 0) so grid lines remain readable.
    val minEmptyColumnWidthPx: Int = 8,
    val emptyColumnWidthPx: Int = 10,
    val maxColumnWidthPx: Int = 1200,
    val autoWrapMinTextLength: Int = 12,
    val autoWrapExcludeNumeric: Boolean = true,
    // Cap auto row height growth in *base* pixels (before applying `scale`).
    val maxAutoRowHeightPx: Int = 1200,
    val minFontPt: Int = 8,
    val maxFontPt: Int = 28,
)

data class RenderResult(
    val bitmaps: List<Bitmap>,
    val wasSplit: Boolean,
    val warnings: List<String>,
)

data class PdfWriteResult(
    val pageCount: Int,
    val wasSplit: Boolean,
    val warnings: List<String>,
)

// Debug helper: surfaced for instrumentation/local repro scripts.
// Not used by the production UI.
data class DebugColumnWidths(
    val sheetName: String,
    val firstCol: Int,
    val lastCol: Int,
    val colWidthsPx: IntArray,
    val scaledWidthPx: Int,
    val scale: Float,
    val headers: List<String>,
)

data class DebugWrapIssue(
    // 1-based indices for easier cross-checking in Excel.
    val row: Int,
    val col: Int,
    val header: String,
    val lineCount: Int,
    val textPreview: String,
)

data class RenderPartsResult(
    val partCount: Int,
    val wasSplit: Boolean,
    val warnings: List<String>,
)

private data class MergeRegion(
    val firstRow: Int,
    val lastRow: Int,
    val firstCol: Int,
    val lastCol: Int,
)

object ExcelBitmapRenderer {
    private fun cleanCellText(raw: String): String {
        if (raw.isEmpty()) return ""
        // Excel exports sometimes include non-breaking spaces / full-width spaces that look empty but
        // should be treated as whitespace for trimming and "is empty" checks.
        return raw.replace('\u00A0', ' ').replace('\u3000', ' ').trim()
    }

    // Use ceil when scaling pixel sizes to avoid rounding down causing:
    // - unexpected extra wraps (scaled width slightly smaller)
    // - clipped multiline text (scaled height slightly smaller)
    private fun scaledPx(basePx: Int, scale: Float): Int {
        if (basePx <= 0) return 0
        return max(1, ceil(basePx.toDouble() * scale.toDouble()).toInt())
    }

    private fun scaledSumPx(basePx: IntArray, scale: Float): Int {
        var sum = 0
        for (v in basePx) sum += scaledPx(v, scale)
        return sum
    }

    private data class FitScaleResult(
        val scale: Float,
        val widthPx: Int,
    )

    private fun fitScaleToMaxWidth(
        baseColWidthsPx: IntArray,
        requestedScale: Float,
        maxWidthPx: Int,
    ): FitScaleResult {
        var scale = requestedScale.coerceAtLeast(0.1f)
        var width = scaledSumPx(baseColWidthsPx, scale)
        if (maxWidthPx <= 0 || width <= maxWidthPx) return FitScaleResult(scale = scale, widthPx = width)

        // Iteratively shrink: ceil rounding can keep us slightly above the max after a single pass.
        repeat(6) {
            val shrink = maxWidthPx.toFloat() / width.toFloat()
            if (shrink >= 1f) return@repeat
            val nextScale = (scale * shrink).coerceAtLeast(0.01f)
            if (nextScale >= scale) return@repeat

            val nextWidth = scaledSumPx(baseColWidthsPx, nextScale)
            if (nextWidth >= width) return@repeat // no progress (likely hit min 1px per column)

            scale = nextScale
            width = nextWidth
            if (width <= maxWidthPx) return FitScaleResult(scale = scale, widthPx = width)
        }

        return FitScaleResult(scale = scale, widthPx = width)
    }

    fun renderSheet(workbook: Workbook, sheetIndex: Int, options: RenderOptions): RenderResult {
        val sheet = workbook.getSheetAt(sheetIndex)

        val candidateRange = findPrintAreaRange(workbook, sheetIndex) ?: findUsedRange(sheet)
        if (candidateRange == null) {
            val bmp = Bitmap.createBitmap(800, 400, Bitmap.Config.ARGB_8888)
            val canvas = Canvas(bmp)
            canvas.drawColor(Color.WHITE)
            val p = Paint(Paint.ANTI_ALIAS_FLAG).apply {
                color = Color.DKGRAY
                textSize = 40f
            }
            canvas.drawText("空表", 40f, 120f, p)
            return RenderResult(bitmaps = listOf(bmp), wasSplit = false, warnings = emptyList())
        }

        val warnings = mutableListOf<String>()
        val formatter = DataFormatter()
        val evaluator = workbook.creationHelper.createFormulaEvaluator()

        val used = if (options.trimBlankEdges) {
            val trim = trimUsedRange(
                sheet = sheet,
                range = candidateRange,
                formatter = formatter,
                evaluator = evaluator,
                maxCells = options.trimMaxCells,
            )
            trim.warning?.let { warnings += it }
            if (trim.didTrim) warnings += "已自动裁剪空白边缘"
            trim.range
        } else {
            candidateRange
        }

        val (firstRow, lastRow, firstCol, lastCol) = used

        val mergeInfo = buildMergeInfo(sheet, firstRow, lastRow, firstCol, lastCol)

        val columnFontPts = if (options.uniformFontPerColumn) {
            val fit = computeColumnFontPts(
                workbook = workbook,
                sheet = sheet,
                formatter = formatter,
                evaluator = evaluator,
                firstRow = firstRow,
                lastRow = lastRow,
                firstCol = firstCol,
                lastCol = lastCol,
                mergeInfo = mergeInfo,
                maxCells = options.columnFontMaxCells,
                minFontPt = options.minFontPt,
                maxFontPt = options.maxFontPt,
            )
            fit.warning?.let { warnings += it }
            warnings += "已按列统一字号"
            fit.fontPts
        } else {
            null
        }

        val maxDigitWidthPx = computeMaxDigitWidthPx(workbook)
        val baseColWidthsPxRaw = IntArray(lastCol - firstCol + 1) { idx ->
            val col = firstCol + idx
            if (sheet.isColumnHidden(col)) return@IntArray 0
            // Excel column width is in 1/256 character units. This is an approximation that works
            // well enough for "shareable images" without pulling AWT font metrics.
            val widthChars = sheet.getColumnWidth(col) / 256f
            max(12, (widthChars * maxDigitWidthPx + 5f).roundToInt())
        }

        val baseColWidthsPx = if (options.adaptiveColumnWidths) {
            val fit = adaptColumnWidthsPx(
                workbook = workbook,
                sheet = sheet,
                formatter = formatter,
                evaluator = evaluator,
                requestedScale = options.scale,
                maxScaledWidthPx = options.maxBitmapDimension,
                firstRow = firstRow,
                lastRow = lastRow,
                firstCol = firstCol,
                lastCol = lastCol,
                baseColWidthsPx = baseColWidthsPxRaw,
                mergeInfo = mergeInfo,
                columnFontPts = columnFontPts,
                maxCells = options.columnWidthMaxCells,
                sampleMaxTextLength = options.columnWidthSampleMaxTextLength,
                minColumnWidthPx = options.minColumnWidthPx,
                minEmptyColumnWidthPx = options.minEmptyColumnWidthPx,
                emptyColumnWidthPx = options.emptyColumnWidthPx,
                maxColumnWidthPx = options.maxColumnWidthPx,
                minFontPt = options.minFontPt,
                maxFontPt = options.maxFontPt,
            )
            fit.warning?.let { warnings += it }
            if (fit.didChange) warnings += "已自动调整列宽"
            fit.colWidthsPx
        } else {
            baseColWidthsPxRaw
        }

        // Determine final render scale early so auto-fit row heights can use the exact same
        // padding/text metrics logic as drawing (important when scale is small and padding has a floor).
        val fitScale =
            fitScaleToMaxWidth(
                baseColWidthsPx = baseColWidthsPx,
                requestedScale = options.scale,
                maxWidthPx = options.maxBitmapDimension,
            )
        var scale = fitScale.scale
        var width = fitScale.widthPx

        if (width > options.maxBitmapDimension) {
            warnings += "表格太宽，已强制缩小"
        }

        val baseRowHeightsPxRaw = IntArray(lastRow - firstRow + 1) { idx ->
            val rowNum = firstRow + idx
            val row = sheet.getRow(rowNum)
            if (row?.zeroHeight == true) return@IntArray 0
            val htPt = row?.heightInPoints ?: sheet.defaultRowHeightInPoints
            max(16, ceil(htPt * 4f / 3f).toInt())
        }

        val baseRowHeightsPx = if (options.autoFitRowHeights) {
            val fit = autoFitRowHeightsPx(
                workbook = workbook,
                sheet = sheet,
                formatter = formatter,
                evaluator = evaluator,
                renderScale = scale,
                firstRow = firstRow,
                lastRow = lastRow,
                firstCol = firstCol,
                lastCol = lastCol,
                baseColWidthsPx = baseColWidthsPx,
                baseRowHeightsPx = baseRowHeightsPxRaw,
                mergeInfo = mergeInfo,
                columnFontPts = columnFontPts,
                maxCells = options.autoFitMaxCells,
                maxAutoRowHeightPx = options.maxAutoRowHeightPx,
                autoWrapOverflowText = options.autoWrapOverflowText,
                autoWrapMinTextLength = options.autoWrapMinTextLength,
                autoWrapExcludeNumeric = options.autoWrapExcludeNumeric,
                minFontPt = options.minFontPt,
                maxFontPt = options.maxFontPt,
            )
            fit.warning?.let { warnings += it }
            if (fit.didChange) warnings += "已自动调整行高以适配换行"
            fit.rowHeightsPx
        } else {
            baseRowHeightsPxRaw
        }

        val parts = planVerticalParts(
            baseRowHeightsPx = baseRowHeightsPx,
            scaledWidthPx = width,
            scale = scale,
            maxBitmapDimension = options.maxBitmapDimension,
            maxTotalPixels = options.maxTotalPixels,
        )
        if (parts.isEmpty()) {
            val bmp = Bitmap.createBitmap(800, 400, Bitmap.Config.ARGB_8888)
            val canvas = Canvas(bmp)
            canvas.drawColor(Color.WHITE)
            val p = Paint(Paint.ANTI_ALIAS_FLAG).apply {
                color = Color.DKGRAY
                textSize = 36f
            }
            canvas.drawText("无可见内容（可能都被隐藏）", 40f, 120f, p)
            return RenderResult(bitmaps = listOf(bmp), wasSplit = false, warnings = warnings)
        }

        val bitmaps = parts.mapIndexed { partIndex, part ->
            val partHeight = part.heightPx
            val bmp = Bitmap.createBitmap(width, partHeight, Bitmap.Config.ARGB_8888)
            val canvas = Canvas(bmp)
            canvas.drawColor(Color.WHITE)

            drawSheetPart(
                canvas = canvas,
                workbook = workbook,
                sheet = sheet,
                formatter = formatter,
                evaluator = evaluator,
                firstRow = firstRow,
                lastRow = lastRow,
                firstCol = firstCol,
                lastCol = lastCol,
                baseColWidthsPx = baseColWidthsPx,
                baseRowHeightsPx = baseRowHeightsPx,
                scale = scale,
                partRowStart = part.rowStart,
                partRowEnd = part.rowEnd,
                partTopOffsetPx = part.topOffsetPx,
                mergeInfo = mergeInfo,
                columnFontPts = columnFontPts,
                autoWrapOverflowText = options.autoWrapOverflowText,
                autoWrapMinTextLength = options.autoWrapMinTextLength,
                autoWrapExcludeNumeric = options.autoWrapExcludeNumeric,
                minFontPt = options.minFontPt,
                maxFontPt = options.maxFontPt,
            )

            if (parts.size > 1 && partIndex == 0) {
                warnings += "因尺寸限制自动分段"
            }

            bmp
        }

        return RenderResult(
            bitmaps = bitmaps,
            wasSplit = bitmaps.size > 1,
            warnings = warnings.distinct(),
        )
    }

    @Suppress("unused")
    fun debugComputeColumnWidths(workbook: Workbook, sheetIndex: Int, options: RenderOptions): DebugColumnWidths {
        val sheet = workbook.getSheetAt(sheetIndex)
        val formatter = DataFormatter()
        val evaluator = workbook.creationHelper.createFormulaEvaluator()

        val candidateRange = findPrintAreaRange(workbook, sheetIndex) ?: findUsedRange(sheet)
            ?: return DebugColumnWidths(
                sheetName = sheet.sheetName,
                firstCol = 0,
                lastCol = -1,
                colWidthsPx = IntArray(0),
                scaledWidthPx = 0,
                scale = options.scale,
                headers = emptyList(),
            )

        val used = if (options.trimBlankEdges) {
            trimUsedRange(
                sheet = sheet,
                range = candidateRange,
                formatter = formatter,
                evaluator = evaluator,
                maxCells = options.trimMaxCells,
            ).range
        } else {
            candidateRange
        }

        val (firstRow, lastRow, firstCol, lastCol) = used
        val mergeInfo = buildMergeInfo(sheet, firstRow, lastRow, firstCol, lastCol)

        val columnFontPts = if (options.uniformFontPerColumn) {
            computeColumnFontPts(
                workbook = workbook,
                sheet = sheet,
                formatter = formatter,
                evaluator = evaluator,
                firstRow = firstRow,
                lastRow = lastRow,
                firstCol = firstCol,
                lastCol = lastCol,
                mergeInfo = mergeInfo,
                maxCells = options.columnFontMaxCells,
                minFontPt = options.minFontPt,
                maxFontPt = options.maxFontPt,
            ).fontPts
        } else {
            null
        }

        val maxDigitWidthPx = computeMaxDigitWidthPx(workbook)
        val baseColWidthsPxRaw = IntArray(lastCol - firstCol + 1) { idx ->
            val col = firstCol + idx
            if (sheet.isColumnHidden(col)) return@IntArray 0
            val widthChars = sheet.getColumnWidth(col) / 256f
            max(12, (widthChars * maxDigitWidthPx + 5f).roundToInt())
        }

        val baseColWidthsPx = if (options.adaptiveColumnWidths) {
            adaptColumnWidthsPx(
                workbook = workbook,
                sheet = sheet,
                formatter = formatter,
                evaluator = evaluator,
                requestedScale = options.scale,
                maxScaledWidthPx = options.maxBitmapDimension,
                firstRow = firstRow,
                lastRow = lastRow,
                firstCol = firstCol,
                lastCol = lastCol,
                baseColWidthsPx = baseColWidthsPxRaw,
                mergeInfo = mergeInfo,
                columnFontPts = columnFontPts,
                maxCells = options.columnWidthMaxCells,
                sampleMaxTextLength = options.columnWidthSampleMaxTextLength,
                minColumnWidthPx = options.minColumnWidthPx,
                minEmptyColumnWidthPx = options.minEmptyColumnWidthPx,
                emptyColumnWidthPx = options.emptyColumnWidthPx,
                maxColumnWidthPx = options.maxColumnWidthPx,
                minFontPt = options.minFontPt,
                maxFontPt = options.maxFontPt,
            ).colWidthsPx
        } else {
            baseColWidthsPxRaw
        }

        val fitScale =
            fitScaleToMaxWidth(
                baseColWidthsPx = baseColWidthsPx,
                requestedScale = options.scale,
                maxWidthPx = options.maxBitmapDimension,
            )

        val headers = (firstCol..lastCol).map { c ->
            val cell = sheet.getRow(firstRow)?.getCell(c)
            if (cell == null) {
                ""
            } else {
                cleanCellText(runCatching { formatter.formatCellValue(cell, evaluator) }.getOrNull().orEmpty())
            }
        }

        return DebugColumnWidths(
            sheetName = sheet.sheetName,
            firstCol = firstCol,
            lastCol = lastCol,
            colWidthsPx = baseColWidthsPx,
            scaledWidthPx = fitScale.widthPx,
            scale = fitScale.scale,
            headers = headers,
        )
    }

    @Suppress("unused")
    fun debugFindWrapIssues(
        workbook: Workbook,
        sheetIndex: Int,
        options: RenderOptions,
        minLineCount: Int = 4,
        maxIssues: Int = 50,
    ): List<DebugWrapIssue> {
        val sheet = workbook.getSheetAt(sheetIndex)
        val formatter = DataFormatter()
        val evaluator = workbook.creationHelper.createFormulaEvaluator()

        val candidateRange = findPrintAreaRange(workbook, sheetIndex) ?: findUsedRange(sheet) ?: return emptyList()
        val used = if (options.trimBlankEdges) {
            trimUsedRange(
                sheet = sheet,
                range = candidateRange,
                formatter = formatter,
                evaluator = evaluator,
                maxCells = options.trimMaxCells,
            ).range
        } else {
            candidateRange
        }

        val (firstRow, lastRow, firstCol, lastCol) = used
        val mergeInfo = buildMergeInfo(sheet, firstRow, lastRow, firstCol, lastCol)

        val columnFontPts = if (options.uniformFontPerColumn) {
            computeColumnFontPts(
                workbook = workbook,
                sheet = sheet,
                formatter = formatter,
                evaluator = evaluator,
                firstRow = firstRow,
                lastRow = lastRow,
                firstCol = firstCol,
                lastCol = lastCol,
                mergeInfo = mergeInfo,
                maxCells = options.columnFontMaxCells,
                minFontPt = options.minFontPt,
                maxFontPt = options.maxFontPt,
            ).fontPts
        } else {
            null
        }

        val widths = debugComputeColumnWidths(workbook, sheetIndex, options)
        val baseColWidthsPx = widths.colWidthsPx
        val scale = widths.scale.coerceAtLeast(0.1f)
        val padding = max(2f, 4f * scale)

        val (cellToMerge, mergeStarts) = mergeInfo
        val paint = Paint(Paint.ANTI_ALIAS_FLAG)

        val headerRow = sheet.getRow(firstRow)
        val headers = (firstCol..lastCol).map { c ->
            val cell = headerRow?.getCell(c)
            if (cell == null) "" else cleanCellText(runCatching { formatter.formatCellValue(cell, evaluator) }.getOrNull().orEmpty())
        }

        val issues = ArrayList<DebugWrapIssue>()

        val rowIt = sheet.rowIterator()
        while (rowIt.hasNext() && issues.size < maxIssues) {
            val row = rowIt.next()
            val r = row.rowNum
            if (r < firstRow || r > lastRow) continue
            if (row.zeroHeight) continue

            val cellIt = row.cellIterator()
            while (cellIt.hasNext() && issues.size < maxIssues) {
                val cell = cellIt.next()
                val c = cell.columnIndex
                if (c < firstCol || c > lastCol) continue
                if (sheet.isColumnHidden(c)) continue

                val rawText0 = runCatching { formatter.formatCellValue(cell, evaluator) }.getOrNull().orEmpty()
                val rawText = cleanCellText(rawText0)
                if (rawText.isEmpty()) continue

                val key = cellKey(r, c)
                val merge = cellToMerge[key]
                if (merge != null && key !in mergeStarts) continue

                val spanFirstCol = merge?.firstCol ?: c
                val spanLastCol = merge?.lastCol ?: c

                var widthPx = 0
                for (absCol in spanFirstCol..spanLastCol) {
                    val idx = absCol - firstCol
                    if (idx !in baseColWidthsPx.indices) continue
                    widthPx += scaledPx(baseColWidthsPx[idx], scale)
                }
                val availableWidth = max(0f, widthPx.toFloat() - padding * 2)

                val style = cell.cellStyle
                val font = workbook.getFontAt(style.fontIndex)
                val colIdxForFont = spanFirstCol - firstCol
                val overridePt = columnFontPts?.getOrNull(colIdxForFont)
                applyFont(paint, font, scale, options.minFontPt, options.maxFontPt, overridePt = overridePt)

                val hasNewline = rawText.contains('\n') || rawText.contains('\r')
                val wrap =
                    style.wrapText ||
                        hasNewline ||
                        (options.autoWrapOverflowText &&
                            shouldAutoWrapText(
                                text = rawText,
                                paint = paint,
                                availableWidth = availableWidth,
                                minTextLength = options.autoWrapMinTextLength,
                                excludeNumeric = options.autoWrapExcludeNumeric,
                            ))

                val finalWrap =
                    if (!wrap && columnFontPts != null && availableWidth > 0f) {
                        measureMaxLineWidthPx(rawText, paint) > availableWidth * 1.01f
                    } else {
                        wrap
                    }
                if (!finalWrap) continue

                val lineCount = countLinesForLayout(rawText, paint, availableWidth, true)
                if (lineCount < minLineCount) continue

                val header = headers.getOrNull(c - firstCol).orEmpty()
                val preview = if (rawText.length > 60) rawText.take(60) + "…" else rawText
                issues += DebugWrapIssue(
                    row = r + 1,
                    col = c + 1,
                    header = header,
                    lineCount = lineCount,
                    textPreview = preview,
                )
            }
        }

        return issues
    }

    fun renderSheetParts(
        workbook: Workbook,
        sheetIndex: Int,
        options: RenderOptions,
        recycleAfterCallback: Boolean = true,
        maxPartsToRender: Int = Int.MAX_VALUE,
        onPart: (partIndex: Int, partCount: Int, bitmap: Bitmap) -> Unit,
    ): RenderPartsResult {
        val sheet = workbook.getSheetAt(sheetIndex)

        val candidateRange = findPrintAreaRange(workbook, sheetIndex) ?: findUsedRange(sheet)
        if (candidateRange == null) {
            val bmp = Bitmap.createBitmap(800, 400, Bitmap.Config.ARGB_8888)
            val canvas = Canvas(bmp)
            canvas.drawColor(Color.WHITE)
            val p = Paint(Paint.ANTI_ALIAS_FLAG).apply {
                color = Color.DKGRAY
                textSize = 40f
            }
            canvas.drawText("空表", 40f, 120f, p)
            try {
                onPart(0, 1, bmp)
            } finally {
                if (recycleAfterCallback) bmp.recycle()
            }
            return RenderPartsResult(partCount = 1, wasSplit = false, warnings = emptyList())
        }

        val warnings = mutableListOf<String>()
        val formatter = DataFormatter()
        val evaluator = workbook.creationHelper.createFormulaEvaluator()

        val used = if (options.trimBlankEdges) {
            val trim = trimUsedRange(
                sheet = sheet,
                range = candidateRange,
                formatter = formatter,
                evaluator = evaluator,
                maxCells = options.trimMaxCells,
            )
            trim.warning?.let { warnings += it }
            if (trim.didTrim) warnings += "已自动裁剪空白边缘"
            trim.range
        } else {
            candidateRange
        }

        val (firstRow, lastRow, firstCol, lastCol) = used

        val mergeInfo = buildMergeInfo(sheet, firstRow, lastRow, firstCol, lastCol)

        val columnFontPts = if (options.uniformFontPerColumn) {
            val fit = computeColumnFontPts(
                workbook = workbook,
                sheet = sheet,
                formatter = formatter,
                evaluator = evaluator,
                firstRow = firstRow,
                lastRow = lastRow,
                firstCol = firstCol,
                lastCol = lastCol,
                mergeInfo = mergeInfo,
                maxCells = options.columnFontMaxCells,
                minFontPt = options.minFontPt,
                maxFontPt = options.maxFontPt,
            )
            fit.warning?.let { warnings += it }
            warnings += "已按列统一字号"
            fit.fontPts
        } else {
            null
        }

        val maxDigitWidthPx = computeMaxDigitWidthPx(workbook)
        val baseColWidthsPxRaw = IntArray(lastCol - firstCol + 1) { idx ->
            val col = firstCol + idx
            if (sheet.isColumnHidden(col)) return@IntArray 0
            val widthChars = sheet.getColumnWidth(col) / 256f
            max(12, (widthChars * maxDigitWidthPx + 5f).roundToInt())
        }

        val baseColWidthsPx = if (options.adaptiveColumnWidths) {
            val fit = adaptColumnWidthsPx(
                workbook = workbook,
                sheet = sheet,
                formatter = formatter,
                evaluator = evaluator,
                requestedScale = options.scale,
                maxScaledWidthPx = options.maxBitmapDimension,
                firstRow = firstRow,
                lastRow = lastRow,
                firstCol = firstCol,
                lastCol = lastCol,
                baseColWidthsPx = baseColWidthsPxRaw,
                mergeInfo = mergeInfo,
                columnFontPts = columnFontPts,
                maxCells = options.columnWidthMaxCells,
                sampleMaxTextLength = options.columnWidthSampleMaxTextLength,
                minColumnWidthPx = options.minColumnWidthPx,
                minEmptyColumnWidthPx = options.minEmptyColumnWidthPx,
                emptyColumnWidthPx = options.emptyColumnWidthPx,
                maxColumnWidthPx = options.maxColumnWidthPx,
                minFontPt = options.minFontPt,
                maxFontPt = options.maxFontPt,
            )
            fit.warning?.let { warnings += it }
            if (fit.didChange) warnings += "已自动调整列宽"
            fit.colWidthsPx
        } else {
            baseColWidthsPxRaw
        }

        val fitScale =
            fitScaleToMaxWidth(
                baseColWidthsPx = baseColWidthsPx,
                requestedScale = options.scale,
                maxWidthPx = options.maxBitmapDimension,
            )
        var scale = fitScale.scale
        var width = fitScale.widthPx

        if (width > options.maxBitmapDimension) {
            warnings += "表格太宽，已强制缩小"
        }

        val baseRowHeightsPxRaw = IntArray(lastRow - firstRow + 1) { idx ->
            val rowNum = firstRow + idx
            val row = sheet.getRow(rowNum)
            if (row?.zeroHeight == true) return@IntArray 0
            val htPt = row?.heightInPoints ?: sheet.defaultRowHeightInPoints
            max(16, ceil(htPt * 4f / 3f).toInt())
        }

        val baseRowHeightsPx = if (options.autoFitRowHeights) {
            val fit = autoFitRowHeightsPx(
                workbook = workbook,
                sheet = sheet,
                formatter = formatter,
                evaluator = evaluator,
                renderScale = scale,
                firstRow = firstRow,
                lastRow = lastRow,
                firstCol = firstCol,
                lastCol = lastCol,
                baseColWidthsPx = baseColWidthsPx,
                baseRowHeightsPx = baseRowHeightsPxRaw,
                mergeInfo = mergeInfo,
                columnFontPts = columnFontPts,
                maxCells = options.autoFitMaxCells,
                maxAutoRowHeightPx = options.maxAutoRowHeightPx,
                autoWrapOverflowText = options.autoWrapOverflowText,
                autoWrapMinTextLength = options.autoWrapMinTextLength,
                autoWrapExcludeNumeric = options.autoWrapExcludeNumeric,
                minFontPt = options.minFontPt,
                maxFontPt = options.maxFontPt,
            )
            fit.warning?.let { warnings += it }
            if (fit.didChange) warnings += "已自动调整行高以适配换行"
            fit.rowHeightsPx
        } else {
            baseRowHeightsPxRaw
        }

        val parts = planVerticalParts(
            baseRowHeightsPx = baseRowHeightsPx,
            scaledWidthPx = width,
            scale = scale,
            maxBitmapDimension = options.maxBitmapDimension,
            maxTotalPixels = options.maxTotalPixels,
        )
        if (parts.isEmpty()) {
            val bmp = Bitmap.createBitmap(800, 400, Bitmap.Config.ARGB_8888)
            val canvas = Canvas(bmp)
            canvas.drawColor(Color.WHITE)
            val p = Paint(Paint.ANTI_ALIAS_FLAG).apply {
                color = Color.DKGRAY
                textSize = 36f
            }
            canvas.drawText("无可见内容（可能都被隐藏）", 40f, 120f, p)
            try {
                onPart(0, 1, bmp)
            } finally {
                if (recycleAfterCallback) bmp.recycle()
            }
            return RenderPartsResult(partCount = 1, wasSplit = false, warnings = warnings.distinct())
        }

        if (parts.size > 1) {
            warnings += "因尺寸限制自动分段"
        }

        val partLimit = max(1, maxPartsToRender)
        for ((i, part) in parts.withIndex()) {
            if (i >= partLimit) break
            val bmp = Bitmap.createBitmap(width, part.heightPx, Bitmap.Config.ARGB_8888)
            val canvas = Canvas(bmp)
            canvas.drawColor(Color.WHITE)

            drawSheetPart(
                canvas = canvas,
                workbook = workbook,
                sheet = sheet,
                formatter = formatter,
                evaluator = evaluator,
                firstRow = firstRow,
                lastRow = lastRow,
                firstCol = firstCol,
                lastCol = lastCol,
                baseColWidthsPx = baseColWidthsPx,
                baseRowHeightsPx = baseRowHeightsPx,
                scale = scale,
                partRowStart = part.rowStart,
                partRowEnd = part.rowEnd,
                partTopOffsetPx = part.topOffsetPx,
                mergeInfo = mergeInfo,
                columnFontPts = columnFontPts,
                autoWrapOverflowText = options.autoWrapOverflowText,
                autoWrapMinTextLength = options.autoWrapMinTextLength,
                autoWrapExcludeNumeric = options.autoWrapExcludeNumeric,
                minFontPt = options.minFontPt,
                maxFontPt = options.maxFontPt,
            )

            try {
                onPart(i, parts.size, bmp)
            } finally {
                if (recycleAfterCallback) bmp.recycle()
            }
        }

        return RenderPartsResult(
            partCount = parts.size,
            wasSplit = parts.size > 1,
            warnings = warnings.distinct(),
        )
    }

    fun writeSheetPdf(
        workbook: Workbook,
        sheetIndex: Int,
        options: RenderOptions,
        out: OutputStream,
        onPartCanvas: ((partIndex: Int, partCount: Int, canvas: Canvas) -> Unit)? = null,
    ): PdfWriteResult {
        val sheet = workbook.getSheetAt(sheetIndex)

        val candidateRange = findPrintAreaRange(workbook, sheetIndex) ?: findUsedRange(sheet)
        if (candidateRange == null) {
            val pdf = PdfDocument()
            try {
                val pageInfo = PdfDocument.PageInfo.Builder(800, 400, 1).create()
                val page = pdf.startPage(pageInfo)
                val canvas = page.canvas
                canvas.drawColor(Color.WHITE)
                val p = Paint(Paint.ANTI_ALIAS_FLAG).apply {
                    color = Color.DKGRAY
                    textSize = 40f
                }
                canvas.drawText("空表", 40f, 120f, p)
                pdf.finishPage(page)
                pdf.writeTo(out)
            } finally {
                pdf.close()
            }
            return PdfWriteResult(pageCount = 1, wasSplit = false, warnings = emptyList())
        }

        val warnings = mutableListOf<String>()
        val formatter = DataFormatter()
        val evaluator = workbook.creationHelper.createFormulaEvaluator()

        val used = if (options.trimBlankEdges) {
            val trim = trimUsedRange(
                sheet = sheet,
                range = candidateRange,
                formatter = formatter,
                evaluator = evaluator,
                maxCells = options.trimMaxCells,
            )
            trim.warning?.let { warnings += it }
            if (trim.didTrim) warnings += "已自动裁剪空白边缘"
            trim.range
        } else {
            candidateRange
        }

        val (firstRow, lastRow, firstCol, lastCol) = used

        val mergeInfo = buildMergeInfo(sheet, firstRow, lastRow, firstCol, lastCol)

        val columnFontPts = if (options.uniformFontPerColumn) {
            val fit = computeColumnFontPts(
                workbook = workbook,
                sheet = sheet,
                formatter = formatter,
                evaluator = evaluator,
                firstRow = firstRow,
                lastRow = lastRow,
                firstCol = firstCol,
                lastCol = lastCol,
                mergeInfo = mergeInfo,
                maxCells = options.columnFontMaxCells,
                minFontPt = options.minFontPt,
                maxFontPt = options.maxFontPt,
            )
            fit.warning?.let { warnings += it }
            warnings += "已按列统一字号"
            fit.fontPts
        } else {
            null
        }

        val maxDigitWidthPx = computeMaxDigitWidthPx(workbook)
        val baseColWidthsPxRaw = IntArray(lastCol - firstCol + 1) { idx ->
            val col = firstCol + idx
            if (sheet.isColumnHidden(col)) return@IntArray 0
            val widthChars = sheet.getColumnWidth(col) / 256f
            max(12, (widthChars * maxDigitWidthPx + 5f).roundToInt())
        }

        val baseColWidthsPx = if (options.adaptiveColumnWidths) {
            val fit = adaptColumnWidthsPx(
                workbook = workbook,
                sheet = sheet,
                formatter = formatter,
                evaluator = evaluator,
                requestedScale = options.scale,
                maxScaledWidthPx = options.maxBitmapDimension,
                firstRow = firstRow,
                lastRow = lastRow,
                firstCol = firstCol,
                lastCol = lastCol,
                baseColWidthsPx = baseColWidthsPxRaw,
                mergeInfo = mergeInfo,
                columnFontPts = columnFontPts,
                maxCells = options.columnWidthMaxCells,
                sampleMaxTextLength = options.columnWidthSampleMaxTextLength,
                minColumnWidthPx = options.minColumnWidthPx,
                minEmptyColumnWidthPx = options.minEmptyColumnWidthPx,
                emptyColumnWidthPx = options.emptyColumnWidthPx,
                maxColumnWidthPx = options.maxColumnWidthPx,
                minFontPt = options.minFontPt,
                maxFontPt = options.maxFontPt,
            )
            fit.warning?.let { warnings += it }
            if (fit.didChange) warnings += "已自动调整列宽"
            fit.colWidthsPx
        } else {
            baseColWidthsPxRaw
        }

        // Apply initial scale. We'll shrink further if needed to respect page constraints.
        val fitScale =
            fitScaleToMaxWidth(
                baseColWidthsPx = baseColWidthsPx,
                requestedScale = options.scale,
                maxWidthPx = options.maxBitmapDimension,
            )
        var scale = fitScale.scale
        var width = fitScale.widthPx

        if (width > options.maxBitmapDimension) {
            warnings += "表格太宽，已强制缩小"
        }

        val baseRowHeightsPxRaw = IntArray(lastRow - firstRow + 1) { idx ->
            val rowNum = firstRow + idx
            val row = sheet.getRow(rowNum)
            if (row?.zeroHeight == true) return@IntArray 0
            val htPt = row?.heightInPoints ?: sheet.defaultRowHeightInPoints
            max(16, ceil(htPt * 4f / 3f).toInt())
        }

        val baseRowHeightsPx = if (options.autoFitRowHeights) {
            val fit = autoFitRowHeightsPx(
                workbook = workbook,
                sheet = sheet,
                formatter = formatter,
                evaluator = evaluator,
                renderScale = scale,
                firstRow = firstRow,
                lastRow = lastRow,
                firstCol = firstCol,
                lastCol = lastCol,
                baseColWidthsPx = baseColWidthsPx,
                baseRowHeightsPx = baseRowHeightsPxRaw,
                mergeInfo = mergeInfo,
                columnFontPts = columnFontPts,
                maxCells = options.autoFitMaxCells,
                maxAutoRowHeightPx = options.maxAutoRowHeightPx,
                autoWrapOverflowText = options.autoWrapOverflowText,
                autoWrapMinTextLength = options.autoWrapMinTextLength,
                autoWrapExcludeNumeric = options.autoWrapExcludeNumeric,
                minFontPt = options.minFontPt,
                maxFontPt = options.maxFontPt,
            )
            fit.warning?.let { warnings += it }
            if (fit.didChange) warnings += "已自动调整行高以适配换行"
            fit.rowHeightsPx
        } else {
            baseRowHeightsPxRaw
        }

        val parts = planVerticalParts(
            baseRowHeightsPx = baseRowHeightsPx,
            scaledWidthPx = width,
            scale = scale,
            maxBitmapDimension = options.maxBitmapDimension,
            maxTotalPixels = options.maxTotalPixels,
        )
        if (parts.isEmpty()) {
            val pdf = PdfDocument()
            try {
                val pageInfo = PdfDocument.PageInfo.Builder(800, 400, 1).create()
                val page = pdf.startPage(pageInfo)
                val canvas = page.canvas
                canvas.drawColor(Color.WHITE)
                val p = Paint(Paint.ANTI_ALIAS_FLAG).apply {
                    color = Color.DKGRAY
                    textSize = 36f
                }
                canvas.drawText("无可见内容（可能都被隐藏）", 40f, 120f, p)
                pdf.finishPage(page)
                pdf.writeTo(out)
            } finally {
                pdf.close()
            }
            return PdfWriteResult(pageCount = 1, wasSplit = false, warnings = warnings.distinct())
        }

        if (parts.size > 1) {
            warnings += "因尺寸限制自动分段"
        }

        val pdf = PdfDocument()
        try {
            for ((i, part) in parts.withIndex()) {
                val pageInfo = PdfDocument.PageInfo.Builder(width, part.heightPx, i + 1).create()
                val page = pdf.startPage(pageInfo)
                val canvas = page.canvas
                canvas.drawColor(Color.WHITE)

                drawSheetPart(
                    canvas = canvas,
                    workbook = workbook,
                    sheet = sheet,
                    formatter = formatter,
                    evaluator = evaluator,
                    firstRow = firstRow,
                    lastRow = lastRow,
                    firstCol = firstCol,
                    lastCol = lastCol,
                    baseColWidthsPx = baseColWidthsPx,
                    baseRowHeightsPx = baseRowHeightsPx,
                    scale = scale,
                    partRowStart = part.rowStart,
                    partRowEnd = part.rowEnd,
                    partTopOffsetPx = part.topOffsetPx,
                    mergeInfo = mergeInfo,
                    columnFontPts = columnFontPts,
                    autoWrapOverflowText = options.autoWrapOverflowText,
                    autoWrapMinTextLength = options.autoWrapMinTextLength,
                    autoWrapExcludeNumeric = options.autoWrapExcludeNumeric,
                    minFontPt = options.minFontPt,
                    maxFontPt = options.maxFontPt,
                )

                onPartCanvas?.invoke(i, parts.size, canvas)
                pdf.finishPage(page)
            }

            pdf.writeTo(out)
        } finally {
            pdf.close()
        }

        return PdfWriteResult(
            pageCount = parts.size,
            wasSplit = parts.size > 1,
            warnings = warnings.distinct(),
        )
    }

    private data class UsedRange(
        val firstRow: Int,
        val lastRow: Int,
        val firstCol: Int,
        val lastCol: Int,
    )

    private data class PartPlan(
        val rowStart: Int,
        val rowEnd: Int,
        val topOffsetPx: Int,
        val heightPx: Int,
    )

    private data class TrimOutcome(
        val range: UsedRange,
        val didTrim: Boolean,
        val warning: String? = null,
    )

    private fun trimUsedRange(
        sheet: Sheet,
        range: UsedRange,
        formatter: DataFormatter,
        evaluator: org.apache.poi.ss.usermodel.FormulaEvaluator,
        maxCells: Int,
    ): TrimOutcome {
        val mergedRegions = (0 until sheet.numMergedRegions).map { sheet.getMergedRegion(it) }

        var minRow = Int.MAX_VALUE
        var maxRow = -1
        var minCol = Int.MAX_VALUE
        var maxCol = -1

        var processed = 0

        val rowIt = sheet.rowIterator()
        while (rowIt.hasNext()) {
            val row = rowIt.next()
            val r = row.rowNum
            if (r < range.firstRow || r > range.lastRow) continue
            if (row.zeroHeight) continue

            val cellIt = row.cellIterator()
            while (cellIt.hasNext()) {
                val cell = cellIt.next()
                val c = cell.columnIndex
                if (c < range.firstCol || c > range.lastCol) continue
                if (sheet.isColumnHidden(c)) continue

                processed++
                if (processed > maxCells) {
                    return TrimOutcome(range = range, didTrim = false, warning = "表格较大，已跳过裁剪空白边缘")
                }

                if (!cellIsVisuallyUsed(cell, formatter, evaluator)) continue

                val mr = mergedRegions.firstOrNull {
                    r in it.firstRow..it.lastRow && c in it.firstColumn..it.lastColumn
                }
                if (mr != null) {
                    minRow = min(minRow, max(range.firstRow, mr.firstRow))
                    maxRow = max(maxRow, min(range.lastRow, mr.lastRow))
                    minCol = min(minCol, max(range.firstCol, mr.firstColumn))
                    maxCol = max(maxCol, min(range.lastCol, mr.lastColumn))
                } else {
                    minRow = min(minRow, r)
                    maxRow = max(maxRow, r)
                    minCol = min(minCol, c)
                    maxCol = max(maxCol, c)
                }
            }
        }

        if (maxRow < 0 || maxCol < 0 || minRow == Int.MAX_VALUE || minCol == Int.MAX_VALUE) {
            return TrimOutcome(range = range, didTrim = false)
        }

        val trimmed = UsedRange(
            firstRow = minRow.coerceIn(range.firstRow, range.lastRow),
            lastRow = maxRow.coerceIn(range.firstRow, range.lastRow),
            firstCol = minCol.coerceIn(range.firstCol, range.lastCol),
            lastCol = maxCol.coerceIn(range.firstCol, range.lastCol),
        )

        return TrimOutcome(range = trimmed, didTrim = trimmed != range)
    }

    private data class ColumnWidthOutcome(
        val colWidthsPx: IntArray,
        val didChange: Boolean,
        val warning: String? = null,
    )

    private data class ColumnFontOutcome(
        val fontPts: IntArray,
        val warning: String? = null,
    )

    private fun computeColumnFontPts(
        workbook: Workbook,
        sheet: Sheet,
        formatter: DataFormatter,
        evaluator: org.apache.poi.ss.usermodel.FormulaEvaluator,
        firstRow: Int,
        lastRow: Int,
        firstCol: Int,
        lastCol: Int,
        mergeInfo: Pair<Map<Long, MergeRegion>, Set<Long>>,
        maxCells: Int,
        minFontPt: Int,
        maxFontPt: Int,
    ): ColumnFontOutcome {
        val (cellToMerge, mergeStarts) = mergeInfo
        val colCount = lastCol - firstCol + 1
        if (colCount <= 0) {
            return ColumnFontOutcome(fontPts = IntArray(0))
        }

        val defaultPt = runCatching { workbook.getFontAt(0).fontHeightInPoints.toInt() }
            .getOrNull()
            ?.coerceIn(minFontPt, maxFontPt)
            ?: 11.coerceIn(minFontPt, maxFontPt)

        // Keep per-column samples small to avoid memory blowups on large sheets.
        val maxSamplesPerCol = 64
        val samples = Array(colCount) { IntArray(maxSamplesPerCol) }
        val sampleCounts = IntArray(colCount)

        var processed = 0
        val headerRow = firstRow

        val rowIt = sheet.rowIterator()
        while (rowIt.hasNext()) {
            val row = rowIt.next()
            val r = row.rowNum
            if (r < firstRow || r > lastRow) continue
            if (row.zeroHeight) continue

            val isHeader = r == headerRow

            val cellIt = row.cellIterator()
            while (cellIt.hasNext()) {
                val cell = cellIt.next()
                val c = cell.columnIndex
                if (c < firstCol || c > lastCol) continue
                if (sheet.isColumnHidden(c)) continue

                val text =
                    cleanCellText(runCatching { formatter.formatCellValue(cell, evaluator) }.getOrNull().orEmpty())
                if (text.isEmpty()) continue
                // Header text should not affect column width heuristics; we optimize for data rows.
                if (isHeader) continue

                val key = cellKey(r, c)
                val merge = cellToMerge[key]
                if (merge != null && key !in mergeStarts) {
                    continue
                }

                // Only consider the first cell for merged regions to avoid double-counting.
                val spanFirstCol = merge?.firstCol ?: c
                if (spanFirstCol != c) continue

                val colIdx = c - firstCol
                if (colIdx !in 0 until colCount) continue

                processed++
                if (processed > maxCells) {
                    val out = IntArray(colCount) { i ->
                        val count = sampleCounts[i]
                        if (count <= 0) defaultPt
                        else {
                            val arr = samples[i].copyOf(count)
                            arr.sort()
                            // Median font size is robust against a few large headers.
                            arr[count / 2].coerceIn(minFontPt, maxFontPt)
                        }
                    }
                    return ColumnFontOutcome(
                        fontPts = out,
                        warning = "表格较大，已限制按列统一字号的扫描范围",
                    )
                }

                val font = workbook.getFontAt(cell.cellStyle.fontIndex)
                val pt = font.fontHeightInPoints.toInt().coerceIn(minFontPt, maxFontPt)

                val count = sampleCounts[colIdx]
                if (count < maxSamplesPerCol) {
                    samples[colIdx][count] = pt
                    sampleCounts[colIdx] = count + 1
                }
            }
        }

        val out = IntArray(colCount) { i ->
            val count = sampleCounts[i]
            if (count <= 0) defaultPt
            else {
                val arr = samples[i].copyOf(count)
                arr.sort()
                arr[count / 2].coerceIn(minFontPt, maxFontPt)
            }
        }

        return ColumnFontOutcome(fontPts = out)
    }

    private fun adaptColumnWidthsPx(
        workbook: Workbook,
        sheet: Sheet,
        formatter: DataFormatter,
        evaluator: org.apache.poi.ss.usermodel.FormulaEvaluator,
        requestedScale: Float,
        maxScaledWidthPx: Int,
        firstRow: Int,
        lastRow: Int,
        firstCol: Int,
        lastCol: Int,
        baseColWidthsPx: IntArray,
        mergeInfo: Pair<Map<Long, MergeRegion>, Set<Long>>,
        columnFontPts: IntArray?,
        maxCells: Int,
        sampleMaxTextLength: Int,
        minColumnWidthPx: Int,
        minEmptyColumnWidthPx: Int,
        emptyColumnWidthPx: Int,
        maxColumnWidthPx: Int,
        minFontPt: Int,
        maxFontPt: Int,
    ): ColumnWidthOutcome {
        val (cellToMerge, mergeStarts) = mergeInfo
        val colCount = lastCol - firstCol + 1
        if (colCount <= 0) return ColumnWidthOutcome(colWidthsPx = baseColWidthsPx, didChange = false)

        // Keep per-column samples small to avoid memory blowups on large sheets.
        val maxSamplesPerCol = 64
        val samples = Array(colCount) { IntArray(maxSamplesPerCol) }
        val sampleTextLens = Array(colCount) { IntArray(maxSamplesPerCol) }
        val sampleCounts = IntArray(colCount)
        val nonEmptyAnyCounts = IntArray(colCount)
        val nonEmptyDataCounts = IntArray(colCount)
        val minFromMergedSpans = IntArray(colCount)
        // Header-driven minimum width (soft): used to prevent headers wrapping into "too many lines".
        // We keep two variants so empty/sparse columns can be narrower (allow 3-line headers).
        val minFromHeaders = IntArray(colCount) // headerMaxLines
        val minFromHeadersTight = IntArray(colCount) // headerMaxLinesTight
        // Content-type hints (based on a small per-column sample).
        val sampleNumericCounts = IntArray(colCount)
        val sampleMostlyAsciiCounts = IntArray(colCount)
        val sampleHasCjkCounts = IntArray(colCount)
        val seenDataCells = IntArray(colCount)
        val sampleIsNumeric = Array(colCount) { BooleanArray(maxSamplesPerCol) }
        val sampleIsMostlyAscii = Array(colCount) { BooleanArray(maxSamplesPerCol) }
        val sampleHasCjk = Array(colCount) { BooleanArray(maxSamplesPerCol) }
        // Track the longest content we've seen for each column (used to cap extreme wrapping even if
        // it appears outside our limited sample set).
        val maxMeasuredWithPaddingPx = IntArray(colCount)
        val maxTextLenSeen = IntArray(colCount)

        val paint = Paint(Paint.ANTI_ALIAS_FLAG)
        val padding = 4f
        // Small safety buffer to avoid borderline wraps due to font metrics / float rounding differences
        // between `measureText` and `breakText`.
        val measureFudgePx = 2f
        val padding2Int = ceil(padding * 2).toInt()
        // Keep header reasonably compact: prevent a long header label from wrapping into too many lines,
        // which would make the whole table look "very tall" even if data rows are short.
        val headerMaxLines = 2
        // For empty/sparse columns we can allow a slightly taller header to save horizontal space.
        val headerMaxLinesTight = 3
        val headerSafety = 1.10f
        val headerSafetyTight = 1.02f

        // Deterministic per-sheet PRNG for reservoir sampling (stable output across runs).
        val rng =
            kotlin.random.Random(
                (sheet.sheetName.hashCode() * 31) xor firstRow xor (lastRow shl 16) xor (firstCol shl 8) xor colCount,
            )

        fun isCjk(ch: Char): Boolean {
            val c = ch.code
            return (c in 0x4E00..0x9FFF) || (c in 0x3400..0x4DBF)
        }

        fun hasCjk(text: String): Boolean {
            for (ch in text) {
                if (isCjk(ch)) return true
            }
            return false
        }

        fun isMostlyAsciiToken(text: String): Boolean {
            val t = text.trim()
            if (t.isEmpty()) return false
            var ascii = 0
            var nonWs = 0
            for (ch in t) {
                if (ch == '\u00A0' || ch.isWhitespace()) continue
                nonWs += 1
                if (ch.code in 33..126) ascii += 1
            }
            if (nonWs <= 0) return false
            return ascii.toFloat() / nonWs.toFloat() >= 0.95f
        }

        var processed = 0
        val headerRow = firstRow
        var visibleDataRows = 0

        val rowIt = sheet.rowIterator()
        while (rowIt.hasNext()) {
            val row = rowIt.next()
            val r = row.rowNum
            if (r < firstRow || r > lastRow) continue
            if (row.zeroHeight) continue

            val isHeader = r == headerRow
            if (!isHeader) {
                visibleDataRows++
            }

            val cellIt = row.cellIterator()
            while (cellIt.hasNext()) {
                val cell = cellIt.next()
                val c = cell.columnIndex
                if (c < firstCol || c > lastCol) continue
                if (sheet.isColumnHidden(c)) continue

                val text =
                    cleanCellText(runCatching { formatter.formatCellValue(cell, evaluator) }.getOrNull().orEmpty())
                if (text.isEmpty()) continue

                val key = cellKey(r, c)
                val merge = cellToMerge[key]
                if (merge != null && key !in mergeStarts) {
                    continue
                }

                // Only consider the first cell for merged regions to avoid double-counting.
                val spanFirstCol = merge?.firstCol ?: c
                val spanLastCol = merge?.lastCol ?: c
                if (spanFirstCol < firstCol || spanLastCol > lastCol) continue
                if (spanFirstCol != c) continue

                if (spanFirstCol != spanLastCol) {
                    val span = (spanLastCol - spanFirstCol + 1).coerceAtLeast(1)
                    for (absCol in spanFirstCol..spanLastCol) {
                        val idx = absCol - firstCol
                        if (idx in 0 until colCount) {
                            nonEmptyAnyCounts[idx] += 1
                            if (!isHeader) nonEmptyDataCounts[idx] += 1
                        }
                    }

                    if (isHeader) {
                        // Only enforce a soft minimum so header won't explode vertically.
                        val font = workbook.getFontAt(cell.cellStyle.fontIndex)
                        val colIdxForFont = c - firstCol
                        val overridePt = columnFontPts?.getOrNull(colIdxForFont)
                        applyFont(paint, font, 1f, minFontPt, maxFontPt, overridePt = overridePt)

                        val textForMeasure =
                            if (sampleMaxTextLength > 0 && text.length > sampleMaxTextLength) {
                                text.substring(0, sampleMaxTextLength)
                            } else {
                                text
                            }
                        val measured0 = measureMaxLineWidthPx(textForMeasure, paint)
                        val measured =
                            if (textForMeasure.length in 1 until text.length) {
                                measured0 * (text.length.toFloat() / textForMeasure.length.toFloat())
                            } else {
                                measured0
                            }
                        val measuredWithPadding = ceil(measured + padding * 2 + measureFudgePx).toInt()
                        val contentWidth = max(0, measuredWithPadding - padding2Int)

                        // NOTE: padding is per-cell, not per-line. So it must NOT be divided by headerMaxLines.
                        val needTotal =
                            ceil(
                                (contentWidth.toFloat() / headerMaxLines.toFloat() + padding2Int.toFloat()) *
                                    headerSafety,
                            ).toInt()
                        val needTotalTight =
                            ceil(
                                (contentWidth.toFloat() / headerMaxLinesTight.toFloat() + padding2Int.toFloat()) *
                                    headerSafetyTight,
                            ).toInt()
                        val perCol = ceil(needTotal.toFloat() / span.toFloat()).toInt()
                        val perColTight = ceil(needTotalTight.toFloat() / span.toFloat()).toInt()
                        for (absCol in spanFirstCol..spanLastCol) {
                            val idx = absCol - firstCol
                            if (idx in 0 until colCount) {
                                minFromHeaders[idx] = max(minFromHeaders[idx], perCol)
                                minFromHeadersTight[idx] = max(minFromHeadersTight[idx], perColTight)
                            }
                        }
                        continue
                    }

                    processed++
                    if (processed > maxCells) {
                        return ColumnWidthOutcome(
                            colWidthsPx = baseColWidthsPx,
                            didChange = false,
                            warning = "表格较大，已跳过自适应列宽",
                        )
                    }

                    // Measure using the cell's font (clamped) so the estimate matches rendering.
                    val font = workbook.getFontAt(cell.cellStyle.fontIndex)
                    val colIdxForFont = c - firstCol
                    val overridePt = columnFontPts?.getOrNull(colIdxForFont)
                    applyFont(paint, font, 1f, minFontPt, maxFontPt, overridePt = overridePt)

                    val textForMeasure =
                        if (sampleMaxTextLength > 0 && text.length > sampleMaxTextLength) {
                            text.substring(0, sampleMaxTextLength)
                        } else {
                            text
                        }
                    val measured0 = measureMaxLineWidthPx(textForMeasure, paint)
                    val measured =
                        if (textForMeasure.length in 1 until text.length) {
                            measured0 * (text.length.toFloat() / textForMeasure.length.toFloat())
                        } else {
                            measured0
                        }
                    val measuredWithPadding = ceil(measured + padding * 2 + measureFudgePx).toInt()

                    // Distribute merged-cell width across its spanned columns (best-effort).
                    val perCol = ceil(measuredWithPadding.toFloat() / span.toFloat()).toInt()
                    for (absCol in spanFirstCol..spanLastCol) {
                        val idx = absCol - firstCol
                        if (idx in 0 until colCount) {
                            minFromMergedSpans[idx] = max(minFromMergedSpans[idx], perCol)
                        }
                    }
                    continue
                }

                val colIdx = c - firstCol
                if (colIdx !in 0 until colCount) continue
                nonEmptyAnyCounts[colIdx] += 1
                if (!isHeader) nonEmptyDataCounts[colIdx] += 1

                // Width should be driven by data rows; skip sampling header.
                if (isHeader) {
                    // Header contributes only a soft min width to prevent extreme wrapping.
                    val font = workbook.getFontAt(cell.cellStyle.fontIndex)
                    val overridePt = columnFontPts?.getOrNull(colIdx)
                    applyFont(paint, font, 1f, minFontPt, maxFontPt, overridePt = overridePt)

                    val textForMeasure =
                        if (sampleMaxTextLength > 0 && text.length > sampleMaxTextLength) {
                            text.substring(0, sampleMaxTextLength)
                        } else {
                            text
                        }
                    val measured0 = measureMaxLineWidthPx(textForMeasure, paint)
                    val measured =
                        if (textForMeasure.length in 1 until text.length) {
                            measured0 * (text.length.toFloat() / textForMeasure.length.toFloat())
                        } else {
                            measured0
                        }
                    val measuredWithPadding = ceil(measured + padding * 2 + measureFudgePx).toInt()
                    val contentWidth = max(0, measuredWithPadding - padding2Int)

                    // NOTE: padding is per-cell, not per-line. So it must NOT be divided by headerMaxLines.
                    val need =
                        ceil(
                            (contentWidth.toFloat() / headerMaxLines.toFloat() + padding2Int.toFloat()) *
                                headerSafety,
                        ).toInt()
                    val needTight =
                        ceil(
                            (contentWidth.toFloat() / headerMaxLinesTight.toFloat() + padding2Int.toFloat()) *
                                headerSafetyTight,
                        ).toInt()
                    minFromHeaders[colIdx] = max(minFromHeaders[colIdx], need)
                    minFromHeadersTight[colIdx] = max(minFromHeadersTight[colIdx], needTight)
                    continue
                }

                // Deterministic reservoir sampling:
                // - Avoid biasing toward the top rows (important for big sheets)
                // - Still keep per-column samples small to avoid memory blowups
                val seen = (seenDataCells[colIdx] + 1).also { seenDataCells[colIdx] = it }
                val count = sampleCounts[colIdx]
                val slot = when {
                    count < maxSamplesPerCol -> count
                    else -> {
                        val j = rng.nextInt(seen)
                        if (j < maxSamplesPerCol) j else -1
                    }
                }
                val mustMeasureForMax = text.length > maxTextLenSeen[colIdx]
                if (slot < 0 && !mustMeasureForMax) continue

                processed++
                if (processed > maxCells) {
                    return ColumnWidthOutcome(
                        colWidthsPx = baseColWidthsPx,
                        didChange = false,
                        warning = "表格较大，已跳过自适应列宽",
                    )
                }

                // Measure using the cell's font (clamped) so the estimate matches rendering.
                val font = workbook.getFontAt(cell.cellStyle.fontIndex)
                val overridePt = columnFontPts?.getOrNull(colIdx)
                applyFont(paint, font, 1f, minFontPt, maxFontPt, overridePt = overridePt)

                val textForMeasure =
                    if (sampleMaxTextLength > 0 && text.length > sampleMaxTextLength) {
                        text.substring(0, sampleMaxTextLength)
                    } else {
                        text
                    }
                val measured0 = measureMaxLineWidthPx(textForMeasure, paint)
                val measured =
                    if (textForMeasure.length in 1 until text.length) {
                        measured0 * (text.length.toFloat() / textForMeasure.length.toFloat())
                    } else {
                        measured0
                    }
                val measuredWithPadding = ceil(measured + padding * 2 + measureFudgePx).toInt()

                maxTextLenSeen[colIdx] = max(maxTextLenSeen[colIdx], text.length)
                maxMeasuredWithPaddingPx[colIdx] = max(maxMeasuredWithPaddingPx[colIdx], measuredWithPadding)

                if (slot >= 0) {
                    // Update sample classification counters if replacing an existing slot.
                    if (slot < count) {
                        if (sampleIsNumeric[colIdx][slot]) sampleNumericCounts[colIdx] -= 1
                        if (sampleHasCjk[colIdx][slot]) sampleHasCjkCounts[colIdx] -= 1
                        if (sampleIsMostlyAscii[colIdx][slot]) sampleMostlyAsciiCounts[colIdx] -= 1
                    }

                    val isNum = looksNumeric(text)
                    val hasC = hasCjk(text)
                    val isAscii = isMostlyAsciiToken(text)

                    samples[colIdx][slot] = measuredWithPadding
                    sampleTextLens[colIdx][slot] = text.length
                    sampleIsNumeric[colIdx][slot] = isNum
                    sampleHasCjk[colIdx][slot] = hasC
                    sampleIsMostlyAscii[colIdx][slot] = isAscii
                    if (isNum) sampleNumericCounts[colIdx] += 1
                    if (hasC) sampleHasCjkCounts[colIdx] += 1
                    if (isAscii) sampleMostlyAsciiCounts[colIdx] += 1

                    if (slot == count) {
                        sampleCounts[colIdx] = count + 1
                    }
                }
            }
        }

        val out = baseColWidthsPx.clone()
        var changed = false

        val scale = requestedScale.coerceAtLeast(0.1f)
        val budgetScaled = if (maxScaledWidthPx > 0) maxScaledWidthPx else Int.MAX_VALUE
        // For global tuning: store 2-line and 3-line candidates for text columns.
        val cand2 = IntArray(colCount)
        val cand3 = IntArray(colCount)
        val canUse3 = BooleanArray(colCount)

        fun pIndex(count: Int, p: Float): Int {
            if (count <= 1) return 0
            return ((count - 1) * p).toInt().coerceIn(0, count - 1)
        }

        for (i in 0 until colCount) {
            val base = baseColWidthsPx[i]
            if (base <= 0) {
                out[i] = 0
                continue
            }

            val anyCount = nonEmptyAnyCounts[i]
            val dataCount = nonEmptyDataCounts[i]
            val density = dataCount.toFloat() / max(1, visibleDataRows).toFloat()
            val isSparse = density < 0.02f

            val newWidth = if (anyCount <= 0 || dataCount <= 0) {
                // Empty in data rows (even if header has text).
                val shrink = min(base, emptyColumnWidthPx).coerceAtLeast(minEmptyColumnWidthPx)
                max(shrink, minFromHeadersTight[i]).coerceAtMost(maxColumnWidthPx)
            } else if (isSparse) {
                // "Mostly empty" columns waste a lot of horizontal space but add little information.
                // Shrink aggressively so important text columns can stay wider (reducing extreme wrapping).
                val shrink = min(base, emptyColumnWidthPx).coerceAtLeast(minEmptyColumnWidthPx)
                max(shrink, minFromHeadersTight[i]).coerceAtMost(maxColumnWidthPx)
            } else {
                val count = sampleCounts[i]
                if (count <= 0) {
                    // Fallback when we couldn't sample any cell (rare).
                    val withMerged = max(min(base, 220).coerceAtLeast(minColumnWidthPx), max(minFromMergedSpans[i], minFromHeaders[i]))
                    withMerged.coerceIn(minColumnWidthPx, maxColumnWidthPx)
                } else {
                    val lensSorted = run {
                        val lens = sampleTextLens[i].copyOf(count)
                        lens.sort()
                        lens
                    }
                    val p90Idx = pIndex(count, 0.9f)
                    val p90Len = lensSorted[p90Idx]
                    val medianLen = lensSorted[count / 2]

                    val numericRatio = sampleNumericCounts[i].toFloat() / count.toFloat()
                    val asciiRatio = sampleMostlyAsciiCounts[i].toFloat() / count.toFloat()
                    val cjkRatio = sampleHasCjkCounts[i].toFloat() / count.toFloat()

                    val arrW1 = run {
                        val arr = samples[i].copyOf(count)
                        arr.sort()
                        arr
                    }
                    val w1Need = arrW1[p90Idx].coerceAtLeast(minColumnWidthPx)
                    val req2 = IntArray(count) { j ->
                        val sampleW = samples[i][j]
                        val content = max(0, sampleW - padding2Int)
                        ceil(content.toFloat() / 2f + padding2Int.toFloat()).toInt()
                            .coerceIn(minColumnWidthPx, maxColumnWidthPx)
                    }.also { it.sort() }
                    val req3 = IntArray(count) { j ->
                        val sampleW = samples[i][j]
                        val content = max(0, sampleW - padding2Int)
                        ceil(content.toFloat() / 3f + padding2Int.toFloat()).toInt()
                            .coerceIn(minColumnWidthPx, maxColumnWidthPx)
                    }.also { it.sort() }

                    val w2Need0 = req2[p90Idx].coerceAtLeast(minColumnWidthPx)
                    val w3Need0 = req3[p90Idx].coerceAtLeast(minColumnWidthPx)

                    val baseMin = max(minFromMergedSpans[i], minFromHeaders[i]).coerceAtLeast(minColumnWidthPx)
                    val baseCapped = min(base, maxColumnWidthPx).coerceAtLeast(minColumnWidthPx)
                    val floorRatio = when {
                        medianLen >= 18 -> 0.55f
                        medianLen >= 12 -> 0.45f
                        else -> 0.25f
                    }
                    val floorFromExcel = max(minColumnWidthPx, (baseCapped * floorRatio).roundToInt())

                    val isNumericCol = numericRatio >= 0.80f
                    val isAsciiCol = asciiRatio >= 0.80f && cjkRatio < 0.20f
                    val isShortCol = p90Len < 10

                    if (isNumericCol || isAsciiCol || isShortCol) {
                        val withMerged = max(w1Need, baseMin)
                        max(withMerged, floorFromExcel).coerceIn(minColumnWidthPx, maxColumnWidthPx)
                    } else {
                        val w2 = max(w2Need0, baseMin).coerceIn(minColumnWidthPx, maxColumnWidthPx)
                        val w3 = max(w3Need0, baseMin).coerceIn(minColumnWidthPx, maxColumnWidthPx)
                        cand2[i] = w2
                        cand3[i] = min(w2, w3)
                        canUse3[i] = p90Len >= 12 && cand3[i] < cand2[i]
                        w2
                    }
                }
            }

            if (newWidth != base) {
                out[i] = newWidth
                changed = true
            }
        }

        // Global re-balancing under a width budget:
        // Prefer a 2-line layout for text columns, but downgrade a few "heavy" text columns to 3 lines
        // when the table would otherwise exceed max width (so we avoid tiny scale / unreadable exports).
        if (budgetScaled != Int.MAX_VALUE) {
            var total = scaledSumPx(out, scale)

            if (total > budgetScaled) {
                val degradeOrder =
                    (0 until colCount)
                        .filter { canUse3[it] && cand3[it] > 0 && cand2[it] > 0 && cand3[it] < out[it] }
                        .sortedByDescending { scaledPx(out[it], scale) - scaledPx(cand3[it], scale) }

                for (i in degradeOrder) {
                    if (total <= budgetScaled) break
                    val old = out[i]
                    val target = cand3[i]
                    if (target <= 0 || target >= old) continue
                    val saving = scaledPx(old, scale) - scaledPx(target, scale)
                    if (saving <= 0) continue
                    out[i] = target
                    total -= saving
                    changed = true
                }
            }

            // If the last downgrade over-saved, use remaining slack to upgrade the cheapest columns
            // back to 2 lines (minimizes the number of 3-line columns).
            var slack = budgetScaled - total
            if (slack > 0) {
                val upgradeOrder =
                    (0 until colCount)
                        .filter { cand2[it] > 0 && cand3[it] > 0 && out[it] == cand3[it] && cand2[it] > cand3[it] }
                        .sortedBy { scaledPx(cand2[it], scale) - scaledPx(out[it], scale) }

                for (i in upgradeOrder) {
                    if (slack <= 0) break
                    val old = out[i]
                    val target = cand2[i]
                    val cost = scaledPx(target, scale) - scaledPx(old, scale)
                    if (cost <= 0) continue
                    if (cost > slack) continue
                    out[i] = target
                    slack -= cost
                    changed = true
                }
            }
        }

        return ColumnWidthOutcome(colWidthsPx = out, didChange = changed)
    }

    private data class AutoFitOutcome(
        val rowHeightsPx: IntArray,
        val didChange: Boolean,
        val warning: String? = null,
    )

    private fun autoFitRowHeightsPx(
        workbook: Workbook,
        sheet: Sheet,
        formatter: DataFormatter,
        evaluator: org.apache.poi.ss.usermodel.FormulaEvaluator,
        renderScale: Float,
        firstRow: Int,
        lastRow: Int,
        firstCol: Int,
        lastCol: Int,
        baseColWidthsPx: IntArray,
        baseRowHeightsPx: IntArray,
        mergeInfo: Pair<Map<Long, MergeRegion>, Set<Long>>,
        columnFontPts: IntArray?,
        maxCells: Int,
        maxAutoRowHeightPx: Int,
        autoWrapOverflowText: Boolean,
        autoWrapMinTextLength: Int,
        autoWrapExcludeNumeric: Boolean,
        minFontPt: Int,
        maxFontPt: Int,
    ): AutoFitOutcome {
        val (cellToMerge, mergeStarts) = mergeInfo

        val out = baseRowHeightsPx.clone()
        var changed = false

        val scale = renderScale.coerceAtLeast(0.1f)
        // Must match drawSheetPart's padding so height estimates are consistent with rendering.
        val padding = max(2f, 4f * scale)
        val paint = Paint(Paint.ANTI_ALIAS_FLAG)

        var processed = 0

        val rowIt = sheet.rowIterator()
        while (rowIt.hasNext()) {
            val row = rowIt.next()
            val r = row.rowNum
            if (r < firstRow || r > lastRow) continue
            if (row.zeroHeight) continue

            val rowIdx = r - firstRow
            if (rowIdx !in out.indices) continue
            if (out[rowIdx] <= 0) continue

            val cellIt = row.cellIterator()
            while (cellIt.hasNext()) {
                val cell = cellIt.next()
                val c = cell.columnIndex
                if (c < firstCol || c > lastCol) continue
                if (sheet.isColumnHidden(c)) continue

                val style = cell.cellStyle
                val rawText = runCatching { formatter.formatCellValue(cell, evaluator) }.getOrNull().orEmpty()
                if (rawText.isBlank()) continue

                val key = cellKey(r, c)
                val merge = cellToMerge[key]
                if (merge != null && key !in mergeStarts) {
                    continue
                }

                processed++
                if (processed > maxCells) {
                    return AutoFitOutcome(
                        rowHeightsPx = out,
                        didChange = changed,
                        warning = "表格较大，已限制自适应行高的扫描范围",
                    )
                }

                val spanFirstRow = merge?.firstRow ?: r
                val spanLastRow = merge?.lastRow ?: r
                val spanFirstCol = merge?.firstCol ?: c
                val spanLastCol = merge?.lastCol ?: c

                var widthPx = 0
                for (absCol in spanFirstCol..spanLastCol) {
                    val idx = absCol - firstCol
                    if (idx !in baseColWidthsPx.indices) continue
                    widthPx += scaledPx(baseColWidthsPx[idx], scale)
                }
                if (widthPx <= 0) continue

                val availableWidth = max(0f, widthPx.toFloat() - padding * 2)
                val font = workbook.getFontAt(style.fontIndex)
                val colIdxForFont = spanFirstCol - firstCol
                val overridePt = columnFontPts?.getOrNull(colIdxForFont)
                applyFont(paint, font, scale, minFontPt, maxFontPt, overridePt = overridePt)

                val hasNewline = rawText.contains('\n') || rawText.contains('\r')
                val wrap =
                    style.wrapText ||
                        hasNewline ||
                        (autoWrapOverflowText &&
                            shouldAutoWrapText(
                                text = rawText,
                                paint = paint,
                                availableWidth = availableWidth,
                                minTextLength = autoWrapMinTextLength,
                                excludeNumeric = autoWrapExcludeNumeric,
                            ))
                val finalWrap =
                    if (!wrap && columnFontPts != null && availableWidth > 0f) {
                        // When text size is uniform per column we avoid per-cell shrinking; force wrap on overflow.
                        measureMaxLineWidthPx(rawText, paint) > availableWidth * 1.01f
                    } else {
                        wrap
                    }
                if (!finalWrap) continue

                val lineCount = countLinesForLayout(rawText, paint, availableWidth, true)
                val required = ceil(padding * 2 + lineCount * paint.fontSpacing).toInt()

                if (spanFirstRow == spanLastRow) {
                    // Convert back to *base* px (before applying `scale`) because the rest of the
                    // pipeline stores row heights in base px.
                    val requiredBase = ceil(required.toDouble() / scale.toDouble()).toInt()
                    val capped = min(requiredBase, maxAutoRowHeightPx)
                    if (capped > out[rowIdx]) {
                        out[rowIdx] = capped
                        changed = true
                    }
                } else {
                    // Multi-row merged cells: ensure the *total* merged height can fit the wrapped text.
                    val spanStartIdx = spanFirstRow - firstRow
                    val spanEndIdx = spanLastRow - firstRow
                    if (spanStartIdx !in out.indices || spanEndIdx !in out.indices) continue

                    var currentTotal = 0
                    val rows = ArrayList<Int>(spanEndIdx - spanStartIdx + 1)
                    for (i in spanStartIdx..spanEndIdx) {
                        val h = out[i]
                        // Don't "unhide" zero-height rows.
                        if (h > 0) {
                            currentTotal += scaledPx(h, scale)
                            rows += i
                        }
                    }
                    if (rows.isEmpty()) continue

                    val maxTotal = scaledPx(maxAutoRowHeightPx, scale) * rows.size
                    val cappedTotal = min(required, maxTotal)
                    if (cappedTotal <= currentTotal) continue

                    var extra = cappedTotal - currentTotal
                    val per = extra / rows.size
                    var rem = extra % rows.size

                    for (i in rows) {
                        var add = per
                        if (rem > 0) {
                            add += 1
                            rem -= 1
                        }
                        if (add <= 0) continue

                        val old = out[i]
                        // Convert scaled "add" back to base px and cap per-row.
                        val addBase = ceil(add.toDouble() / scale.toDouble()).toInt().coerceAtLeast(1)
                        val newH = min(old + addBase, maxAutoRowHeightPx)
                        if (newH != old) {
                            out[i] = newH
                            changed = true
                            extra -= (scaledPx(newH, scale) - scaledPx(old, scale))
                        }
                        if (extra <= 0) break
                    }
                }
            }
        }

        return AutoFitOutcome(rowHeightsPx = out, didChange = changed)
    }

    private fun countLinesForLayout(text: String, paint: Paint, maxWidth: Float, wrap: Boolean): Int {
        if (text.isEmpty()) return 0

        val paragraphs = text.split("\r\n", "\n", "\r")
        if (paragraphs.isEmpty()) return 0

        var count = 0
        for (para in paragraphs) {
            if (!wrap || maxWidth <= 0f) {
                count += 1
            } else {
                count += breakSingleLine(para, paint, maxWidth).size
            }
        }

        return max(1, count)
    }

    private fun cellIsVisuallyUsed(
        cell: Cell,
        formatter: DataFormatter,
        evaluator: org.apache.poi.ss.usermodel.FormulaEvaluator,
    ): Boolean {
        val style = cell.cellStyle
        val hasBorder =
            style.borderTop != BorderStyle.NONE ||
                style.borderRight != BorderStyle.NONE ||
                style.borderBottom != BorderStyle.NONE ||
                style.borderLeft != BorderStyle.NONE

        val hasFill = backgroundColorArgb(style) != null

        if (hasBorder || hasFill) return true

        val text = runCatching { formatter.formatCellValue(cell, evaluator) }.getOrNull().orEmpty()
        return text.isNotBlank()
    }

    private fun findUsedRange(sheet: Sheet): UsedRange? {
        var firstRow = Int.MAX_VALUE
        var lastRow = -1
        var firstCol = Int.MAX_VALUE
        var lastCol = -1

        val rowIt = sheet.rowIterator()
        while (rowIt.hasNext()) {
            val row = rowIt.next()
            val r = row.rowNum
            val fc = row.firstCellNum.toInt()
            val lc = row.lastCellNum.toInt() - 1
            if (fc >= 0 && lc >= fc) {
                firstRow = min(firstRow, r)
                lastRow = max(lastRow, r)
                firstCol = min(firstCol, fc)
                lastCol = max(lastCol, lc)
            }
        }

        // Include merged regions even if the row/cell iterators are sparse.
        for (i in 0 until sheet.numMergedRegions) {
            val region: CellRangeAddress = sheet.getMergedRegion(i)
            firstRow = min(firstRow, region.firstRow)
            lastRow = max(lastRow, region.lastRow)
            firstCol = min(firstCol, region.firstColumn)
            lastCol = max(lastCol, region.lastColumn)
        }

        if (lastRow < 0 || lastCol < 0 || firstRow == Int.MAX_VALUE || firstCol == Int.MAX_VALUE) {
            return null
        }

        return UsedRange(firstRow, lastRow, firstCol, lastCol)
    }

    private fun findPrintAreaRange(workbook: Workbook, sheetIndex: Int): UsedRange? {
        val printArea = workbook.getPrintArea(sheetIndex) ?: return null
        if (printArea.isBlank()) return null

        // Examples:
        // - Sheet1!$A$1:$H$20
        // - 'My Sheet'!$A$1:$H$20,$J$1:$L$10
        val afterBang = printArea.substringAfter('!', printArea)
        val ranges = afterBang.split(',').map { it.replace("$", "").trim() }.filter { it.isNotEmpty() }
        if (ranges.isEmpty()) return null

        var firstRow = Int.MAX_VALUE
        var lastRow = -1
        var firstCol = Int.MAX_VALUE
        var lastCol = -1

        for (r in ranges) {
            val addr = runCatching { CellRangeAddress.valueOf(r) }.getOrNull() ?: continue
            firstRow = min(firstRow, addr.firstRow)
            lastRow = max(lastRow, addr.lastRow)
            firstCol = min(firstCol, addr.firstColumn)
            lastCol = max(lastCol, addr.lastColumn)
        }

        if (lastRow < 0 || lastCol < 0 || firstRow == Int.MAX_VALUE || firstCol == Int.MAX_VALUE) {
            return null
        }

        return UsedRange(firstRow, lastRow, firstCol, lastCol)
    }

    private fun buildMergeInfo(
        sheet: Sheet,
        exportFirstRow: Int,
        exportLastRow: Int,
        exportFirstCol: Int,
        exportLastCol: Int,
    ): Pair<Map<Long, MergeRegion>, Set<Long>> {
        val cellToRegion = HashMap<Long, MergeRegion>()
        val regionStarts = HashSet<Long>()

        for (i in 0 until sheet.numMergedRegions) {
            val region: CellRangeAddress = sheet.getMergedRegion(i)
            val mr = MergeRegion(
                firstRow = max(region.firstRow, exportFirstRow),
                lastRow = min(region.lastRow, exportLastRow),
                firstCol = max(region.firstColumn, exportFirstCol),
                lastCol = min(region.lastColumn, exportLastCol),
            )
            if (mr.firstRow > mr.lastRow || mr.firstCol > mr.lastCol) continue

            for (r in mr.firstRow..mr.lastRow) {
                for (c in mr.firstCol..mr.lastCol) {
                    cellToRegion[cellKey(r, c)] = mr
                }
            }
            regionStarts.add(cellKey(mr.firstRow, mr.firstCol))
        }

        return cellToRegion to regionStarts
    }

    private fun planVerticalParts(
        baseRowHeightsPx: IntArray,
        scaledWidthPx: Int,
        scale: Float,
        maxBitmapDimension: Int,
        maxTotalPixels: Long,
    ): List<PartPlan> {
        if (scaledWidthPx <= 0) return emptyList()

        val scaledRowHeights = baseRowHeightsPx.map { h -> scaledPx(h, scale) }
        val totalHeight = scaledRowHeights.sum()

        // Single-part fast path.
        if (scaledWidthPx <= maxBitmapDimension &&
            totalHeight <= maxBitmapDimension &&
            scaledWidthPx.toLong() * totalHeight.toLong() <= maxTotalPixels
        ) {
            return listOf(
                PartPlan(
                    rowStart = 0,
                    rowEnd = scaledRowHeights.lastIndex,
                    topOffsetPx = 0,
                    heightPx = totalHeight,
                ),
            )
        }

        if (totalHeight <= 0) return emptyList()

        val maxHeightByPixels = run {
            val h = maxTotalPixels / scaledWidthPx.toLong()
            val hInt = if (h > Int.MAX_VALUE.toLong()) Int.MAX_VALUE else h.toInt()
            max(200, hInt)
        }
        val partMaxHeight = min(maxBitmapDimension, maxHeightByPixels)

        val parts = mutableListOf<PartPlan>()
        var start = 0
        var acc = 0
        var topOffset = 0

        for (i in scaledRowHeights.indices) {
            val h = scaledRowHeights[i]
            if (acc > 0 && acc + h > partMaxHeight) {
                parts += PartPlan(
                    rowStart = start,
                    rowEnd = i - 1,
                    topOffsetPx = topOffset,
                    heightPx = acc,
                )
                topOffset += acc
                start = i
                acc = 0
            }
            acc += h
        }

        if (acc > 0) {
            parts += PartPlan(
                rowStart = start,
                rowEnd = scaledRowHeights.lastIndex,
                topOffsetPx = topOffset,
                heightPx = acc,
            )
        }

        return parts
    }

    private fun drawSheetPart(
        canvas: Canvas,
        workbook: Workbook,
        sheet: Sheet,
        formatter: DataFormatter,
        evaluator: org.apache.poi.ss.usermodel.FormulaEvaluator,
        firstRow: Int,
        lastRow: Int,
        firstCol: Int,
        lastCol: Int,
        baseColWidthsPx: IntArray,
        baseRowHeightsPx: IntArray,
        scale: Float,
        partRowStart: Int,
        partRowEnd: Int,
        partTopOffsetPx: Int,
        mergeInfo: Pair<Map<Long, MergeRegion>, Set<Long>>,
        columnFontPts: IntArray?,
        autoWrapOverflowText: Boolean,
        autoWrapMinTextLength: Int,
        autoWrapExcludeNumeric: Boolean,
        minFontPt: Int,
        maxFontPt: Int,
    ) {
        val (cellToMerge, mergeStarts) = mergeInfo

        val colCount = lastCol - firstCol + 1
        val rowCount = lastRow - firstRow + 1

        val uniformTextSize = columnFontPts != null

        val colWidths = IntArray(colCount) { idx -> scaledPx(baseColWidthsPx[idx], scale) }
        val rowHeights = IntArray(rowCount) { idx -> scaledPx(baseRowHeightsPx[idx], scale) }

        val x = IntArray(colCount + 1)
        for (i in 0 until colCount) x[i + 1] = x[i] + colWidths[i]

        val y = IntArray(rowCount + 1)
        for (i in 0 until rowCount) y[i + 1] = y[i] + rowHeights[i]

        val gridPaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
            style = Paint.Style.STROKE
            color = Color.rgb(210, 210, 210)
            strokeWidth = max(1f, scale)
        }

        val borderPaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
            style = Paint.Style.STROKE
            color = Color.BLACK
            strokeWidth = max(1f, scale)
        }

        val fillPaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
            style = Paint.Style.FILL
            color = Color.WHITE
        }

        val textPaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
            color = Color.BLACK
            textSize = max(10f, 14f * scale)
        }

        val padding = max(2f, 4f * scale)

        val absoluteRowStart = firstRow + partRowStart
        val absoluteRowEnd = firstRow + partRowEnd

        for (absRow in absoluteRowStart..absoluteRowEnd) {
            val rowIdx = absRow - firstRow
            val top = y[rowIdx] - partTopOffsetPx
            val bottom = y[rowIdx + 1] - partTopOffsetPx
            if (bottom <= top) continue

            val row = sheet.getRow(absRow)

            for (absCol in firstCol..lastCol) {
                val colIdx = absCol - firstCol
                val left = x[colIdx]
                val right = x[colIdx + 1]
                if (right <= left) continue

                val key = cellKey(absRow, absCol)
                val merge = cellToMerge[key]
                if (merge != null && key !in mergeStarts) {
                    continue
                }

                val rect = if (merge != null) {
                    val mergeTop = y[merge.firstRow - firstRow] - partTopOffsetPx
                    val mergeBottom = y[merge.lastRow - firstRow + 1] - partTopOffsetPx
                    val mergeLeft = x[merge.firstCol - firstCol]
                    val mergeRight = x[merge.lastCol - firstCol + 1]
                    RectF(
                        mergeLeft.toFloat(),
                        mergeTop.toFloat(),
                        mergeRight.toFloat(),
                        mergeBottom.toFloat(),
                    )
                } else {
                    RectF(left.toFloat(), top.toFloat(), right.toFloat(), bottom.toFloat())
                }

                val cell = row?.getCell(absCol)

                val bg = cell?.let { backgroundColorArgb(it.cellStyle) }
                if (bg != null) {
                    fillPaint.color = bg
                    canvas.drawRect(rect, fillPaint)
                }

                // Default light grid.
                canvas.drawRect(rect, gridPaint)

                if (cell != null) {
                    val style = cell.cellStyle
                    val borders = listOf(
                        style.borderTop to "top",
                        style.borderRight to "right",
                        style.borderBottom to "bottom",
                        style.borderLeft to "left",
                    )
                    if (borders.any { it.first != BorderStyle.NONE }) {
                        borderPaint.strokeWidth = max(1f, scale)
                        borderPaint.color = Color.BLACK
                        // Draw a simple rectangle border for now (best-effort).
                        canvas.drawRect(rect, borderPaint)
                    }

                    val text =
                        runCatching { formatter.formatCellValue(cell, evaluator) }.getOrNull().orEmpty()
                    if (text.isNotBlank()) {
                        val font = workbook.getFontAt(style.fontIndex)
                        val colIdxForFont = (merge?.firstCol ?: absCol) - firstCol
                        val overridePt = columnFontPts?.getOrNull(colIdxForFont)
                        applyFont(textPaint, font, scale, minFontPt, maxFontPt, overridePt = overridePt)
                        val alignH = style.alignment
                        val alignV = style.verticalAlignment
                        val wrap0 =
                            style.wrapText ||
                                text.contains('\n') ||
                                text.contains('\r') ||
                                (autoWrapOverflowText &&
                                    shouldAutoWrapText(
                                        text = text,
                                        paint = textPaint,
                                        availableWidth = max(0f, rect.width() - padding * 2),
                                        minTextLength = autoWrapMinTextLength,
                                        excludeNumeric = autoWrapExcludeNumeric,
                                    ))
                        val availableWidth = max(0f, rect.width() - padding * 2)
                        val wrap =
                            if (!wrap0 && uniformTextSize && availableWidth > 0f) {
                                // When text size is uniform per column we avoid per-cell shrinking; force wrap on overflow.
                                measureMaxLineWidthPx(text, textPaint) > availableWidth * 1.01f
                            } else {
                                wrap0
                            }
                        drawTextInRect(
                            canvas = canvas,
                            paint = textPaint,
                            text = text,
                            rect = rect,
                            padding = padding,
                            alignH = alignH,
                            alignV = alignV,
                            wrap = wrap,
                            uniformTextSize = uniformTextSize,
                        )
                    }
                }
            }
        }
    }

    private fun drawTextInRect(
        canvas: Canvas,
        paint: Paint,
        text: String,
        rect: RectF,
        padding: Float,
        alignH: HorizontalAlignment,
        alignV: VerticalAlignment,
        wrap: Boolean,
        uniformTextSize: Boolean,
    ) {
        val availableWidth = max(0f, rect.width() - padding * 2)
        val availableHeight = max(0f, rect.height() - padding * 2)

        fun layoutLines(): List<String> {
            // Newlines in Excel cell text should always produce new visual lines.
            val paragraphs = text.split("\r\n", "\n", "\r")
            if (paragraphs.isEmpty()) return emptyList()

            val out = ArrayList<String>()
            for (para in paragraphs) {
                if (!wrap || availableWidth <= 0f) {
                    out.add(para)
                } else {
                    out.addAll(breakSingleLine(para, paint, availableWidth))
                }
            }
            return out
        }

        // First pass layout.
        val originalTextSize = paint.textSize
        var lines = layoutLines()

        if (!uniformTextSize) {
            val minSizeWidth = max(8f, originalTextSize * 0.80f)
            val minSizeHeight = max(8f, originalTextSize * 0.70f)

            // When not wrapping, try to shrink to fit width to avoid truncation (e.g., long numbers).
            if (!wrap && availableWidth > 0f && lines.isNotEmpty()) {
                var maxLineWidth = lines.maxOf { paint.measureText(it) }
                var tries = 0
                while (maxLineWidth > availableWidth && paint.textSize > minSizeWidth && tries < 8) {
                    paint.textSize *= 0.9f
                    maxLineWidth = lines.maxOf { paint.measureText(it) }
                    tries++
                }
            }

            // If wrapped/multiline text doesn't fit vertically, shrink a bit to avoid overlap.
            if (availableHeight > 0f) {
                fun calcTotalHeight(lineCount: Int): Float {
                    if (lineCount <= 1) {
                        val fm = paint.fontMetrics
                        return fm.descent - fm.ascent
                    }
                    return lineCount * paint.fontSpacing
                }

                var totalTextHeight = calcTotalHeight(lines.size)
                var tries = 0
                while (totalTextHeight > availableHeight && paint.textSize > minSizeHeight && tries < 8) {
                    paint.textSize *= 0.9f
                    lines = layoutLines()
                    totalTextHeight = calcTotalHeight(lines.size)
                    tries++
                }
            }
        }

        val fm = paint.fontMetrics
        val lineHeight = if (lines.size <= 1) (fm.descent - fm.ascent) else paint.fontSpacing
        val totalTextHeight = lines.size * lineHeight

        val startY = when (alignV) {
            VerticalAlignment.TOP -> rect.top + padding - fm.ascent
            VerticalAlignment.BOTTOM -> rect.bottom - padding - totalTextHeight - fm.ascent
            else -> rect.centerY() - totalTextHeight / 2f - fm.ascent
        }

        // Clip to cell bounds so text never bleeds into other cells/rows.
        val save = canvas.save()
        canvas.clipRect(rect)
        try {
            for ((i, line) in lines.withIndex()) {
                val y = startY + i * lineHeight
                val lineWidth = paint.measureText(line)
                val x = when (alignH) {
                    HorizontalAlignment.RIGHT -> rect.right - padding - lineWidth
                    HorizontalAlignment.CENTER -> rect.centerX() - lineWidth / 2f
                    else -> rect.left + padding
                }
                canvas.drawText(line, x, y, paint)
            }
        } finally {
            canvas.restoreToCount(save)
        }
    }

    private fun breakSingleLine(text: String, paint: Paint, maxWidth: Float): List<String> {
        if (text.isEmpty()) return listOf("")
        if (maxWidth <= 0f) return listOf(text)

        val out = ArrayList<String>()
        var start = 0
        while (start < text.length) {
            var count = paint.breakText(text, start, text.length, true, maxWidth, null)
            if (count <= 0) break
            val eps = 1.01f
            val remainingStart0 = start + count
            if (remainingStart0 < text.length) {
                // If the remaining part fits in ONE line, we can finish this paragraph in 2 lines.
                // In that case, prefer a more balanced split instead of an almost-full first line
                // and a tiny second line (e.g. 13+2).
                val remainingWidth0 = paint.measureText(text, remainingStart0, text.length)
                if (remainingWidth0 <= maxWidth * eps) {
                    val total = text.length - start
                    var desiredFirst = (total + 1) / 2 // prefer first >= second
                    desiredFirst = desiredFirst.coerceIn(1, count)

                    // Avoid the very last line being a single orphan char when possible.
                    if (total - desiredFirst == 1 && desiredFirst > 1) {
                        val cand = desiredFirst - 1
                        val tailWidth = paint.measureText(text, start + cand, text.length)
                        if (tailWidth <= maxWidth * eps) {
                            desiredFirst = cand
                        }
                    }

                    // Ensure the tail still fits in one line; otherwise move split back to the right.
                    while (desiredFirst < count) {
                        val tailWidth = paint.measureText(text, start + desiredFirst, text.length)
                        if (tailWidth <= maxWidth * eps) break
                        desiredFirst += 1
                    }

                    count = desiredFirst
                }
            }
            out.add(text.substring(start, start + count))
            start += count
        }
        // Safety: avoid returning empty on weird measurement results.
        return if (out.isNotEmpty()) out else listOf(text)
    }

    private fun measureMaxLineWidthPx(text: String, paint: Paint): Float {
        if (text.isEmpty()) return 0f
        val lines = text.split("\r\n", "\n", "\r")
        var maxWidth = 0f
        for (line in lines) {
            maxWidth = max(maxWidth, paint.measureText(line))
        }
        return maxWidth
    }

    private fun looksNumeric(text: String): Boolean {
        val t = text.trim()
        if (t.isEmpty()) return false
        for (ch in t) {
            if (ch.isDigit()) continue
            when (ch) {
                '.', ',', '-', '+', '/', '%', '(', ')', ' ', '\u00A0' -> continue
                else -> return false
            }
        }
        return true
    }

    private fun shouldAutoWrapText(
        text: String,
        paint: Paint,
        availableWidth: Float,
        minTextLength: Int,
        excludeNumeric: Boolean,
    ): Boolean {
        if (availableWidth <= 0f) return false
        val t = text.trim()
        if (t.isEmpty()) return false

        val width = measureMaxLineWidthPx(t, paint)
        if (width <= availableWidth * 1.05f) return false

        // If keeping one-line would require shrinking too much, prefer wrapping even if Excel didn't.
        val requiredScale = if (width > 0f) availableWidth / width else 1f
        if (requiredScale < 0.80f) return true

        if (t.length < minTextLength) return false
        if (excludeNumeric && looksNumeric(t)) return false
        return true
    }

    private fun applyFont(
        paint: Paint,
        font: Font,
        scale: Float,
        minPt: Int = 8,
        maxPt: Int = 28,
        overridePt: Int? = null,
    ) {
        val pt = (overridePt ?: font.fontHeightInPoints.toInt()).coerceIn(minPt, maxPt)
        val basePx = max(10f, pt * 4f / 3f)
        paint.textSize = basePx * scale
        val style = when {
            font.bold && font.italic -> Typeface.BOLD_ITALIC
            font.bold -> Typeface.BOLD
            font.italic -> Typeface.ITALIC
            else -> Typeface.NORMAL
        }
        paint.typeface = androidTypefaceFor(font.fontName, style)
        // Avoid resolving XSSF font colors: XSSFColor references java.awt which isn't on Android.
        paint.color = Color.BLACK
    }

    private fun androidTypefaceFor(fontName: String?, style: Int): Typeface {
        val name = (fontName ?: "").trim()
        if (name.isEmpty()) return Typeface.create(Typeface.DEFAULT, style)

        val lower = name.lowercase()
        val family = when {
            lower.contains("courier") || lower.contains("consolas") || lower.contains("mono") -> Typeface.MONOSPACE
            lower.contains("times") || lower.contains("serif") || lower.contains("roman") -> Typeface.SERIF
            // Common Chinese fonts on Windows/macOS.
            lower.contains("yahei") || name.contains("微软雅黑") -> Typeface.SANS_SERIF
            lower.contains("simsun") || name.contains("宋体") -> Typeface.SERIF
            lower.contains("simhei") || name.contains("黑体") -> Typeface.SANS_SERIF
            else -> Typeface.SANS_SERIF
        }
        return Typeface.create(family, style)
    }

    private fun computeMaxDigitWidthPx(workbook: Workbook): Float {
        val paint = Paint(Paint.ANTI_ALIAS_FLAG)
        val font = runCatching { workbook.getFontAt(0) }.getOrNull()
        if (font != null) {
            applyFont(paint, font, 1f)
        } else {
            paint.textSize = 14f
            paint.typeface = Typeface.DEFAULT
        }

        val mdw = paint.measureText("0")
        return if (mdw.isFinite() && mdw > 0f) mdw else 7f
    }

    @Suppress("unused")
    private fun isCellVisuallyUsed(cell: Cell): Boolean {
        // Heuristic used for future "auto-trim" improvements.
        return when (cell.cellType) {
            CellType.STRING -> cell.stringCellValue?.isNotBlank() == true
            CellType.BLANK -> {
                val style = cell.cellStyle
                style.fillPattern != FillPatternType.NO_FILL ||
                    style.borderTop != BorderStyle.NONE ||
                    style.borderRight != BorderStyle.NONE ||
                    style.borderBottom != BorderStyle.NONE ||
                    style.borderLeft != BorderStyle.NONE
            }
            else -> true
        }
    }

    private fun backgroundColorArgb(style: org.apache.poi.ss.usermodel.CellStyle): Int? {
        if (style.fillPattern == FillPatternType.NO_FILL) return null
        // Only treat as background when it's actually solid.
        if (style.fillPattern != FillPatternType.SOLID_FOREGROUND) return null

        // Use indexed colors only. Custom XSSF colors require XSSFColor which references java.awt.
        val idx = style.fillForegroundColor.toInt() and 0xFF
        return indexedColorToArgb(idx)
    }

    private fun indexedColorToArgb(idx: Int): Int? {
        // HSSF/XSSF share the same indices for the basic palette.
        // We only map the common ones; unknown indices fall back to "no fill".
        return when (idx) {
            8 -> Color.BLACK
            9 -> Color.WHITE
            10 -> Color.RED
            11 -> Color.rgb(0, 255, 0) // bright green
            12 -> Color.BLUE
            13 -> Color.YELLOW
            14 -> Color.MAGENTA
            15 -> Color.CYAN
            22 -> Color.rgb(192, 192, 192) // grey 25%
            23 -> Color.rgb(128, 128, 128) // grey 50%
            else -> null
        }
    }

    private fun cellKey(row: Int, col: Int): Long {
        return (row.toLong() shl 32) or (col.toLong() and 0xFFFF_FFFFL)
    }
}
