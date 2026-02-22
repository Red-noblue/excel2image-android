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
    val columnWidthSampleMaxTextLength: Int = 30,
    val minColumnWidthPx: Int = 12,
    val emptyColumnWidthPx: Int = 24,
    val maxColumnWidthPx: Int = 800,
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

        // Apply initial scale. We'll shrink further if needed to respect bitmap constraints.
        var scale = options.scale.coerceAtLeast(0.1f)

        fun scaledSum(arr: IntArray): Int = arr.sumOf { (it * scale).toInt() }

        var width = scaledSum(baseColWidthsPx)

        if (width > options.maxBitmapDimension) {
            val shrink = options.maxBitmapDimension.toFloat() / width.toFloat()
            scale *= shrink
            width = scaledSum(baseColWidthsPx)
        }

        if (width > options.maxBitmapDimension) {
            warnings += "表格太宽，已强制缩小"
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

    fun renderSheetParts(
        workbook: Workbook,
        sheetIndex: Int,
        options: RenderOptions,
        recycleAfterCallback: Boolean = true,
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

        var scale = options.scale.coerceAtLeast(0.1f)

        fun scaledSum(arr: IntArray): Int = arr.sumOf { (it * scale).toInt() }

        var width = scaledSum(baseColWidthsPx)

        if (width > options.maxBitmapDimension) {
            val shrink = options.maxBitmapDimension.toFloat() / width.toFloat()
            scale *= shrink
            width = scaledSum(baseColWidthsPx)
        }

        if (width > options.maxBitmapDimension) {
            warnings += "表格太宽，已强制缩小"
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

        for ((i, part) in parts.withIndex()) {
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

        // Apply initial scale. We'll shrink further if needed to respect page constraints.
        var scale = options.scale.coerceAtLeast(0.1f)

        fun scaledSum(arr: IntArray): Int = arr.sumOf { (it * scale).toInt() }

        var width = scaledSum(baseColWidthsPx)

        if (width > options.maxBitmapDimension) {
            val shrink = options.maxBitmapDimension.toFloat() / width.toFloat()
            scale *= shrink
            width = scaledSum(baseColWidthsPx)
        }

        if (width > options.maxBitmapDimension) {
            warnings += "表格太宽，已强制缩小"
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

        val rowIt = sheet.rowIterator()
        while (rowIt.hasNext()) {
            val row = rowIt.next()
            val r = row.rowNum
            if (r < firstRow || r > lastRow) continue
            if (row.zeroHeight) continue

            val cellIt = row.cellIterator()
            while (cellIt.hasNext()) {
                val cell = cellIt.next()
                val c = cell.columnIndex
                if (c < firstCol || c > lastCol) continue
                if (sheet.isColumnHidden(c)) continue

                val text =
                    runCatching { formatter.formatCellValue(cell, evaluator) }.getOrNull().orEmpty().trim()
                if (text.isEmpty()) continue

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
        val sampleCounts = IntArray(colCount)
        val hasAnyText = BooleanArray(colCount)
        val minFromMergedSpans = IntArray(colCount)

        val paint = Paint(Paint.ANTI_ALIAS_FLAG)
        val padding = 4f

        var processed = 0

        val rowIt = sheet.rowIterator()
        while (rowIt.hasNext()) {
            val row = rowIt.next()
            val r = row.rowNum
            if (r < firstRow || r > lastRow) continue
            if (row.zeroHeight) continue

            val cellIt = row.cellIterator()
            while (cellIt.hasNext()) {
                val cell = cellIt.next()
                val c = cell.columnIndex
                if (c < firstCol || c > lastCol) continue
                if (sheet.isColumnHidden(c)) continue

                val text =
                    runCatching { formatter.formatCellValue(cell, evaluator) }.getOrNull().orEmpty().trim()
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

                    val measured = measureMaxLineWidthPx(text, paint)
                    val measuredWithPadding = ceil(measured + padding * 2).toInt()

                    // Distribute merged-cell width across its spanned columns (best-effort).
                    val span = (spanLastCol - spanFirstCol + 1).coerceAtLeast(1)
                    val perCol = ceil(measuredWithPadding.toFloat() / span.toFloat()).toInt()
                    for (absCol in spanFirstCol..spanLastCol) {
                        val idx = absCol - firstCol
                        if (idx in 0 until colCount) {
                            hasAnyText[idx] = true
                            minFromMergedSpans[idx] = max(minFromMergedSpans[idx], perCol)
                        }
                    }
                    continue
                }

                val colIdx = c - firstCol
                if (colIdx !in 0 until colCount) continue
                hasAnyText[colIdx] = true

                // Ignore extremely long texts when deciding column width; those will be wrapped.
                if (text.length > sampleMaxTextLength) continue

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

                val measured = measureMaxLineWidthPx(text, paint)
                val measuredWithPadding = ceil(measured + padding * 2).toInt()

                val count = sampleCounts[colIdx]
                if (count < maxSamplesPerCol) {
                    samples[colIdx][count] = measuredWithPadding
                    sampleCounts[colIdx] = count + 1
                }
            }
        }

        val out = baseColWidthsPx.clone()
        var changed = false

        for (i in 0 until colCount) {
            val base = baseColWidthsPx[i]
            if (base <= 0) {
                out[i] = 0
                continue
            }

            val newWidth = if (!hasAnyText[i]) {
                min(base, emptyColumnWidthPx).coerceAtLeast(minColumnWidthPx)
            } else {
                val count = sampleCounts[i]
                val typical = if (count <= 0) {
                    // Fallback for "all texts are long": keep it moderate, don't force huge columns.
                    min(base, 220).coerceAtLeast(minColumnWidthPx)
                } else {
                    val arr = samples[i].copyOf(count)
                    arr.sort()
                    val idx = ((count - 1) * 0.8f).toInt().coerceIn(0, count - 1)
                    arr[idx].coerceAtLeast(minColumnWidthPx)
                }

                val withMerged = max(typical, minFromMergedSpans[i])
                min(base, withMerged).coerceIn(minColumnWidthPx, maxColumnWidthPx)
            }

            if (newWidth != base) {
                out[i] = newWidth
                changed = true
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

        val padding = 4f
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

                val hasNewline = rawText.contains('\n') || rawText.contains('\r')
                val trimmedLen = rawText.trim().length
                val isWrapCandidate =
                    style.wrapText ||
                        hasNewline ||
                        (autoWrapOverflowText && trimmedLen >= autoWrapMinTextLength) ||
                        // When font size is uniform per column, even shorter texts may overflow and need wrapping.
                        (columnFontPts != null && trimmedLen >= 4)
                if (!isWrapCandidate) continue

                processed++
                if (processed > maxCells) {
                    return AutoFitOutcome(
                        rowHeightsPx = out,
                        didChange = changed,
                        warning = "表格较大，已限制自适应行高的扫描范围",
                    )
                }

                val key = cellKey(r, c)
                val merge = cellToMerge[key]
                if (merge != null && key !in mergeStarts) {
                    continue
                }

                // Only handle single-row merged cells for now (common in headers).
                if (merge != null && merge.firstRow != merge.lastRow) {
                    continue
                }

                val spanFirstCol = merge?.firstCol ?: c
                val spanLastCol = merge?.lastCol ?: c

                var widthPx = 0
                for (absCol in spanFirstCol..spanLastCol) {
                    val idx = absCol - firstCol
                    if (idx !in baseColWidthsPx.indices) continue
                    widthPx += baseColWidthsPx[idx]
                }
                if (widthPx <= 0) continue

                val availableWidth = max(0f, widthPx.toFloat() - padding * 2)
                val font = workbook.getFontAt(style.fontIndex)
                val colIdxForFont = spanFirstCol - firstCol
                val overridePt = columnFontPts?.getOrNull(colIdxForFont)
                applyFont(paint, font, 1f, minFontPt, maxFontPt, overridePt = overridePt)

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

                val capped = min(required, maxAutoRowHeightPx)
                if (capped > out[rowIdx]) {
                    out[rowIdx] = capped
                    changed = true
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

        val scaledRowHeights = baseRowHeightsPx.map { h ->
            if (h <= 0) 0 else max(1, (h * scale).toInt())
        }
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

        val colWidths = IntArray(colCount) { idx ->
            val base = baseColWidthsPx[idx]
            if (base <= 0) 0 else max(1, (base * scale).toInt())
        }
        val rowHeights = IntArray(rowCount) { idx ->
            val base = baseRowHeightsPx[idx]
            if (base <= 0) 0 else max(1, (base * scale).toInt())
        }

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
            val count = paint.breakText(text, start, text.length, true, maxWidth, null)
            if (count <= 0) break
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
