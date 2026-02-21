package com.zys.excel2image

import android.graphics.Bitmap
import android.graphics.Canvas
import android.graphics.Color
import android.graphics.Paint
import android.graphics.RectF
import android.graphics.Typeface
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFFont
import kotlin.math.ceil
import kotlin.math.max
import kotlin.math.min

data class RenderOptions(
    val scale: Float,
    val maxBitmapDimension: Int = 16_000,
    val maxTotalPixels: Long = 100_000_000L,
)

data class RenderResult(
    val bitmaps: List<Bitmap>,
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

        val used = findPrintAreaRange(workbook, sheetIndex) ?: findUsedRange(sheet)
        if (used == null) {
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

        val (firstRow, lastRow, firstCol, lastCol) = used

        val baseColWidthsPx = IntArray(lastCol - firstCol + 1) { idx ->
            val col = firstCol + idx
            // Excel column width is in 1/256 character units. This is an approximation that works
            // well enough for "shareable images" without pulling AWT font metrics.
            val widthChars = sheet.getColumnWidth(col) / 256f
            max(12, (widthChars * 7f + 5f).toInt())
        }

        val baseRowHeightsPx = IntArray(lastRow - firstRow + 1) { idx ->
            val rowNum = firstRow + idx
            val row = sheet.getRow(rowNum)
            val htPt = row?.heightInPoints ?: sheet.defaultRowHeightInPoints
            max(16, ceil(htPt * 4f / 3f).toInt())
        }

        // Apply initial scale. We'll shrink further if needed to respect bitmap constraints.
        var scale = options.scale.coerceAtLeast(0.1f)

        fun scaledSum(arr: IntArray): Int = arr.sumOf { (it * scale).toInt() }

        var width = scaledSum(baseColWidthsPx)
        var height = scaledSum(baseRowHeightsPx)

        if (width > options.maxBitmapDimension) {
            val shrink = options.maxBitmapDimension.toFloat() / width.toFloat()
            scale *= shrink
            width = scaledSum(baseColWidthsPx)
            height = scaledSum(baseRowHeightsPx)
        }

        val warnings = mutableListOf<String>()
        if (width > options.maxBitmapDimension) {
            warnings += "表格太宽，已强制缩小"
        }

        val mergeInfo = buildMergeInfo(sheet, firstRow, lastRow, firstCol, lastCol)

        val parts = planVerticalParts(
            baseRowHeightsPx = baseRowHeightsPx,
            scaledWidthPx = width,
            scale = scale,
            maxBitmapDimension = options.maxBitmapDimension,
            maxTotalPixels = options.maxTotalPixels,
        )

        val formatter = DataFormatter()
        val evaluator = workbook.creationHelper.createFormulaEvaluator()

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
        val scaledRowHeights = baseRowHeightsPx.map { max(1, (it * scale).toInt()) }
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

        val maxHeightByPixels = max(200, (maxTotalPixels / scaledWidthPx.toLong()).toInt())
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
    ) {
        val (cellToMerge, mergeStarts) = mergeInfo

        val colCount = lastCol - firstCol + 1
        val rowCount = lastRow - firstRow + 1

        val colWidths = IntArray(colCount) { idx -> max(1, (baseColWidthsPx[idx] * scale).toInt()) }
        val rowHeights = IntArray(rowCount) { idx -> max(1, (baseRowHeightsPx[idx] * scale).toInt()) }

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

            val row = sheet.getRow(absRow)

            for (absCol in firstCol..lastCol) {
                val colIdx = absCol - firstCol
                val left = x[colIdx]
                val right = x[colIdx + 1]

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

                    val text = formatter.formatCellValue(cell, evaluator).orEmpty()
                    if (text.isNotBlank()) {
                        val font = workbook.getFontAt(style.fontIndex)
                        applyFont(textPaint, font, scale)
                        val alignH = style.alignment
                        val alignV = style.verticalAlignment
                        val wrap = style.wrapText
                        drawTextInRect(
                            canvas = canvas,
                            paint = textPaint,
                            text = text,
                            rect = rect,
                            padding = padding,
                            alignH = alignH,
                            alignV = alignV,
                            wrap = wrap,
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
    ) {
        val availableWidth = max(0f, rect.width() - padding * 2)
        val lines = if (wrap) breakLines(text, paint, availableWidth) else listOf(text)

        val fm = paint.fontMetrics
        val lineHeight = paint.fontSpacing
        val totalTextHeight = lines.size * lineHeight

        val startY = when (alignV) {
            VerticalAlignment.TOP -> rect.top + padding - fm.ascent
            VerticalAlignment.BOTTOM -> rect.bottom - padding - totalTextHeight - fm.ascent
            else -> rect.centerY() - totalTextHeight / 2f - fm.ascent
        }

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
    }

    private fun breakLines(text: String, paint: Paint, maxWidth: Float): List<String> {
        if (text.isEmpty()) return emptyList()
        if (maxWidth <= 0f) return listOf(text)

        val out = ArrayList<String>()
        var start = 0
        while (start < text.length) {
            val count = paint.breakText(text, start, text.length, true, maxWidth, null)
            if (count <= 0) break
            out.add(text.substring(start, start + count))
            start += count
        }
        return out
    }

    private fun applyFont(paint: Paint, font: Font, scale: Float) {
        val basePx = max(10f, font.fontHeightInPoints * 4f / 3f)
        paint.textSize = basePx * scale
        paint.typeface = when {
            font.bold && font.italic -> Typeface.create(Typeface.DEFAULT, Typeface.BOLD_ITALIC)
            font.bold -> Typeface.create(Typeface.DEFAULT, Typeface.BOLD)
            font.italic -> Typeface.create(Typeface.DEFAULT, Typeface.ITALIC)
            else -> Typeface.DEFAULT
        }

        val xssfFont = font as? XSSFFont
        val argb = xssfFont?.xssfColor?.let(::xssfColorToArgb)
        paint.color = argb ?: Color.BLACK
    }

    private fun backgroundColorArgb(style: org.apache.poi.ss.usermodel.CellStyle): Int? {
        if (style.fillPattern == FillPatternType.NO_FILL) return null
        val xStyle = style as? XSSFCellStyle ?: return null
        val xColor = xStyle.fillForegroundColorColor as? XSSFColor ?: return null
        // Only treat as background when it's actually solid.
        if (xStyle.fillPattern != FillPatternType.SOLID_FOREGROUND) return null
        return xssfColorToArgb(xColor)
    }

    private fun xssfColorToArgb(color: XSSFColor): Int? {
        val argb = color.argb
        if (argb != null && argb.size == 4) {
            val a = argb[0].toInt() and 0xFF
            val r = argb[1].toInt() and 0xFF
            val g = argb[2].toInt() and 0xFF
            val b = argb[3].toInt() and 0xFF
            return Color.argb(a, r, g, b)
        }
        val rgb = color.rgb
        if (rgb != null && rgb.size == 3) {
            val r = rgb[0].toInt() and 0xFF
            val g = rgb[1].toInt() and 0xFF
            val b = rgb[2].toInt() and 0xFF
            return Color.rgb(r, g, b)
        }
        return null
    }

    private fun cellKey(row: Int, col: Int): Long {
        return (row.toLong() shl 32) or (col.toLong() and 0xFFFF_FFFFL)
    }
}
