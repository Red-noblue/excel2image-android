package com.zys.excel2image

import android.graphics.Bitmap
import android.graphics.BitmapFactory
import android.graphics.Canvas
import android.graphics.Paint
import android.graphics.Path
import android.graphics.PointF
import android.net.Uri
import android.os.Bundle
import android.widget.AdapterView
import android.widget.ArrayAdapter
import androidx.appcompat.app.AlertDialog
import androidx.appcompat.app.AppCompatActivity
import androidx.core.view.isVisible
import androidx.lifecycle.lifecycleScope
import com.davemorrissey.labs.subscaleview.ImageSource
import com.zys.excel2image.databinding.ActivityAnnotateBinding
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import java.io.File
import java.io.FileOutputStream

class AnnotateActivity : AppCompatActivity() {

    companion object {
        const val EXTRA_EXCEL_URI = "excel_uri"
        const val EXTRA_SHEET_INDEX = "sheet_index"
        const val EXTRA_BASE_NAME = "base_name"
    }

    private lateinit var binding: ActivityAnnotateBinding

    private data class Part(
        val index: Int,
        val file: File,
        val strokes: MutableList<Stroke> = mutableListOf(),
    )

    private var excelUri: Uri? = null
    private var sheetIndex: Int = 0
    private var baseName: String = "workbook"

    private var sheetName: String = "sheet"

    private val parts = mutableListOf<Part>()
    private var currentPartIndex = 0
    private var drawMode = false

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        binding = ActivityAnnotateBinding.inflate(layoutInflater)
        setContentView(binding.root)

        excelUri = intent.getStringExtra(EXTRA_EXCEL_URI)?.let { Uri.parse(it) }
        sheetIndex = intent.getIntExtra(EXTRA_SHEET_INDEX, 0)
        baseName = intent.getStringExtra(EXTRA_BASE_NAME) ?: "workbook"

        val uri = excelUri
        if (uri == null) {
            finish()
            return
        }

        binding.overlay.imageView = binding.imgLarge
        binding.overlay.onStrokeFinished = { stroke ->
            parts.getOrNull(currentPartIndex)?.strokes?.add(stroke)
            refreshOverlay()
            updateButtons()
        }
        binding.imgLarge.setOnStateChangedListener(object : com.davemorrissey.labs.subscaleview.SubsamplingScaleImageView.OnStateChangedListener {
            override fun onScaleChanged(newScale: Float, origin: Int) {
                binding.overlay.invalidate()
            }

            override fun onCenterChanged(newCenter: PointF, origin: Int) {
                binding.overlay.invalidate()
            }
        })

        binding.btnMode.setOnClickListener {
            drawMode = !drawMode
            applyModeUi()
        }

        binding.btnUndo.setOnClickListener {
            val list = parts.getOrNull(currentPartIndex)?.strokes ?: return@setOnClickListener
            if (list.isNotEmpty()) {
                list.removeAt(list.lastIndex)
                refreshOverlay()
                updateButtons()
            }
        }

        binding.btnClear.setOnClickListener {
            val list = parts.getOrNull(currentPartIndex)?.strokes ?: return@setOnClickListener
            if (list.isEmpty()) return@setOnClickListener
            AlertDialog.Builder(this)
                .setTitle("清空标注？")
                .setMessage("将清空当前段的所有涂鸦标注。")
                .setPositiveButton("清空") { _, _ ->
                    list.clear()
                    refreshOverlay()
                    updateButtons()
                }
                .setNegativeButton("取消", null)
                .show()
        }

        binding.btnExportImage.setOnClickListener { exportAnnotatedImages() }
        binding.btnExportPdf.setOnClickListener { exportAnnotatedPdf() }

        binding.spinnerParts.onItemSelectedListener = object : AdapterView.OnItemSelectedListener {
            override fun onItemSelected(parent: AdapterView<*>?, view: android.view.View?, position: Int, id: Long) {
                if (position != currentPartIndex) {
                    currentPartIndex = position
                    showPart(position)
                }
            }

            override fun onNothingSelected(parent: AdapterView<*>?) = Unit
        }

        applyModeUi()
        updateButtons()

        renderPartsToCache(uri)
    }

    override fun onDestroy() {
        super.onDestroy()
        // Best-effort cleanup of cached preview images.
        for (p in parts) {
            runCatching { p.file.delete() }
        }
    }

    private fun applyModeUi() {
        binding.overlay.drawEnabled = drawMode
        binding.btnMode.text = if (drawMode) "涂鸦" else "移动"
        binding.txtStatus.text = buildString {
            append(sheetName)
            if (parts.size > 1) append("  第${currentPartIndex + 1}/${parts.size}段")
            append(if (drawMode) "（涂鸦模式）" else "（移动模式）")
        }
    }

    private fun refreshOverlay() {
        binding.overlay.strokes = parts.getOrNull(currentPartIndex)?.strokes ?: emptyList()
    }

    private fun updateButtons() {
        val hasParts = parts.isNotEmpty()
        val strokes = parts.getOrNull(currentPartIndex)?.strokes.orEmpty()
        binding.btnUndo.isEnabled = hasParts && strokes.isNotEmpty() && !binding.progress.isVisible
        binding.btnClear.isEnabled = hasParts && strokes.isNotEmpty() && !binding.progress.isVisible
        binding.btnExportImage.isEnabled = hasParts && !binding.progress.isVisible
        binding.btnExportPdf.isEnabled = hasParts && !binding.progress.isVisible
        binding.spinnerParts.isEnabled = parts.size > 1 && !binding.progress.isVisible
        binding.btnMode.isEnabled = hasParts && !binding.progress.isVisible
    }

    private fun renderPartsToCache(uri: Uri) {
        binding.progress.isVisible = true
        binding.txtStatus.text = "正在生成可标注预览…"
        updateButtons()

        lifecycleScope.launch {
            val result = withContext(Dispatchers.Default) {
                runCatching {
                    val loaded = ExcelLoader.load(contentResolver, uri)
                    val safeSheetIndex = sheetIndex.coerceIn(0, loaded.workbook.numberOfSheets - 1)
                    sheetIndex = safeSheetIndex
                    sheetName = loaded.workbook.getSheetAt(safeSheetIndex).sheetName

                    val outDir = File(cacheDir, "annotate").apply { mkdirs() }
                    // Clear old cache.
                    outDir.listFiles()?.forEach { runCatching { it.delete() } }

                    val localParts = mutableListOf<Part>()

                    ExcelBitmapRenderer.renderSheetParts(
                        workbook = loaded.workbook,
                        sheetIndex = safeSheetIndex,
                        options = annotateRenderOptions(),
                    ) { partIndex, partCount, bmp ->
                        val name = buildString {
                            append("sheet_")
                            append(partIndex.toString().padStart(2, '0'))
                            append("_of_")
                            append(partCount.toString().padStart(2, '0'))
                            append(".png")
                        }
                        val file = File(outDir, name)
                        FileOutputStream(file).use { out ->
                            if (!bmp.compress(Bitmap.CompressFormat.PNG, 100, out)) {
                                error("Bitmap compress failed")
                            }
                        }
                        localParts += Part(index = partIndex, file = file)
                    }

                    loaded.workbook.closeQuietly()
                    localParts.sortBy { it.index }
                    localParts
                }
            }

            binding.progress.isVisible = false

            result.onSuccess { localParts ->
                parts.clear()
                parts.addAll(localParts)

                if (parts.isEmpty()) {
                    binding.txtStatus.text = "无可标注内容"
                    updateButtons()
                    return@launch
                }

                val labels = parts.map { p -> "第${p.index + 1}段" }
                binding.spinnerParts.adapter = ArrayAdapter(
                    this@AnnotateActivity,
                    android.R.layout.simple_spinner_dropdown_item,
                    labels,
                )
                binding.spinnerParts.isEnabled = parts.size > 1
                currentPartIndex = 0
                showPart(0)
                applyModeUi()
                updateButtons()
            }.onFailure { e ->
                binding.txtStatus.text = "生成失败：${e.message ?: e.javaClass.simpleName}"
                AlertDialog.Builder(this@AnnotateActivity)
                    .setTitle("生成失败")
                    .setMessage((e.stackTraceToString()).take(10_000))
                    .setPositiveButton("关闭") { _, _ -> finish() }
                    .show()
                updateButtons()
            }
        }
    }

    private fun showPart(index: Int) {
        val part = parts.getOrNull(index) ?: return
        binding.imgLarge.setImage(ImageSource.uri(Uri.fromFile(part.file)))
        refreshOverlay()
        applyModeUi()
        updateButtons()
    }

    private fun exportAnnotatedImages() {
        if (parts.isEmpty()) return

        binding.progress.isVisible = true
        binding.txtStatus.text = "正在导出标注图片…"
        updateButtons()

        lifecycleScope.launch {
            val result = withContext(Dispatchers.Default) {
                runCatching {
                    val out = ArrayList<Uri>()
                    for (part in parts) {
                        val bmp = decodeMutable(part.file)
                        try {
                            val canvas = Canvas(bmp)
                            drawStrokesOnCanvas(canvas, part.strokes, scale = 1f)

                            val displayName = buildString {
                                append(baseName)
                                append("_")
                                append(sanitizeForFileName(sheetName))
                                append("_marked")
                                if (parts.size > 1) {
                                    append("_p")
                                    append((part.index + 1).toString().padStart(2, '0'))
                                }
                                append(".png")
                            }
                            out += ImageSaver.savePngToPictures(this@AnnotateActivity, displayName, bmp)
                        } finally {
                            bmp.recycle()
                        }
                    }
                    out
                }
            }

            binding.progress.isVisible = false

            result.onSuccess { uris ->
                binding.txtStatus.text = "导出完成：已保存 ${uris.size} 张标注图片"
            }.onFailure { e ->
                binding.txtStatus.text = "导出失败：${e.message ?: e.javaClass.simpleName}"
                AlertDialog.Builder(this@AnnotateActivity)
                    .setTitle("导出失败")
                    .setMessage((e.stackTraceToString()).take(10_000))
                    .setPositiveButton("关闭", null)
                    .show()
            }
            updateButtons()
        }
    }

    private fun exportAnnotatedPdf() {
        if (parts.isEmpty()) return

        binding.progress.isVisible = true
        binding.txtStatus.text = "正在导出标注PDF…"
        updateButtons()

        lifecycleScope.launch {
            val result = withContext(Dispatchers.Default) {
                runCatching {
                    val uri = excelUri ?: error("Missing excel uri")
                    val displayName = buildString {
                        append(baseName)
                        append("_")
                        append(sanitizeForFileName(sheetName))
                        append("_marked.pdf")
                    }

                    DocumentSaver.savePdfToDownloads(this@AnnotateActivity, displayName) { out ->
                        // Re-render the sheet directly into PDF (vector-like), then overlay doodles.
                        // This avoids embedding huge full-resolution bitmaps into the PDF.
                        val loaded = ExcelLoader.load(contentResolver, uri)
                        try {
                            ExcelBitmapRenderer.writeSheetPdf(
                                workbook = loaded.workbook,
                                sheetIndex = sheetIndex.coerceIn(0, loaded.workbook.numberOfSheets - 1),
                                options = annotateRenderOptions(),
                                out = out,
                            ) { partIndex, _, canvas ->
                                val strokeList = parts.firstOrNull { it.index == partIndex }?.strokes.orEmpty()
                                drawStrokesOnCanvas(canvas, strokeList, scale = 1f)
                            }
                        } finally {
                            loaded.workbook.closeQuietly()
                        }
                    }
                }
            }

            binding.progress.isVisible = false

            result.onSuccess {
                binding.txtStatus.text = "导出完成：已保存 1 个标注PDF"
            }.onFailure { e ->
                binding.txtStatus.text = "导出失败：${e.message ?: e.javaClass.simpleName}"
                AlertDialog.Builder(this@AnnotateActivity)
                    .setTitle("导出失败")
                    .setMessage((e.stackTraceToString()).take(10_000))
                    .setPositiveButton("关闭", null)
                    .show()
            }

            updateButtons()
        }
    }

    private fun drawStrokesOnCanvas(canvas: Canvas, strokes: List<Stroke>, scale: Float) {
        if (strokes.isEmpty()) return

        val paint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
            style = Paint.Style.STROKE
            strokeJoin = Paint.Join.ROUND
            strokeCap = Paint.Cap.ROUND
        }
        val path = Path()
        for (s in strokes) {
            if (s.points.size < 2) continue
            paint.color = s.colorArgb
            paint.strokeWidth = (s.widthSourcePx * scale).coerceAtLeast(1f)

            path.reset()
            val first = s.points.first()
            path.moveTo(first.x * scale, first.y * scale)
            for (i in 1 until s.points.size) {
                val p = s.points[i]
                path.lineTo(p.x * scale, p.y * scale)
            }
            canvas.drawPath(path, paint)
        }
    }

    private fun decodeMutable(file: File): Bitmap {
        val opts = BitmapFactory.Options().apply {
            inPreferredConfig = Bitmap.Config.ARGB_8888
            inMutable = true
        }
        val bmp = BitmapFactory.decodeFile(file.absolutePath, opts)
            ?: error("Decode failed: ${file.name}")
        return if (bmp.isMutable) bmp else bmp.copy(Bitmap.Config.ARGB_8888, true)
    }

    private fun annotateRenderOptions(): RenderOptions {
        // Export-like scale for clear markup. (SSIV will keep it smooth with tiling.)
        // NOTE: Keep this consistent across:
        // - cached part images (for drawing)
        // - annotated PDF export (for exact coordinate match)
        return RenderOptions(
            scale = 2.0f,
            maxBitmapDimension = 16_000,
            maxTotalPixels = 20_000_000L,
            uniformFontPerColumn = true,
            trimMaxCells = 250_000,
            columnWidthMaxCells = 250_000,
            columnFontMaxCells = 250_000,
            autoFitMaxCells = 250_000,
            maxAutoRowHeightPx = 2000,
            minFontPt = 8,
            maxFontPt = 20,
        )
    }

    private fun sanitizeForFileName(name: String): String {
        return name.replace(Regex("""[\\/:*?"<>|]"""), "_").trim().ifBlank { "sheet" }
    }
}
