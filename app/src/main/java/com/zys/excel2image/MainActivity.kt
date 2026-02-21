package com.zys.excel2image

import android.Manifest
import android.content.ClipData
import android.content.ClipboardManager
import android.content.Context
import android.content.Intent
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.view.View
import android.widget.AdapterView
import android.widget.ArrayAdapter
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.appcompat.app.AlertDialog
import androidx.core.view.isVisible
import androidx.lifecycle.lifecycleScope
import com.google.android.material.snackbar.Snackbar
import com.zys.excel2image.databinding.ActivityMainBinding
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.Job
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.apache.poi.ss.usermodel.Workbook

class MainActivity : AppCompatActivity() {

    private lateinit var binding: ActivityMainBinding

    private var workbook: Workbook? = null
    private var workbookName: String = "workbook"
    private var currentSheetIndex: Int = 0
    private var lastExportUris: List<Uri> = emptyList()
    private var previewBitmap: android.graphics.Bitmap? = null

    private var renderJob: Job? = null

    private val openDocumentLauncher =
        registerForActivityResult(ActivityResultContracts.OpenDocument()) { uri ->
            if (uri != null) {
                // Persist permission so user can export later without the original grant.
                try {
                    contentResolver.takePersistableUriPermission(
                        uri,
                        Intent.FLAG_GRANT_READ_URI_PERMISSION,
                    )
                } catch (_: SecurityException) {
                    // Ignore: not all URIs are persistable.
                }
                loadWorkbookFromUri(uri)
            }
        }

    private val writeStoragePermissionLauncher =
        registerForActivityResult(ActivityResultContracts.RequestPermission()) { granted ->
            if (granted) {
                exportCurrentSheet()
            } else {
                showMessage("需要存储权限才能保存到相册（Android 9 及以下）。")
            }
        }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        binding = ActivityMainBinding.inflate(layoutInflater)
        setContentView(binding.root)

        binding.btnOpen.setOnClickListener {
            openDocumentLauncher.launch(
                arrayOf(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "application/vnd.ms-excel",
                    "text/csv",
                    "*/*",
                ),
            )
        }

        binding.btnExport.setOnClickListener {
            if (needsLegacyWritePermission()) {
                writeStoragePermissionLauncher.launch(Manifest.permission.WRITE_EXTERNAL_STORAGE)
            } else {
                exportCurrentSheet()
            }
        }

        binding.btnShare.setOnClickListener {
            shareLastExport()
        }

        binding.spinnerSheets.onItemSelectedListener = object : AdapterView.OnItemSelectedListener {
            override fun onItemSelected(
                parent: AdapterView<*>?,
                view: View?,
                position: Int,
                id: Long,
            ) {
                if (position != currentSheetIndex) {
                    currentSheetIndex = position
                    renderPreview()
                }
            }

            override fun onNothingSelected(parent: AdapterView<*>?) = Unit
        }

        handleIntent(intent)
    }

    override fun onNewIntent(intent: Intent) {
        super.onNewIntent(intent)
        handleIntent(intent)
    }

    override fun onDestroy() {
        super.onDestroy()
        renderJob?.cancel()
        previewBitmap?.recycle()
        previewBitmap = null
        workbook?.closeQuietly()
        workbook = null
    }

    private fun handleIntent(intent: Intent) {
        when (intent.action) {
            Intent.ACTION_VIEW -> {
                intent.data?.let { loadWorkbookFromUri(it) }
            }

            Intent.ACTION_SEND -> {
                val uriFromExtra = if (Build.VERSION.SDK_INT >= 33) {
                    intent.getParcelableExtra(Intent.EXTRA_STREAM, Uri::class.java)
                } else {
                    @Suppress("DEPRECATION")
                    intent.getParcelableExtra(Intent.EXTRA_STREAM) as? Uri
                }
                (uriFromExtra ?: intent.data)?.let { loadWorkbookFromUri(it) }
            }
        }
    }

    private fun loadWorkbookFromUri(uri: Uri) {
        renderJob?.cancel()
        workbook?.closeQuietly()
        workbook = null
        previewBitmap?.recycle()
        previewBitmap = null
        binding.imgPreview.setImageDrawable(null)
        lastExportUris = emptyList()
        updateButtons()

        binding.progress.isVisible = true
        binding.txtStatus.text = "正在读取文件…"

        lifecycleScope.launch {
            val result = withContext(Dispatchers.IO) {
                runCatching { ExcelLoader.load(contentResolver, uri) }
            }

            binding.progress.isVisible = false

            result.onSuccess { loaded ->
                workbook = loaded.workbook
                workbookName = loaded.displayNameWithoutExt.ifBlank { "workbook" }
                currentSheetIndex = 0

                val adapter = ArrayAdapter(
                    this@MainActivity,
                    android.R.layout.simple_spinner_dropdown_item,
                    loaded.sheetNames,
                )
                binding.spinnerSheets.adapter = adapter
                binding.spinnerSheets.isEnabled = loaded.sheetNames.size > 1

                binding.txtStatus.text =
                    "已打开：${loaded.displayName}（${loaded.sheetNames.size} 个工作表）"
                updateButtons()
                renderPreview()
            }.onFailure { e ->
                binding.txtStatus.text = "打开失败：${e.message ?: e.javaClass.simpleName}"
                showMessage("打开失败：${e.message ?: e.javaClass.simpleName}")
            }
        }
    }

    private fun renderPreview() {
        val wb = workbook ?: return
        renderJob?.cancel()

        binding.progress.isVisible = true
        binding.txtStatus.text = "正在生成预览…"
        updateButtons()

        renderJob = lifecycleScope.launch {
            val result = withContext(Dispatchers.Default) {
                runCatching {
                    // Preview should be lightweight to avoid OOM on large sheets.
                    ExcelBitmapRenderer.renderSheet(
                        workbook = wb,
                        sheetIndex = currentSheetIndex,
                        options = RenderOptions(
                            scale = 0.6f,
                            maxBitmapDimension = 4096,
                            maxTotalPixels = 8_000_000L,
                            // Keep preview responsive; export will do a deeper pass.
                            trimMaxCells = 20_000,
                            columnWidthMaxCells = 20_000,
                            autoFitMaxCells = 20_000,
                            maxAutoRowHeightPx = 900,
                        ),
                    )
                }
            }

            binding.progress.isVisible = false

            result.onSuccess { renderResult ->
                val bmp = renderResult.bitmaps.firstOrNull()
                previewBitmap?.recycle()
                previewBitmap = bmp
                binding.imgPreview.setImageBitmap(bmp)
                // Free additional parts if preview got split unexpectedly.
                renderResult.bitmaps.drop(1).forEach { it.recycle() }

                binding.txtStatus.text = buildString {
                    append("预览：")
                    append(wb.getSheetAt(currentSheetIndex).sheetName)
                    if (renderResult.wasSplit) append("（预览已分段）")
                    if (renderResult.warnings.isNotEmpty()) {
                        append("  ")
                        append(renderResult.warnings.joinToString("；"))
                    }
                }
            }.onFailure { e ->
                previewBitmap?.recycle()
                previewBitmap = null
                binding.imgPreview.setImageDrawable(null)
                binding.txtStatus.text = "预览失败：${e.message ?: e.javaClass.simpleName}"
                showErrorDialog("预览失败", e)
            }

            updateButtons()
        }
    }

    private fun exportCurrentSheet() {
        val wb = workbook ?: return

        binding.progress.isVisible = true
        binding.txtStatus.text = "正在导出图片…"
        updateButtons()

        lifecycleScope.launch {
            val exportResult = withContext(Dispatchers.Default) {
                runCatching {
                    ExcelExporter.exportSheetAsPng(
                        context = this@MainActivity,
                        workbook = wb,
                        sheetIndex = currentSheetIndex,
                        baseName = workbookName,
                        options = RenderOptions(
                            scale = 1.4f,
                            maxBitmapDimension = 16_000,
                            maxTotalPixels = 20_000_000L,
                            // Export can spend more time to improve readability.
                            trimMaxCells = 250_000,
                            columnWidthMaxCells = 250_000,
                            autoFitMaxCells = 250_000,
                            maxAutoRowHeightPx = 2000,
                        ),
                    )
                }
            }

            binding.progress.isVisible = false

            exportResult.onSuccess { uris ->
                lastExportUris = uris
                binding.txtStatus.text =
                    if (uris.size == 1) "导出完成：已保存 1 张图片（可直接分享）"
                    else "导出完成：已保存 ${uris.size} 张图片（因尺寸限制自动分段）"
                updateButtons()
            }.onFailure { e ->
                binding.txtStatus.text = "导出失败：${e.message ?: e.javaClass.simpleName}"
                showMessage("导出失败：${e.message ?: e.javaClass.simpleName}")
                showErrorDialog("导出失败", e)
                updateButtons()
            }
        }
    }

    private fun shareLastExport() {
        if (lastExportUris.isEmpty()) {
            showMessage("请先导出图片。")
            return
        }

        val intent = if (lastExportUris.size == 1) {
            Intent(Intent.ACTION_SEND).apply {
                type = "image/png"
                putExtra(Intent.EXTRA_STREAM, lastExportUris.first())
                addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
            }
        } else {
            Intent(Intent.ACTION_SEND_MULTIPLE).apply {
                type = "image/png"
                putParcelableArrayListExtra(Intent.EXTRA_STREAM, ArrayList(lastExportUris))
                addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
            }
        }

        startActivity(Intent.createChooser(intent, "分享图片"))
    }

    private fun updateButtons() {
        val hasWorkbook = workbook != null
        val busy = binding.progress.isVisible
        binding.btnExport.isEnabled = hasWorkbook && !busy
        binding.btnShare.isEnabled = lastExportUris.isNotEmpty() && !busy
    }

    private fun needsLegacyWritePermission(): Boolean {
        // MediaStore on Android 10+ doesn't require legacy write permission.
        return Build.VERSION.SDK_INT <= 28
    }

    private fun showMessage(msg: String) {
        Snackbar.make(binding.root, msg, Snackbar.LENGTH_LONG).show()
    }

    private fun showErrorDialog(title: String, t: Throwable) {
        val details = (t.stackTraceToString()).take(10_000)
        AlertDialog.Builder(this)
            .setTitle(title)
            .setMessage(details)
            .setPositiveButton("复制") { _, _ ->
                val clipboard = getSystemService(Context.CLIPBOARD_SERVICE) as ClipboardManager
                clipboard.setPrimaryClip(ClipData.newPlainText("error", details))
                showMessage("已复制错误信息")
            }
            .setNegativeButton("关闭", null)
            .show()
    }
}
