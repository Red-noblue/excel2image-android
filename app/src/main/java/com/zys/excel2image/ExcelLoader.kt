package com.zys.excel2image

import android.content.ContentResolver
import android.database.Cursor
import android.net.Uri
import android.provider.OpenableColumns
import org.apache.poi.openxml4j.util.ZipSecureFile
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.BufferedReader
import java.io.InputStreamReader
import java.nio.charset.StandardCharsets

data class LoadedWorkbook(
    val workbook: Workbook,
    val displayName: String,
    val displayNameWithoutExt: String,
    val sheetNames: List<String>,
)

object ExcelLoader {
    fun load(contentResolver: ContentResolver, uri: Uri): LoadedWorkbook {
        // Some real-world xlsx exported by other tools can trip the zip-bomb protection.
        // Lowering the ratio avoids false positives while still being safe for typical files.
        ZipSecureFile.setMinInflateRatio(0.001)

        val displayName = contentResolver.getDisplayName(uri) ?: "workbook.xlsx"
        val nameWithoutExt = displayName.substringBeforeLast('.', displayName)

        val workbook = when (displayName.substringAfterLast('.', "").lowercase()) {
            "csv" -> contentResolver.openInputStream(uri).useOrThrow { input ->
                loadCsvWorkbook(input)
            }
            else -> contentResolver.openInputStream(uri).useOrThrow { input ->
                WorkbookFactory.create(input)
            }
        }

        val sheetNames = (0 until workbook.numberOfSheets).map { workbook.getSheetAt(it).sheetName }

        return LoadedWorkbook(
            workbook = workbook,
            displayName = displayName,
            displayNameWithoutExt = nameWithoutExt,
            sheetNames = sheetNames,
        )
    }

    private fun loadCsvWorkbook(input: java.io.InputStream): Workbook {
        val wb = XSSFWorkbook()
        val sheet = wb.createSheet("Sheet1")

        BufferedReader(InputStreamReader(input, StandardCharsets.UTF_8)).use { br ->
            var rowIndex = 0
            while (true) {
                val line = br.readLine() ?: break
                val values = parseCsvLine(line)
                val row = sheet.createRow(rowIndex++)
                for ((col, v) in values.withIndex()) {
                    row.createCell(col).setCellValue(v)
                }
                // Prevent extreme files from locking the UI; V1 is for "normal" csv.
                if (rowIndex > 10_000) break
            }
        }

        return wb
    }

    // Minimal CSV parser: supports quotes, escaped quotes, commas.
    private fun parseCsvLine(line: String): List<String> {
        val out = ArrayList<String>()
        val sb = StringBuilder()
        var i = 0
        var inQuotes = false

        while (i < line.length) {
            val ch = line[i]
            when {
                ch == '"' -> {
                    if (inQuotes && i + 1 < line.length && line[i + 1] == '"') {
                        // Escaped quote: ""
                        sb.append('"')
                        i++
                    } else {
                        inQuotes = !inQuotes
                    }
                }
                ch == ',' && !inQuotes -> {
                    out.add(sb.toString())
                    sb.setLength(0)
                }
                else -> sb.append(ch)
            }
            i++
        }
        out.add(sb.toString())
        return out
    }
}

private fun ContentResolver.getDisplayName(uri: Uri): String? {
    var cursor: Cursor? = null
    return try {
        cursor = query(uri, arrayOf(OpenableColumns.DISPLAY_NAME), null, null, null)
        if (cursor != null && cursor.moveToFirst()) cursor.getString(0) else null
    } catch (_: Exception) {
        null
    } finally {
        cursor?.close()
    }
}
