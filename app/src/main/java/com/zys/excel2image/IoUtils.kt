package com.zys.excel2image

import java.io.Closeable
import java.io.InputStream
import java.io.OutputStream

inline fun <T : Closeable?, R> T.useOrThrow(block: (T) -> R): R {
    if (this == null) error("Stream is null")
    return this.use(block)
}

fun Closeable.closeQuietly() {
    try {
        close()
    } catch (_: Exception) {
        // ignore
    }
}

