package com.zys.excel2image

import java.io.Closeable

inline fun <T : Closeable, R> T?.useOrThrow(block: (T) -> R): R {
    val value = this ?: error("Stream is null")
    return value.use(block)
}

fun Closeable.closeQuietly() {
    try {
        close()
    } catch (_: Exception) {
        // ignore
    }
}
