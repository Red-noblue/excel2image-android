package com.zys.excel2image

import android.graphics.PointF

data class Stroke(
    val colorArgb: Int,
    // Stroke width in *source image* pixels.
    val widthSourcePx: Float,
    val points: MutableList<PointF>,
)

