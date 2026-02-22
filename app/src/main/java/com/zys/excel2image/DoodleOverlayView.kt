package com.zys.excel2image

import android.content.Context
import android.graphics.Canvas
import android.graphics.Paint
import android.graphics.Path
import android.graphics.PointF
import android.util.AttributeSet
import android.view.MotionEvent
import android.view.View
import com.davemorrissey.labs.subscaleview.SubsamplingScaleImageView
import kotlin.math.abs
import kotlin.math.max

class DoodleOverlayView @JvmOverloads constructor(
    context: Context,
    attrs: AttributeSet? = null,
) : View(context, attrs) {

    var imageView: SubsamplingScaleImageView? = null

    var drawEnabled: Boolean = false
        set(value) {
            field = value
            // Ensure we can intercept touches only when needed.
            isClickable = value
        }

    var brushColorArgb: Int = 0xFFFF0000.toInt()
    var brushWidthViewPx: Float = 10f

    var strokes: List<Stroke> = emptyList()
        set(value) {
            field = value
            invalidate()
        }

    var onStrokeFinished: ((Stroke) -> Unit)? = null

    private var inProgress: Stroke? = null

    private val strokePaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        style = Paint.Style.STROKE
        strokeJoin = Paint.Join.ROUND
        strokeCap = Paint.Cap.ROUND
    }

    private val tmpPath = Path()
    private val tmpPoint = PointF()

    override fun onDraw(canvas: Canvas) {
        super.onDraw(canvas)
        val iv = imageView ?: return
        if (!iv.isReady) return

        val scale = iv.scale
        for (stroke in strokes) {
            drawStroke(canvas, iv, stroke, scale)
        }
        inProgress?.let { drawStroke(canvas, iv, it, scale) }
    }

    private fun drawStroke(
        canvas: Canvas,
        iv: SubsamplingScaleImageView,
        stroke: Stroke,
        scale: Float,
    ) {
        if (stroke.points.isEmpty()) return

        strokePaint.color = stroke.colorArgb
        // Make the stroke scale with the content when zooming.
        strokePaint.strokeWidth = max(1f, stroke.widthSourcePx * scale)

        tmpPath.reset()
        val first = stroke.points.first()
        val firstV = iv.sourceToViewCoord(first.x, first.y) ?: return
        tmpPath.moveTo(firstV.x, firstV.y)

        for (i in 1 until stroke.points.size) {
            val p = stroke.points[i]
            val v = iv.sourceToViewCoord(p.x, p.y) ?: continue
            tmpPath.lineTo(v.x, v.y)
        }

        canvas.drawPath(tmpPath, strokePaint)
    }

    override fun onTouchEvent(event: MotionEvent): Boolean {
        if (!drawEnabled) return false

        val iv = imageView ?: return false
        if (!iv.isReady) return false
        if (event.pointerCount != 1) {
            // Ignore multi-touch in draw mode for now; use "移动" mode to pan/zoom.
            return true
        }

        val src = iv.viewToSourceCoord(event.x, event.y) ?: return true
        tmpPoint.set(src.x, src.y)

        when (event.actionMasked) {
            MotionEvent.ACTION_DOWN -> {
                parent?.requestDisallowInterceptTouchEvent(true)
                val widthSource = (brushWidthViewPx / max(0.0001f, iv.scale)).coerceAtLeast(0.5f)
                inProgress = Stroke(
                    colorArgb = brushColorArgb,
                    widthSourcePx = widthSource,
                    points = mutableListOf(PointF(tmpPoint.x, tmpPoint.y)),
                )
                invalidate()
                return true
            }

            MotionEvent.ACTION_MOVE -> {
                val stroke = inProgress ?: return true
                val last = stroke.points.lastOrNull()
                if (last == null) {
                    stroke.points += PointF(tmpPoint.x, tmpPoint.y)
                } else {
                    // Drop very tiny moves to avoid huge point lists.
                    if (abs(tmpPoint.x - last.x) >= 0.75f || abs(tmpPoint.y - last.y) >= 0.75f) {
                        stroke.points += PointF(tmpPoint.x, tmpPoint.y)
                    }
                }
                invalidate()
                return true
            }

            MotionEvent.ACTION_UP, MotionEvent.ACTION_CANCEL -> {
                val stroke = inProgress
                inProgress = null
                if (stroke != null && stroke.points.size >= 2) {
                    onStrokeFinished?.invoke(stroke)
                }
                parent?.requestDisallowInterceptTouchEvent(false)
                invalidate()
                return true
            }
        }

        return super.onTouchEvent(event)
    }
}
