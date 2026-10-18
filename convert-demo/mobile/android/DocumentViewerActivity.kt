// Android (Kotlin + built-in PdfRenderer). No third-party libraries.
//
// Start with:
//   startActivity(Intent(this, DocumentViewerActivity::class.java)
//       .putExtra("path", localFile.absolutePath))
// Manifest: <uses-permission android:name="android.permission.INTERNET" />

package com.example.docviewer

import android.graphics.Bitmap
import android.graphics.Color
import android.graphics.pdf.PdfRenderer
import android.os.Bundle
import android.os.ParcelFileDescriptor
import android.widget.ImageView
import android.widget.LinearLayout
import android.widget.ScrollView
import android.widget.TextView
import androidx.appcompat.app.AppCompatActivity
import java.io.File
import java.net.HttpURLConnection
import java.net.URL
import java.net.URLEncoder
import java.util.concurrent.Executors

class DocumentViewerActivity : AppCompatActivity() {

    private val serverUrl = "https://docs.internal.example/api/convert" // your /api/convert
    private val io = Executors.newSingleThreadExecutor()
    private lateinit var pages: LinearLayout

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        pages = LinearLayout(this).apply { orientation = LinearLayout.VERTICAL; setBackgroundColor(Color.LTGRAY) }
        setContentView(ScrollView(this).apply { addView(pages) })

        val source = File(requireNotNull(intent.getStringExtra("path")))
        title = source.name

        io.execute {
            try {
                val pdf = if (source.extension.equals("pdf", true)) source else convert(source)
                runOnUiThread { render(pdf) }
            } catch (e: Exception) {
                runOnUiThread { pages.addView(TextView(this).apply { text = "Could not open: ${e.message}" }) }
            }
        }
    }

    /** Sends the file to our own server; returns the converted PDF as a temp file in the app cache. */
    private fun convert(file: File): File {
        val conn = (URL(serverUrl).openConnection() as HttpURLConnection).apply {
            requestMethod = "POST"
            doOutput = true
            setRequestProperty("X-File-Name", URLEncoder.encode(file.name, "UTF-8"))
            setFixedLengthStreamingMode(file.length())
        }
        file.inputStream().use { input -> conn.outputStream.use { input.copyTo(it) } }
        check(conn.responseCode == 200) { "server returned ${conn.responseCode}" }

        val out = File.createTempFile("converted", ".pdf", cacheDir) // app sandbox only
        conn.inputStream.use { input -> out.outputStream().use { input.copyTo(it) } }
        return out
    }

    /** Renders every page with Android's built-in PdfRenderer. */
    private fun render(pdf: File) {
        val width = resources.displayMetrics.widthPixels
        ParcelFileDescriptor.open(pdf, ParcelFileDescriptor.MODE_READ_ONLY).use { fd ->
            PdfRenderer(fd).use { renderer ->
                for (i in 0 until renderer.pageCount) {
                    renderer.openPage(i).use { page ->
                        val height = width * page.height / page.width
                        val bitmap = Bitmap.createBitmap(width, height, Bitmap.Config.ARGB_8888)
                        bitmap.eraseColor(Color.WHITE)
                        page.render(bitmap, null, null, PdfRenderer.Page.RENDER_MODE_FOR_DISPLAY)
                        pages.addView(ImageView(this).apply {
                            setImageBitmap(bitmap)
                            setPadding(0, 0, 0, 16)
                        })
                    }
                }
            }
        }
        if (pdf.parentFile == cacheDir) pdf.delete() // don't leave customer documents on disk
    }

    override fun onDestroy() {
        io.shutdownNow()
        super.onDestroy()
    }
}

// For large documents, render pages lazily in a RecyclerView, or use Jetpack's
// androidx.pdf PdfViewerFragment, which adds zoom, search and text selection.
