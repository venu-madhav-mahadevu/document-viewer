// iOS (SwiftUI + PDFKit). No third-party libraries.
//
// Usage:
//   DocumentViewer(fileURL: localFileURL, serverURL: URL(string: "https://docs.internal.example/api/convert")!)

import SwiftUI
import PDFKit

struct DocumentViewer: View {
    let fileURL: URL      // file already in the app sandbox
    let serverURL: URL    // your /api/convert endpoint

    @State private var document: PDFDocument?
    @State private var errorText: String?

    var body: some View {
        Group {
            if let document {
                PDFKitView(document: document)
            } else if let errorText {
                Text(errorText).foregroundStyle(.red).padding()
            } else {
                ProgressView("Converting…")
            }
        }
        .navigationTitle(fileURL.lastPathComponent)
        .task { await load() }
    }

    private func load() async {
        do {
            let data = try Data(contentsOf: fileURL)
            if fileURL.pathExtension.lowercased() == "pdf" {
                document = PDFDocument(data: data)   // PDFs need no conversion
                return
            }
            var request = URLRequest(url: serverURL)
            request.httpMethod = "POST"
            request.setValue(fileURL.lastPathComponent.addingPercentEncoding(withAllowedCharacters: .urlPathAllowed),
                             forHTTPHeaderField: "X-File-Name")
            let (pdf, response) = try await URLSession.shared.upload(for: request, from: data)
            guard (response as? HTTPURLResponse)?.statusCode == 200 else {
                errorText = "This document could not be converted."
                return
            }
            document = PDFDocument(data: pdf)
        } catch {
            errorText = error.localizedDescription
        }
    }
}

/// Wraps Apple's built-in PDFView (zoom, scroll, search, text selection).
struct PDFKitView: UIViewRepresentable {
    let document: PDFDocument

    func makeUIView(context: Context) -> PDFView {
        let view = PDFView()
        view.autoScales = true
        view.displayMode = .singlePageContinuous
        view.document = document
        return view
    }

    func updateUIView(_ view: PDFView, context: Context) {
        view.document = document
    }
}

// Tip: for a zero-server option on iOS, QLPreviewController (Quick Look)
// previews .docx/.xlsx/.pptx natively on the device.
