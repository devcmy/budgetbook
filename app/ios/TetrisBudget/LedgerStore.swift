import Foundation
import SwiftUI
import UniformTypeIdentifiers

// Local only. Writes fail visibly; no silent reset or fallback to seed on corrupt files.
struct LedgerFileStore {
    let url: URL
    let previousURL: URL
    let beforeRestoreURL: URL
    init(directory supplied: URL? = nil) throws {
        let directory: URL
        if let supplied { directory = supplied }
        else { directory = try FileManager.default.url(for: .applicationSupportDirectory, in: .userDomainMask,
                                                   appropriateFor: nil, create: true).appendingPathComponent("TetrisBudget", isDirectory: true) }
        try FileManager.default.createDirectory(at: directory, withIntermediateDirectories: true)
        url = directory.appendingPathComponent("ledger-v1.json")
        previousURL = directory.appendingPathComponent("ledger-v1.previous.json")
        beforeRestoreURL = directory.appendingPathComponent("ledger-v1.before-restore.json")
    }
    func load() throws -> Data? {
        guard FileManager.default.fileExists(atPath: url.path) else { return nil }
        return try Data(contentsOf: url)
    }
    func write(_ data: Data, restoring: Bool = false) throws {
        if let previous = try load() {
            try previous.write(to: restoring ? beforeRestoreURL : previousURL, options: [.atomic, .completeUntilFirstUserAuthentication])
        }
        try data.write(to: url, options: [.atomic, .completeUntilFirstUserAuthentication])
    }
}

struct LedgerDocument: FileDocument {
    static var readableContentTypes: [UTType] { [.json] }
    var data: Data
    init(data: Data) { self.data = data }
    init(configuration: ReadConfiguration) throws {
        guard let data = configuration.file.regularFileContents else { throw LedgerError(message: "JSON 파일을 읽을 수 없습니다.") }
        self.data = data
    }
    func fileWrapper(configuration: WriteConfiguration) throws -> FileWrapper { FileWrapper(regularFileWithContents: data) }
}

struct BackupReview: Identifiable {
    let id = UUID()
    let data: Data
    let count: Int
    let wrap: Int
}

@MainActor
struct BackupControls: View {
    @ObservedObject var model: BudgetViewModel
    @State private var exporting = false
    @State private var importing = false
    @State private var document = LedgerDocument(data: Data())
    @State private var review: BackupReview?
    @State private var spreadsheetReview: SpreadsheetReview?
    var body: some View {
        BudgetCard {
            Text(model.name("settings.backup")).font(.subheadline.weight(.semibold))
            Text("엑셀은 거래 추가 · JSON은 전체 백업 복원입니다.").font(.caption2).foregroundStyle(Palette.muted)
            Button(model.name("settings.export")) {
                do { document = LedgerDocument(data: try model.backupData()); exporting = true }
                catch { model.errorMessage = error.localizedDescription }
            }.frame(minHeight: 44)
            Button(model.name("settings.import")) { importing = true }.frame(minHeight: 44)
            Button(model.name("settings.previous")) {
                do { review = try model.reviewPreviousSave() }
                catch { model.errorMessage = error.localizedDescription }
            }.font(.caption).frame(minHeight: 44)
            if let data = model.snapshot {
                Text("운영 거래 \(model.entries.count)건 · 원본 RAW \(data.sourceCount)건 보관")
                    .font(.caption2).foregroundStyle(Palette.muted)
            }
        }
        .fileExporter(isPresented: $exporting, document: document, contentType: .json,
                      defaultFilename: "테트리스_가계부_백업") { result in
            switch result {
            case .success: model.showToast("백업을 저장했어요.")
            case .failure(let error): model.errorMessage = error.localizedDescription
            }
        }
        .fileImporter(isPresented: $importing, allowedContentTypes: [.json, UTType(filenameExtension: "xlsx") ?? .spreadsheet, UTType(filenameExtension: "xls") ?? .spreadsheet]) { result in
            do {
                let url = try result.get()
                let granted = url.startAccessingSecurityScopedResource()
                defer { if granted { url.stopAccessingSecurityScopedResource() } }
                let size = try url.resourceValues(forKeys: [.fileSizeKey]).fileSize ?? 0
                guard size <= 32 * 1024 * 1024 else { throw LedgerError(message: "32MB 이하의 엑셀 또는 JSON 파일을 선택해주세요.") }
                let bytes = try Data(contentsOf: url)
                if ["xlsx", "xls"].contains(url.pathExtension.lowercased()) {
                    spreadsheetReview = try model.reviewSpreadsheet(bytes, fileName: url.lastPathComponent)
                } else { review = try model.reviewBackup(bytes) }
            } catch { model.errorMessage = error.localizedDescription }
        }
        .alert(item: $review) { backup in
            Alert(title: Text("이 백업으로 복원할까요?"),
                  message: Text("거래 \(backup.count)건 · WRAP \(won(backup.wrap))원\n현재 데이터 전체를 바꿉니다. 먼저 현재 백업을 저장해주세요."),
                  primaryButton: .destructive(Text("복원")) { model.restore(backup.data) }, secondaryButton: .cancel(Text("취소")))
        }
        .sheet(item: $spreadsheetReview) { value in
            SpreadsheetReviewScreen(model: model, review: value)
        }
    }
}
