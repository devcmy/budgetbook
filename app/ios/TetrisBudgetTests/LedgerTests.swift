import XCTest
@testable import TetrisBudget

final class LedgerTests: XCTestCase {
    private func fixture() throws -> Data {
        let url = try XCTUnwrap(Bundle.main.url(forResource: "seed-ledger", withExtension: "json"))
        return try Data(contentsOf: url)
    }
    @MainActor
    func testJavaScriptCoreReproducesExcelAndAcceptsNativeOptionalFields() async throws {
        let core = try LedgerCoreBridge(), state = try fixture()
        let summary = try JSONDecoder().decode(BudgetSnapshot.self, from: core.call("summary", state: state, month: "2026-10"))
        XCTAssertEqual(summary.remaining, 1_195_895)
        XCTAssertEqual(summary.wrap, 8_480_881)
        XCTAssertEqual(summary.variableRecordIDs.count, 19)
        var entry = LedgerEntry()
        entry.title = "기기 검증"; entry.date = "2026-10-06T12:00:00"; entry.amount = 10_000; entry.funding = "living"
        let changed = try core.call("mutate", state: state, change: ["type": "save", "entry": entry.dictionary()])
        let live = try JSONDecoder().decode(BudgetSnapshot.self, from: core.call("summary", state: changed, month: "2026-10"))
        XCTAssertEqual(live.wrap, 8_470_881)
        XCTAssertEqual(live.remaining, 1_185_895)
        XCTAssertEqual(live.monthlyFocus.cashAdditional, 0)
    }
    @MainActor
    func testNativeDiskSaveRelaunchRoundTripAndPreviousCopy() async throws {
        let directory = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        defer { try? FileManager.default.removeItem(at: directory) }
        let store = try LedgerFileStore(directory: directory), core = try LedgerCoreBridge()
        let before = try fixture()
        try store.write(before)
        var entry = LedgerEntry()
        entry.title = "재실행 검증"; entry.date = "2026-10-06T12:00:00"; entry.amount = 12_345
        let after = try core.call("mutate", state: before, change: ["type": "save", "entry": entry.dictionary()])
        try store.write(after)
        let reopened = try LedgerFileStore(directory: directory)
        XCTAssertEqual(try reopened.load(), after)
        XCTAssertEqual(try Data(contentsOf: store.previousURL), before)
        let validated = try core.call("validate", state: XCTUnwrap(reopened.load()))
        let metadata = try JSONDecoder().decode(LedgerMetadata.self, from: validated)
        XCTAssertEqual(metadata.transactions.count, 46)
    }
    @MainActor
    func testCorruptFileIsPreservedAndRejectedInsteadOfReset() async throws {
        let directory = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        defer { try? FileManager.default.removeItem(at: directory) }
        let store = try LedgerFileStore(directory: directory), core = try LedgerCoreBridge()
        let corrupt = Data("{broken backup".utf8)
        try store.write(corrupt)
        let read = try XCTUnwrap(store.load())
        XCTAssertThrowsError(try core.call("validate", state: read))
        XCTAssertEqual(try store.load(), corrupt)
    }
    @MainActor
    func testFailedWriteThrowsWithoutReplacingLastSavedData() async throws {
        let directory = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        defer { try? FileManager.default.removeItem(at: directory) }
        let store = try LedgerFileStore(directory: directory)
        let before = try fixture()
        try store.write(before)
        // A directory at the previous-file destination simulates an unavailable backup destination.
        try FileManager.default.createDirectory(at: store.previousURL, withIntermediateDirectories: false)
        XCTAssertThrowsError(try store.write(Data("{}".utf8)))
        XCTAssertEqual(try store.load(), before)
    }
    @MainActor
    func testRestorePreservesGoodPreviousCopyAndCorruptOriginal() async throws {
        let directory = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        defer { try? FileManager.default.removeItem(at: directory) }
        let store = try LedgerFileStore(directory: directory)
        let good = try fixture()
        try store.write(good); try store.write(good)
        let corrupt = Data("{corrupt".utf8)
        try corrupt.write(to: store.url, options: .atomic)
        try store.write(good, restoring: true)
        XCTAssertEqual(try store.load(), good)
        XCTAssertEqual(try Data(contentsOf: store.previousURL), good)
        XCTAssertEqual(try Data(contentsOf: store.beforeRestoreURL), corrupt)
    }

}
