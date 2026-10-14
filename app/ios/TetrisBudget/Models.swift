import Foundation
import SwiftUI

// A derived view of the persisted ledger, calculated by the shared accounting engine.
struct BudgetSnapshot: Decodable {
    let asOf, month: String
    let actualIncome, actualOutflow, actualRemaining: Int
    let hasRecords, hasForecast, tetrisPending, historicalPending: Bool
    let archivedCount: Int
    let familyPlan, sourceCount: Int
    let wrapRecordIDs: [Int]
    let reserveUseItems: [BudgetTransaction]
    let income, outflow, remaining, tetrisOutflow, tetrisRemaining: Int
    let wrap, living, reserve, reserveAllocated, reserveUsed: Int
    let reserveMonthPlan, reserveMonthActual: Int
    let transactions: [BudgetTransaction]
    let reserves: [ReserveItem]
    let fixed: [FixedItem]
    let monthlyFocus: MonthlyFocus
    let cashItems: [CashItem]
    let dailyAmounts: [DayAmount]
    let variableRecordIDs: [Int]
}

struct BudgetTransaction: Codable, Identifiable {
    var id: Int
    var date, title, account, category: String
    var amount: Int
    var kind, note: String
    var isPreview: Bool
    var operation: String = "expense"
    var status: String = "confirmed"
    var reserveID: Int? = nil
    var recordIncome: Int = 0
    var recordExpense: Int = 0
    var symbol: String {
        kind == "수입" ? "arrow.down.to.line" : kind == "지출" ? "cart" : "arrow.left.arrow.right"
    }
    var tint: Color { kind == "수입" ? Palette.green : kind == "지출" ? Palette.ink : Palette.blue }
    var day: String { date.contains("날짜 미확인") ? String(date.prefix(7)) + " · 날짜 미확인" : String(date.prefix(10)) }
}

struct ReserveItem: Decodable, Identifiable {
    let id: Int
    let group, title, date, note: String
    let plan, allocated, used, balance: Int
}

struct FixedItem: Decodable, Identifiable {
    var id: String { planID ?? title + account }
    let planID: String?
    let title: String
    let amount: Int
    let status, account: String
}

enum BudgetTab: String, CaseIterable {
    case home, records, reserve, balance, settings
    var labelKey: String { "tab." + rawValue }
    var symbol: String {
        switch self {
        case .home: "square.grid.2x2"
        case .records: "list.bullet.rectangle"
        case .reserve: "gift"
        case .balance: "wallet.pass"
        case .settings: "gearshape"
        }
    }
}

enum BudgetSheet: Identifiable {
    case plan(LedgerPlan?), input(BudgetTransaction?), inputEntry(LedgerEntry), transaction(BudgetTransaction), reserve(ReserveItem)
    case archive, settlement, fixed, tetris, month, about, living, purpose(String)
    case variable, cash, reserveUse, classification, rename(String), resetNames
    var id: String {
        switch self {
        case .plan(let plan): "plan-\(plan?.id ?? "new")"
        case .input(let t): "input-\(t?.id ?? 0)"
        case .inputEntry(let entry): "entry-\(entry.id)-\(entry.planID ?? "new")"
        case .transaction(let t): "transaction-\(t.id)"
        case .reserve(let r): "reserve-\(r.id)"
        case .purpose(let group): "purpose-\(group)"
        case .archive: "archive"
        case .settlement: "settlement"
        case .fixed: "fixed"
        case .tetris: "tetris"
        case .month: "month"
        case .about: "about"
        case .living: "living"
        case .variable: "variable"
        case .cash: "cash"
        case .reserveUse: "reserveUse"
        case .classification: "classification"
        case .rename(let key): "rename-\(key)"
        case .resetNames: "resetNames"
        }
    }
}

@MainActor
final class BudgetViewModel: ObservableObject {
    @Published private(set) var snapshot: BudgetSnapshot?
    let labels: [InterfaceLabel]
    @Published var personalization: Personalization { didSet { personalization.save() } }
    @Published var selectedTab: BudgetTab = .home
    @Published private(set) var year = 2026
    @Published private(set) var month = 10
    @Published var sheet: BudgetSheet?
    @Published private(set) var entries: [LedgerEntry] = []
    @Published private(set) var plans: [LedgerPlan] = []
    @Published private(set) var rawArchive: [RawRecord] = []
    @Published var errorMessage: String?
    @Published var toast: String?
    private var toastTask: Task<Void, Never>?
    private var core: LedgerCoreBridge?
    private var store: LedgerFileStore?
    private var ledger: Data?
    var monthKey: String { String(format: "%04d-%02d", year, month) }

    init() {
        let loadedLabels: [InterfaceLabel]
        if let url = Bundle.main.url(forResource: "labels", withExtension: "json"), let data = try? Data(contentsOf: url) {
            loadedLabels = (try? JSONDecoder().decode([InterfaceLabel].self, from: data)) ?? []
        } else { loadedLabels = [] }
        labels = loadedLabels
        personalization = Personalization.load(catalog: loadedLabels)
        do {
            let engine = try LedgerCoreBridge(), fileStore = try LedgerFileStore()
            core = engine; store = fileStore
            let bytes: Data
            if let saved = try fileStore.load() { bytes = saved }
            else {
                guard let url = Bundle.main.url(forResource: "seed-ledger", withExtension: "json") else {
                    throw LedgerError(message: "초기 이관 데이터를 찾을 수 없습니다.")
                }
                bytes = try Data(contentsOf: url)
            }
            let validated = try engine.call("validate", state: bytes)
            let data = try decodeSummary(validated, month: monthKey)
            let meta = try JSONDecoder().decode(LedgerMetadata.self, from: validated)
            // Existing corrupt data is never replaced by seed automatically.
            if try fileStore.load() == nil { try fileStore.write(validated) }
            ledger = validated; entries = meta.transactions; plans = meta.plans; rawArchive = meta.sourceArchive; snapshot = data
        } catch { errorMessage = "데이터를 열지 못했습니다. 원본은 유지했습니다.\n" + error.localizedDescription }
    }

    func defaultName(_ key: String) -> String { labels.first(where: { $0.id == key })?.defaultName ?? key }
    func name(_ key: String) -> String { personalization.names[key] ?? defaultName(key) }
    var transactions: [BudgetTransaction] { snapshot?.transactions ?? [] }
    func entry(_ id: Int) -> LedgerEntry? { entries.first { $0.id == id } }
    var wrapTransactions: [BudgetTransaction] {
        let ids = snapshot?.wrapRecordIDs ?? []
        return ids.compactMap { id in entries.first(where: { $0.id == id }).map { $0.displayTransaction } }
    }
    func selectPeriod(year: Int, month: Int) {
        do {
            guard let ledger else { throw LedgerError(message: "가계부를 먼저 복원해주세요.") }
            let key = String(format: "%04d-%02d", year, month)
            let value = try decodeSummary(ledger, month: key)
            self.year = year; self.month = month; snapshot = value
        } catch { errorMessage = error.localizedDescription }
    }
    func moveMonth(_ delta: Int) {
        let index = year * 12 + month - 1 + delta
        let nextYear = index / 12, nextMonth = index % 12 + 1
        guard (1...9999).contains(nextYear) else { return }
        selectPeriod(year: nextYear, month: nextMonth)
    }
    private func decodeSummary(_ state: Data, month: String) throws -> BudgetSnapshot {
        guard let core else { throw LedgerError(message: "계산 모듈이 준비되지 않았습니다.") }
        return try JSONDecoder().decode(BudgetSnapshot.self, from: core.call("summary", state: state, month: month))
    }
    @discardableResult
    func save(_ entry: LedgerEntry) -> Bool {
        do {
            guard let ledger, let core else { throw LedgerError(message: "가계부를 먼저 복원해주세요.") }
            let candidate = try core.call("mutate", state: ledger, change: ["type": "save", "entry": entry.dictionary()])
            try commit(candidate, month: entry.month)
            selectedTab = .records; sheet = nil; showToast("거래를 저장했어요."); return true
        } catch { errorMessage = error.localizedDescription; return false }
    }
    func delete(_ transaction: BudgetTransaction) {
        do {
            guard let ledger, let core else { throw LedgerError(message: "가계부를 먼저 복원해주세요.") }
            let candidate = try core.call("mutate", state: ledger, change: ["type": "delete", "id": transaction.id])
            try commit(candidate, month: monthKey)
            sheet = nil; showToast("거래를 삭제했어요.")
        } catch { errorMessage = error.localizedDescription }
    }
    private func commit(_ bytes: Data, month: String, restoring: Bool = false) throws {
        guard let store else { throw LedgerError(message: "기기 저장소에 접근할 수 없습니다.") }
        guard bytes.count <= 32 * 1024 * 1024 else { throw LedgerError(message: "가계부 파일이 32MB를 초과했습니다. 저장 전 백업과 이관 범위를 확인해주세요.") }
        let data = try decodeSummary(bytes, month: month)
        let meta = try JSONDecoder().decode(LedgerMetadata.self, from: bytes)
        try store.write(bytes, restoring: restoring) // Only publish successful disk writes.
        ledger = bytes; entries = meta.transactions; plans = meta.plans; rawArchive = meta.sourceArchive
        year = Int(month.prefix(4)) ?? year; self.month = Int(month.suffix(2)) ?? self.month
        snapshot = data
    }
    func savePlan(_ plan: LedgerPlan) {
        do {
            guard let ledger, let core else { throw LedgerError(message: "가계부를 먼저 복원해주세요.") }
            let object = try JSONSerialization.jsonObject(with: JSONEncoder().encode(plan))
            let candidate = try core.call("mutate", state: ledger, change: ["type": "savePlan", "plan": object])
            try commit(candidate, month: plan.month)
            sheet = nil; showToast("예상·계획을 저장했어요.")
        } catch { errorMessage = error.localizedDescription }
    }
    func backupData() throws -> Data {
        // Permit exporting unreadable/corrupt original for rescue without modifying it.
        guard let ledger else {
            if let bytes = try store?.load() { return bytes }
            throw LedgerError(message: "백업할 파일이 없습니다.")
        }
        var object = try JSONSerialization.jsonObject(with: ledger) as? [String: Any] ?? [:]
        object["personalization"] = try JSONSerialization.jsonObject(with: JSONEncoder().encode(personalization))
        return try JSONSerialization.data(withJSONObject: object, options: [.prettyPrinted, .sortedKeys])
    }
    func reviewBackup(_ bytes: Data) throws -> BackupReview {
        guard bytes.count <= 32 * 1024 * 1024, let core else { throw LedgerError(message: "백업 크기 또는 계산 모듈을 확인해주세요.") }
        let validated = try core.call("validate", state: bytes)
        let summary = try decodeSummary(validated, month: monthKey)
        let meta = try JSONDecoder().decode(LedgerMetadata.self, from: validated)
        if let settings = meta.personalization {
            for (key, value) in settings.names {
                guard let label = labels.first(where: { $0.id == key }), Personalization.validName(value, limit: label.maxLength) != nil else {
                    throw LedgerError(message: "백업의 메뉴 이름을 확인해주세요.")
                }
            }
        }
        return BackupReview(data: validated, count: meta.transactions.count, wrap: summary.wrap)
    }
    func reviewPreviousSave() throws -> BackupReview {
        guard let store else { throw LedgerError(message: "기기 저장소에 접근할 수 없습니다.") }
        return try reviewBackup(Data(contentsOf: store.previousURL))
    }
    func restore(_ bytes: Data) {
        do {
            let backup = try reviewBackup(bytes)
            let meta = try JSONDecoder().decode(LedgerMetadata.self, from: backup.data)
            try commit(backup.data, month: monthKey, restoring: true)
            if let settings = meta.personalization { personalization = settings }
            showToast("백업을 복원했어요.")
        } catch { errorMessage = error.localizedDescription }
    }
    func reviewSpreadsheet(_ bytes: Data, fileName: String) throws -> SpreadsheetReview {
        guard bytes.count <= 32 * 1024 * 1024, let ledger, let core else {
            throw LedgerError(message: "32MB 이하의 엑셀을 선택하고 가계부를 먼저 열어주세요.")
        }
        let workbook = try core.spreadsheet("read", state: ledger, payload: ["file": bytes.base64EncodedString(), "fileName": fileName])
        return try reviewSpreadsheetTable(workbook, sheet: nil)
    }
    func reviewSpreadsheetTable(_ workbook: Data, sheet: String?) throws -> SpreadsheetReview {
        guard let ledger, let core else { throw LedgerError(message: "가계부를 먼저 열어주세요.") }
        var payload: [String: Any] = ["workbook": try JSONSerialization.jsonObject(with: workbook)]
        if let sheet { payload["sheet"] = sheet }
        let proposal = try core.spreadsheet("review", state: ledger, payload: payload)
        return try SpreadsheetReview(workbook: workbook, proposal: proposal)
    }
    @discardableResult
    func importSpreadsheet(_ review: SpreadsheetReview) -> Bool {
        do {
            guard let ledger, let core else { throw LedgerError(message: "가계부를 먼저 열어주세요.") }
            let candidate = try core.spreadsheet("apply", state: ledger, payload: ["proposal": try review.dictionary()])
            let month = review.rows.first { $0.include && $0.disposition == "new" }?.entry?.month ?? monthKey
            try commit(candidate, month: month)
            selectedTab = .records; showToast("\(review.selectedCount)건을 가져왔어요."); return true
        } catch { errorMessage = error.localizedDescription; return false }
    }
    func showToast(_ text: String) {
        toastTask?.cancel(); toast = text
        toastTask = Task { [weak self] in
            try? await Task.sleep(for: .seconds(2.5))
            guard !Task.isCancelled else { return }
            self?.toast = nil
        }
    }
}

extension LedgerEntry {
    var displayTransaction: BudgetTransaction {
        let incoming = ["income", "refund"].contains(operation) || (operation == "jointActivity" && sourceKind == "수입")
        let spending = ["expense", "reserveUse"].contains(operation) || (operation == "jointActivity" && sourceKind == "지출")
        return BudgetTransaction(id: id, date: date ?? month + " · 날짜 미확인", title: title, account: account, category: category,
                                 amount: incoming ? amount : -amount, kind: incoming ? "수입" : spending ? "지출" : "이체", note: note, isPreview: false,
                                 operation: operation, status: status, reserveID: reserveID)
    }
}

func won(_ value: Int) -> String { value.formatted(.number.locale(Locale(identifier: "ko_KR"))) }

struct MonthlyFocus: Decodable {
    let fixed, variable, unplannedCash, cashAlreadyCounted, cashAdditional, reserveUse, refund, jointAllocation, reserveAllocation, investmentAllocation: Int
}
struct CashItem: Decodable, Identifiable {
    var id: Int { sourceRow }
    let sourceRow: Int
    let title, date: String
    let amount, alreadyCounted, additional: Int
    let note: String
}
struct DayAmount: Decodable, Identifiable {
    var id: Int { day }
    let day, actual, tetris: Int
}
