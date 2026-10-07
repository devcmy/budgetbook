import SwiftUI
import Foundation

struct SpreadsheetRow: Decodable, Identifiable {
    let id: String
    let row: Int
    let disposition, reason: String
    var include: Bool
    var entry: LedgerEntry?
}

struct SpreadsheetProposal: Decodable {
    let fileName, sheet: String
    let total, existing, invalid: Int
    var rows: [SpreadsheetRow]
}

struct SpreadsheetReview: Identifiable {
    let id = UUID()
    let workbook, proposal: Data
    let fileName, sheet: String
    let total, existing, invalid: Int
    let sheets: [String]
    var rows: [SpreadsheetRow]
    var selectedCount: Int { rows.filter { $0.include && $0.disposition == "new" }.count }
    var pendingCount: Int { rows.filter { $0.include && $0.disposition == "new" && $0.entry?.operation == "pending" }.count }
    init(workbook: Data, proposal: Data) throws {
        self.workbook = workbook; self.proposal = proposal
        let decoded = try JSONDecoder().decode(SpreadsheetProposal.self, from: proposal)
        fileName = decoded.fileName; sheet = decoded.sheet; total = decoded.total
        existing = decoded.existing; invalid = decoded.invalid; rows = decoded.rows
        let object = try JSONSerialization.jsonObject(with: workbook) as? [String: Any]
        sheets = (object?["sheets"] as? [[String: Any]] ?? []).compactMap { $0["name"] as? String }
    }
    func dictionary() throws -> [String: Any] {
        var object = try JSONSerialization.jsonObject(with: proposal) as? [String: Any] ?? [:]
        var originals = object["rows"] as? [[String: Any]] ?? []
        guard originals.count == rows.count else { throw LedgerError(message: "엑셀 검토 행 수가 다릅니다.") }
        for index in rows.indices {
            originals[index]["include"] = rows[index].include
            if let entry = rows[index].entry { originals[index]["entry"] = try entry.dictionary() }
        }
        object["rows"] = originals; return object
    }
}

@MainActor
struct SpreadsheetReviewScreen: View {
    @ObservedObject var model: BudgetViewModel
    @State var review: SpreadsheetReview
    @State private var showExisting = false
    @State private var amountTexts: [String: String] = [:]
    @State private var importError: String?
    @Environment(\.dismiss) private var dismiss
    var body: some View {
        NavigationStack {
            ScrollView {
                VStack(alignment: .leading, spacing: 14) {
                    Text(review.fileName).font(.caption).foregroundStyle(Palette.muted)
                    if review.sheets.count > 1 {
                        Picker("거래 시트", selection: Binding(get: { review.sheet }, set: { name in
                            do { review = try model.reviewSpreadsheetTable(review.workbook, sheet: name) }
                            catch { importError = error.localizedDescription }
                        })) {
                            ForEach(review.sheets, id: \.self) { Text($0).tag($0) }
                        }
                    }
                    BudgetCard {
                        Text("읽은 거래 \(review.total)건 · 중복 제외 \(review.existing)건 · 오류 제외 \(review.invalid)건").font(.caption)
                        Text("선택 \(review.selectedCount)건 · 보류 \(review.pendingCount)건").font(.subheadline.weight(.semibold))
                    }
                    BudgetNote(text: "선택한 거래를 현재 가계부에 추가합니다. 고정비·수입은 대체할 예상 항목을 선택해주세요. 이체·용도 미확인은 보류로 저장한 뒤 기록 수정에서 처리합니다. 월별 결산·예비비·잔액 시트는 자동 교체하지 않습니다.")
                    Toggle("중복 거래도 보기", isOn: $showExisting).font(.caption)
                    LazyVStack(spacing: 10) {
                        ForEach(Array(review.rows.enumerated()), id: \.element.id) { index, row in
                            if showExisting || row.disposition != "existing" { rowCard(row, index: index) }
                        }
                    }
                    Button("\(review.selectedCount)건 추가하기") { performImport() }
                        .buttonStyle(PrimaryButtonStyle()).disabled(review.selectedCount == 0)
                    Button("취소") { dismiss() }.frame(minHeight: 44)
                }.padding(18)
            }.background(Palette.paper)
                .navigationTitle("엑셀 거래 가져오기").navigationBarTitleDisplayMode(.inline)
                .toolbar { ToolbarItem(placement: .cancellationAction) { Button("닫기") { dismiss() } } }
                .alert("가져오지 못했어요", isPresented: Binding(get: { importError != nil }, set: { if !$0 { importError = nil } })) {
                    Button("확인") { importError = nil }
                } message: { Text(importError ?? "") }
        }
    }
    @ViewBuilder private func rowCard(_ row: SpreadsheetRow, index: Int) -> some View {
        BudgetCard {
            if let entry = row.entry {
                DisclosureGroup {
                    Text("\(entry.date ?? entry.month + " · 날짜 미확인")\n\(entry.account) · \(entry.category)")
                        .font(.caption2).foregroundStyle(Palette.muted)
                    Text(row.reason.isEmpty ? "새 거래" : row.reason).font(.caption2).foregroundStyle(Palette.muted)
                    if row.disposition == "new" {
                        Toggle("이 거래 가져오기", isOn: Binding(get: { review.rows[index].include }, set: { review.rows[index].include = $0 }))
                        Picker("반영 상태", selection: stringBinding(index, \.status)) {
                            Text("확정").tag("confirmed"); Text("미확인").tag("unconfirmed"); Text("예상").tag("planned")
                        }
                        TextField("금액 (원)", text: Binding(get: { amountTexts[row.id] ?? String(review.rows[index].entry?.amount ?? 0) }, set: { amountTexts[row.id] = $0 }))
                            .keyboardType(.numbersAndPunctuation)
                        if entry.operation == "expense" {
                            Picker("지출 성격", selection: stringBinding(index, \.expenseClass)) {
                                Text(model.name("monthly.fixed")).tag("fixed"); Text(model.name("monthly.variable")).tag("variable")
                            }
                        }
                        if ["expense", "income"].contains(entry.operation) {
                            Picker("계좌 위치", selection: stringBinding(index, \.funding)) {
                                Text("카드").tag("card"); Text("일반 계좌·현금").tag("bank"); Text("WRAP 생활비").tag("living")
                            }
                            let plans = model.plans.filter { $0.month == entry.month && (($0.type == "income" && entry.operation == "income") || ($0.type == "fixed" && entry.operation == "expense" && entry.expenseClass == "fixed")) }
                            if !plans.isEmpty {
                                Picker("대체할 예상 항목", selection: Binding(get: { review.rows[index].entry?.planID ?? "" }, set: { review.rows[index].entry?.planID = $0.isEmpty ? nil : $0 })) {
                                    Text("새 거래 · 예상과 별도").tag("")
                                    ForEach(plans) { Text("\($0.title) · \(won($0.amount))원").tag($0.id) }
                                }
                            }
                        }
                    }
                } label: {
                    VStack(alignment: .leading, spacing: 4) {
                        Text("\(row.row)행 · \(entry.title)").font(.subheadline)
                        Text("\(won(entry.amount))원 · \(row.disposition == "existing" ? "중복 제외" : entry.operation == "pending" ? "처리 보류" : entry.status == "confirmed" ? "확정" : "미확인·예상")")
                            .font(.caption2).foregroundStyle(Palette.muted)
                    }
                }
            } else { Text("\(row.row)행 · \(row.reason)").font(.caption).foregroundStyle(.red) }
        }
    }
    private func stringBinding(_ index: Int, _ key: WritableKeyPath<LedgerEntry, String>) -> Binding<String> {
        Binding(get: { review.rows[index].entry?[keyPath: key] ?? "" }, set: { value in
            review.rows[index].entry?[keyPath: key] = value
            if key == \LedgerEntry.expenseClass { review.rows[index].entry?.planID = nil }
        })
    }
    private func performImport() {
        for index in review.rows.indices where review.rows[index].include && review.rows[index].disposition == "new" {
            if let raw = amountTexts[review.rows[index].id] {
                let value = raw.trimmingCharacters(in: .whitespacesAndNewlines).replacingOccurrences(of: ",", with: "")
                guard value.range(of: "^-?[0-9]+$", options: .regularExpression) != nil, let amount = Int(value) else {
                    importError = "\(review.rows[index].row)행 금액은 원 단위 정수여야 합니다."; return
                }
                review.rows[index].entry?.amount = amount
            }
        }
        if model.importSpreadsheet(review) { dismiss() }
        else { importError = model.errorMessage; model.errorMessage = nil }
    }
}
