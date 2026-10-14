import SwiftUI

@MainActor
struct TransactionEditor: View {
    @ObservedObject var model: BudgetViewModel
    let transaction: BudgetTransaction?
    @State private var draft: LedgerEntry
    @State private var amount: String
    @State private var date: Date
    @State private var unknownDate: Bool
    @State private var manualDisplay = ""
    @State private var error = ""
    private let initialDate: Date
    private let types: [(String, String)] = [
        ("expense", "지출"), ("income", "수입"), ("refund", "환급·정산"),
        ("reserveAllocate", "예비비 실제적립"), ("reserveUse", "예비비 사용"), ("reserveReturn", "예비비 → 생활비 환입"),
        ("jointAllocate", "공동계좌 납입"), ("jointActivity", "공동계좌 활동"),
        ("investmentAllocate", "투자금 적립"), ("investmentBuy", "투자 매수"),
        ("cardPayment", "카드대금 납부"), ("transfer", "계좌 이동"), ("pending", "대조 차이·보류")
    ]
    private let locations = [("card", "카드"), ("bank", "일반 계좌·현금"), ("living", "WRAP 생활비"), ("joint", "공동계좌")]
    private static let formatter: DateFormatter = {
        let f = DateFormatter(); f.locale = Locale(identifier: "en_US_POSIX")
        f.timeZone = TimeZone(identifier: "Asia/Seoul"); f.dateFormat = "yyyy-MM-dd'T'HH:mm:ss"; return f
    }()
    private static func parse(_ text: String?) -> Date? {
        guard let text else { return nil }
        return formatter.date(from: String(text.prefix(19)))
    }
    init(model: BudgetViewModel, transaction: BudgetTransaction?, template: LedgerEntry? = nil) {
        self.model = model; self.transaction = transaction
        var entry = transaction.flatMap { model.entry($0.id) } ?? template ?? LedgerEntry()
        if entry.id == 0 && template == nil {
            entry.month = model.monthKey
            let now = Date(), nowString = Self.formatter.string(from: now)
            entry.date = nowString.hasPrefix(model.monthKey) ? nowString : model.monthKey + "-01T12:00:00"
            if transaction?.kind == "이체" { entry.operation = "reserveAllocate"; entry.funding = "living"; entry.reserveID = 22 }
        }
        let parsed = Self.parse(entry.date) ?? Self.formatter.date(from: entry.month + "-01T12:00:00") ?? .now
        initialDate = parsed
        _draft = State(initialValue: entry); _amount = State(initialValue: entry.id == 0 ? "" : String(entry.amount))
        _date = State(initialValue: parsed); _unknownDate = State(initialValue: entry.date == nil)
        _manualDisplay = State(initialValue: entry.displayAmount.map(String.init) ?? "")
    }
    private var needsReserve: Bool { ["reserveAllocate", "reserveUse", "reserveReturn"].contains(draft.operation) }
    private var selectablePlans: [LedgerPlan] {
        let m = unknownDate ? draft.month : String(Self.formatter.string(from: date).prefix(7))
        return model.plans.filter { $0.month == m && ((draft.operation == "income" && $0.type == "income") || (draft.operation == "expense" && draft.expenseClass == "fixed" && $0.type == "fixed")) }
    }
    private var relatedExpenses: [LedgerEntry] {
        model.entries.filter { $0.id != draft.id && ["expense", "reserveUse"].contains($0.operation) && (draft.operation != "cardPayment" || $0.funding == "card") }
    }
    var body: some View {
        VStack(alignment: .leading, spacing: 12) {
            BudgetCard {
                Picker("거래 종류", selection: $draft.operation) {
                    ForEach(types, id: \.0) { Text($0.1).tag($0.0) }
                }.onChange(of: draft.operation) { _, _ in
                    draft.planID = nil; draft.linkedID = nil; draft.attributedMonth = nil
                    if needsReserve && draft.reserveID == nil { draft.reserveID = model.snapshot?.reserves.first?.id }
                    if draft.operation == "jointActivity" { draft.funding = "joint"; if draft.sourceKind == nil { draft.sourceKind = "지출" } }
                    if draft.id == 0 { draft.category = types.first(where: { $0.0 == draft.operation })?.1 ?? draft.category }
                }
                TextField("금액 (원)", text: $amount).font(.title2.weight(.semibold)).keyboardType(.numbersAndPunctuation)
                    .accessibilityLabel("원 단위 금액")
                TextField("내용", text: $draft.title).autocorrectionDisabled()
                Picker("상태", selection: $draft.status) {
                    Text("확정").tag("confirmed"); Text("미확인").tag("unconfirmed"); Text("예상·계획").tag("planned")
                }.pickerStyle(.segmented)
                Text("확정 거래만 실제 금액과 잔액에 반영합니다.").font(.caption2).foregroundStyle(Palette.muted)
            }
            BudgetCard {
                Toggle("날짜 미확인 · 월만 기록", isOn: $unknownDate)
                if unknownDate {
                    TextField("거래 월 (YYYY-MM)", text: $draft.month).keyboardType(.numbersAndPunctuation)
                } else { DatePicker("실제 거래일", selection: $date, displayedComponents: [.date, .hourAndMinute]) }
                TextField("원본 계좌·카드 이름", text: $draft.account).autocorrectionDisabled()
                TextField("분류", text: $draft.category).autocorrectionDisabled()
                if draft.operation == "jointActivity" {
                    Picker("공동계좌 입출금", selection: $draft.sourceKind) {
                        Text("지출").tag(Optional("지출")); Text("수입").tag(Optional("수입"))
                    }
                }
                if !["reserveUse", "reserveReturn", "investmentBuy", "pending"].contains(draft.operation) {
                    Picker(draft.operation == "income" || draft.operation == "refund" ? "입금 위치" : "출금 위치", selection: $draft.funding) {
                        ForEach(locations, id: \.0) { Text($0.1).tag($0.0) }
                    }
                }
                if draft.operation == "transfer" {
                    Picker("도착 위치", selection: $draft.destination) {
                        Text("일반 계좌").tag("bank"); Text("WRAP 생활비").tag("living"); Text("공동계좌").tag("joint")
                    }
                    Text("공동계좌 납입은 별도 거래 종류를 선택해주세요.").font(.caption2).foregroundStyle(Palette.muted)
                }
                if needsReserve {
                    Picker("예비비 목적", selection: $draft.reserveID) {
                        Text("선택해주세요").tag(nil as Int?)
                        ForEach(model.snapshot?.reserves ?? []) { r in Text(r.title).tag(Optional(r.id)) }
                    }
                    Text("생활비에서 실제적립하면 WRAP 총잔액은 그대로이고 목적자금만 늘어요.").font(.caption2).foregroundStyle(Palette.muted)
                }
            }
            if ["expense", "reserveUse"].contains(draft.operation) {
                BudgetCard {
                    Picker("지출 성격", selection: $draft.expenseClass) {
                        Text(model.name("monthly.variable")).tag("variable"); Text(model.name("monthly.fixed")).tag("fixed")
                    }.pickerStyle(.segmented)
                    Toggle("이미 정산을 뺀 본인부담 순액", isOn: $draft.alreadyNet)
                    Text("순액이면 연결된 환급을 결산에서 다시 빼지 않아요.").font(.caption2).foregroundStyle(Palette.muted)
                }
            }
            if !selectablePlans.isEmpty {
                BudgetCard {
                    Picker("대체할 예상 항목", selection: $draft.planID) {
                        Text("새 거래 · 예상과 별도").tag(nil as String?)
                        ForEach(selectablePlans) { p in Text("\(p.title) · \(won(p.amount))원").tag(Optional(p.id)) }
                    }
                    Text("예상 월급·고정비와 같은 건이면 연결해서 중복을 막아요.").font(.caption2).foregroundStyle(Palette.muted)
                }
            }
            if ["refund", "cardPayment"].contains(draft.operation) {
                BudgetCard {
                    Picker("원래 지출", selection: $draft.linkedID) {
                        Text(draft.operation == "refund" ? "월 환급·할인 (개별 연결 없음)" : "사용 내역 선택").tag(nil as Int?)
                        ForEach(relatedExpenses) { t in Text("\(t.month) · \(t.title) · \(won(t.amount))원").tag(Optional(t.id)) }
                    }
                    if draft.operation == "refund" {
                        TextField("결산 귀속 월 (YYYY-MM)", text: Binding(get: { draft.attributedMonth ?? "" }, set: { draft.attributedMonth = $0.isEmpty ? nil : $0 }))
                            .keyboardType(.numbersAndPunctuation)
                        Text("다른 달의 정산은 원래 지출 월 또는 받은 월을 직접 적어주세요.").font(.caption2).foregroundStyle(Palette.muted)
                    } else { Text("사용할 때 소비, 납부할 때 계좌 출금만 반영해요.").font(.caption2).foregroundStyle(Palette.muted) }
                }
            }
            BudgetCard {
                TextField("메모", text: $draft.note, axis: .vertical).lineLimit(2...5).autocorrectionDisabled()
                TextField("테트리스 표시 금액 (선택)", text: $manualDisplay).keyboardType(.numberPad)
                Text("새 거래를 자동 반올림하지 않아요. 월 테트리스 결산은 규칙 확정 전까지 검토 대기로 표시합니다.")
                    .font(.caption2).foregroundStyle(Palette.muted)
            }
            if draft.operation == "pending" { BudgetNote(text: "대조 차이는 보류로 보관하며 수입이나 이자로 자동 반영하지 않아요.") }
            if !error.isEmpty { Text(error).font(.caption).foregroundStyle(Palette.red) }
            Button(model.name("entry.save")) { save() }.buttonStyle(PrimaryButtonStyle())
        }
    }
    private func save() {
        let normalized = amount.replacingOccurrences(of: ",", with: "").trimmingCharacters(in: .whitespacesAndNewlines)
        guard let number = Int(normalized), (-9_999_999_999...9_999_999_999).contains(number) else { error = "원 단위 정수 금액을 입력해주세요."; return }
        let title = draft.title.trimmingCharacters(in: .whitespacesAndNewlines)
        guard !title.isEmpty else { error = "내용을 입력해주세요."; return }
        var value = draft; value.amount = number; value.title = title
        if unknownDate { value.date = nil }
        else {
            let text = Self.formatter.string(from: date)
            value.month = String(text.prefix(7))
            value.date = date == initialDate && draft.date != nil ? draft.date : text
        }
        if manualDisplay.isEmpty { value.displayAmount = nil }
        else {
            guard let display = Int(manualDisplay), display >= 0 else { error = "테트리스 표시 금액을 확인해주세요."; return }
            value.displayAmount = display
        }
        if !needsReserve { value.reserveID = nil }
        if !selectablePlans.contains(where: { $0.id == value.planID }) { value.planID = nil }
        _ = model.save(value)
    }
}
