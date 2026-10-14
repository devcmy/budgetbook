import SwiftUI

@MainActor
struct PlanEditor: View {
    @ObservedObject var model: BudgetViewModel
    @State private var draft: LedgerPlan
    @State private var amount: String
    @State private var error = ""
    init(model: BudgetViewModel, plan: LedgerPlan?) {
        self.model = model
        let value = plan ?? LedgerPlan(id: UUID().uuidString, month: model.monthKey, type: "fixed", title: "", account: "개인 계좌", status: "예상", amount: 0)
        _draft = State(initialValue: value); _amount = State(initialValue: String(value.amount))
    }
    var body: some View {
        VStack(alignment: .leading, spacing: 12) {
            BudgetCard {
                TextField("월 (YYYY-MM)", text: $draft.month).keyboardType(.numbersAndPunctuation)
                Picker("계획 종류", selection: $draft.type) {
                    Text("고정비 예상").tag("fixed")
                    Text("수입 예상").tag("income")
                    Text("예비비 적립계획").tag("reserve")
                    if draft.type == "familyContribution" { Text("가족계좌 납입 · 중단").tag("familyContribution") }
                }
                TextField("항목 이름", text: $draft.title).autocorrectionDisabled()
                TextField("예상 금액 (원)", text: $amount).keyboardType(.numberPad)
                TextField("계좌·카드 이름", text: $draft.account).autocorrectionDisabled()
            }
            BudgetNote(text: "예상·계획을 바꿔도 실제 잔액은 바뀌지 않아요. 실제 수입·고정비 입력에서 이 항목을 연결하면 예상 금액을 대체합니다. 취소한 계획은 금액을 0으로 기록해주세요.")
            if draft.type == "familyContribution" {
                BudgetNote(text: "가족계좌 납입은 2026년 10월부터 중단됐어요. 원본 계획은 보존하고 적용 금액은 0원으로 유지합니다.")
            }
            if !error.isEmpty { Text(error).font(.caption).foregroundStyle(Palette.red) }
            Button("예상·계획 저장") {
                guard let value = Int(amount.replacingOccurrences(of: ",", with: "")), value >= 0, value <= 9_999_999_999,
                      !draft.title.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty else {
                    error = "항목 이름과 원 단위 예상 금액을 입력해주세요."; return
                }
                var plan = draft; plan.amount = value
                plan.title = plan.title.trimmingCharacters(in: .whitespacesAndNewlines)
                model.savePlan(plan)
            }.buttonStyle(PrimaryButtonStyle())
        }
    }
}
