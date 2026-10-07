import SwiftUI

@MainActor
struct BudgetSheetView: View {
    let route: BudgetSheet
    @ObservedObject var model: BudgetViewModel
    let data: BudgetSnapshot
    @Environment(\.dismiss) private var dismiss
    @State private var deleteCandidate: BudgetTransaction?
    @State private var pickerYear = 2026
    private var title: String {
        switch route {
        case .input(let t): model.name(t == nil || t?.id == 0 ? "entry.add" : "entry.edit")
        case .inputEntry: model.name("entry.add")
        case .plan: model.name("entry.plan")
        case .transaction: model.name("records.detail")
        case .reserve(let r): r.title
        case .archive: model.name("common.archive")
        case .settlement: model.name("monthly.summary")
        case .fixed: model.name("monthly.fixed")
        case .tetris: model.name("monthly.tetris")
        case .month: model.name("common.month")
        case .about: "테트리스 가계부"
        case .living: model.name("balance.living")
        case .purpose(let group): group
        case .variable: model.name("monthly.variable")
        case .cash: model.name("monthly.cash")
        case .reserveUse: model.name("monthly.reserveUse")
        case .classification: "이번달 집계 기준"
        case .rename: "이름 변경"
        case .resetNames: "이름 복원"
        }
    }
    var body: some View {
        ScrollView {
            VStack(alignment: .leading, spacing: 17) {
                HStack {
                    Text(title).font(.title3.weight(.bold))
                    Spacer()
                    Button { dismiss() } label: {
                        Image(systemName: "xmark").font(.caption.weight(.semibold)).foregroundStyle(Palette.muted)
                            .frame(width: 44, height: 44).background(Palette.line, in: Circle())
                    }.accessibilityLabel("닫기")
                }.padding(.bottom, 5)
                switch route {
                case .input(let transaction): TransactionEditor(model: model, transaction: transaction)
                case .plan(let plan): PlanEditor(model: model, plan: plan)
                case .inputEntry(let entry): TransactionEditor(model: model, transaction: nil, template: entry)
                case .transaction(let t): transactionDetail(t)
                case .reserve(let r): reserveDetail(r)
                case .archive: RawArchiveScreen(model: model)
                case .settlement: settlement
                case .fixed: fixed
                case .tetris: tetris
                case .month: monthPicker
                case .about: about
                case .living: living
                case .variable, .cash, .reserveUse, .classification:
                    MonthlyDetail(route: route, model: model, data: data)
                case .rename(let id): RenameSheet(model: model, labelID: id)
                case .resetNames:
                    Text("탭과 메뉴 이름을 모두 기본값으로 되돌릴까요?").font(.subheadline)
                    Button("이름만 기본값으로 복원") { model.personalization.names = [:]; model.sheet = nil; model.showToast("이름을 기본값으로 되돌렸어요.") }.buttonStyle(PrimaryButtonStyle())
                case .purpose(let group):
                    BudgetCard {
                        ForEach(data.reserves.filter { $0.group == group && $0.balance > 0 }) { r in
                            Button { model.sheet = .reserve(r) } label: { AmountRow(title: r.title, amount: r.balance) }.buttonStyle(.plain)
                        }
                    }
                }
            }.padding(.horizontal, 22).padding(.top, 28).padding(.bottom, 30)
        }.background(Palette.paper).tint(Palette.violet).foregroundStyle(Palette.ink)
            .confirmationDialog("거래를 삭제할까요?", isPresented: Binding(get: { deleteCandidate != nil }, set: { if !$0 { deleteCandidate = nil } }), titleVisibility: .visible, presenting: deleteCandidate) { transaction in
                Button(model.name("entry.delete"), role: .destructive) { model.delete(transaction); deleteCandidate = nil }
                Button("취소", role: .cancel) { deleteCandidate = nil }
            } message: { Text("결산과 잔액을 다시 계산해요. 원본 RAW와 수정 이력은 보관합니다.") }
            .alert("처리하지 못했어요", isPresented: Binding(get: { model.errorMessage != nil }, set: { if !$0 { model.errorMessage = nil } })) {
                Button("확인") { model.errorMessage = nil }
            } message: { Text(model.errorMessage ?? "") }
    }

    private func transactionDetail(_ t: BudgetTransaction) -> some View {
        VStack(alignment: .leading, spacing: 17) {
            Text(t.date.replacingOccurrences(of: "T", with: " ")).font(.caption).foregroundStyle(Palette.muted)
            Text("\(won(t.amount)) 원").font(.largeTitle.weight(.semibold)).monospacedDigit()
            BudgetCard {
                detailRow("내용", t.title)
                Divider()
                detailRow("종류", t.kind)
                Divider()
                detailRow("분류", t.category)
                Divider()
                detailRow("결제수단", t.account)
            }
            if !t.note.isEmpty { BudgetNote(text: t.note) }
            if t.id == 2333 { BudgetNote(text: "변동지출에 이미 포함된 거래예요.\n추가 실지출로 다시 더하지 않아요.") }
            if t.id == 2346 { BudgetNote(text: "공동계좌에 납입한 금액이에요.\n이후 공동계좌 결제는 개인 소비로 다시 세지 않아요.") }
            StatusBadge(text: t.status == "confirmed" ? "확정" : t.status == "unconfirmed" ? "미확인" : "예상")
            Button(model.name("entry.edit")) { model.sheet = .input(t) }.buttonStyle(PrimaryButtonStyle())
            Button(model.name("entry.delete"), role: .destructive) { deleteCandidate = t }.font(.caption).frame(minHeight: 44)
            Text("원본 RAW는 보관하고, 수정 내용은 별도로 기록합니다.").font(.caption2).foregroundStyle(Palette.muted)
        }
    }

    private func reserveDetail(_ r: ReserveItem) -> some View {
        VStack(alignment: .leading, spacing: 17) {
            Text("남은 금액").font(.caption).foregroundStyle(Palette.muted)
            Text("\(won(r.balance)) 원").font(.largeTitle.weight(.semibold)).monospacedDigit()
            BudgetCard {
                AmountRow(title: "누적 실제적립", amount: r.allocated)
                Divider()
                AmountRow(title: "누적 사용", amount: r.used)
                Divider()
                AmountRow(title: "\(model.month)월 적립계획", amount: r.plan)
            }
            if !r.date.isEmpty { detailRow("예정일", r.date) }
            BudgetNote(text: r.note)
            Text("예정일과 실제 사용일은 따로 관리해요. 미래의 계획을 사용 완료로 바꾸지 않아요.")
                .font(.caption).foregroundStyle(Palette.muted).lineSpacing(5)
        }
    }

    private var settlement: some View {
        VStack(alignment: .leading, spacing: 17) {
            StatusBadge(text: data.hasForecast ? "예상 포함" : "확정")
            BudgetNote(text: "남는 돈은 총수입에서 지출·적립을 뺀 결산이에요. 통장 잔액과 구분합니다. 미확인 고정비와 예상 월급은 확정 금액에 포함하지 않아요.")
            BudgetCard {
                AmountRow(title: "확정 기록 수입", amount: data.actualIncome)
                AmountRow(title: "확정 지출·적립", amount: data.actualOutflow)
                Divider()
                AmountRow(title: "확정 기록 남는 돈", amount: data.actualRemaining, emphasized: true)
            }
            BudgetCard {
                AmountRow(title: "총수입 · 예상 포함", amount: data.income)
                AmountRow(title: "고정비 · 예상 포함", amount: data.monthlyFocus.fixed)
                AmountRow(title: "비고정비", amount: data.monthlyFocus.variable)
                AmountRow(title: "공동계좌 적립", amount: data.monthlyFocus.jointAllocation)
                AmountRow(title: "예비비 실제적립", amount: data.monthlyFocus.reserveAllocation)
                AmountRow(title: "투자금 적립", amount: data.monthlyFocus.investmentAllocation)
                AmountRow(title: "환급·할인 차감", amount: -data.monthlyFocus.refund)
                Divider()
                AmountRow(title: "지출·적립 합계", amount: data.outflow, emphasized: true)
                AmountRow(title: "남는 돈 · 예상 포함", amount: data.remaining, emphasized: true)
            }
            Button(model.name("monthly.fixed") + " 확인하기") { model.sheet = .fixed }.buttonStyle(PrimaryButtonStyle())
        }
    }

    private var fixed: some View {
        VStack(alignment: .leading, spacing: 17) {
            Text("항목을 눌러 실제 결제로 바꾸세요. 같은 예상 항목을 대체하므로 두 번 더하지 않아요.")
                .font(.caption).foregroundStyle(Palette.muted)
            BudgetCard {
                ForEach(data.fixed) { item in
                    Button { openFixed(item) } label: {
                        HStack {
                            VStack(alignment: .leading, spacing: 5) {
                                Text(item.title).font(.subheadline)
                                Text(item.account).font(.caption2).foregroundStyle(Palette.muted)
                            }
                            Spacer()
                            VStack(alignment: .trailing, spacing: 5) {
                                Text("\(won(item.amount))원").font(.caption.weight(.medium)).monospacedDigit()
                                StatusBadge(text: item.status)
                            }
                        }.frame(minHeight: 44)
                    }.buttonStyle(.plain)
                    if item.id != data.fixed.last?.id { Divider() }
                }
            }
        }
    }
    private func openFixed(_ item: FixedItem) {
        if let plan = item.planID {
            if let linked = model.entries.first(where: { $0.planID == plan }) {
                model.sheet = .input(linked.displayTransaction)
            } else {
                var entry = LedgerEntry()
                entry.month = model.monthKey; entry.date = model.monthKey + "-01T12:00:00"
                entry.title = item.title; entry.account = item.account; entry.amount = item.amount
                entry.expenseClass = "fixed"; entry.category = "고정비"; entry.planID = plan
                entry.funding = item.account == "CASH" ? "bank" : "card"
                model.sheet = .inputEntry(entry)
            }
        } else if let entry = model.entries.first(where: { $0.month == model.monthKey && $0.title == item.title && $0.account == item.account && $0.expenseClass == "fixed" }) {
            model.sheet = .input(entry.displayTransaction)
        }
    }

    private var tetris: some View {
        VStack(alignment: .leading, spacing: 17) {
            if data.tetrisPending {
                BudgetNote(text: "테트리스 결산 검토 대기\n변경된 거래는 원 단위로 계산했어요. 표시 규칙이 확정되기 전에는 월 테트리스 잔액을 임의로 만들지 않아요.")
                AmountRow(title: "현재 원 단위 결산", amount: data.remaining)
                BudgetCard {
                    ForEach(model.entries.filter { $0.month == model.monthKey && $0.displayAmount != nil }) { entry in
                        AmountRow(title: entry.title + " · 직접 지정", amount: entry.displayAmount ?? 0)
                    }
                }
            } else {
                BudgetCard {
                    AmountRow(title: "원 단위 결산", amount: data.remaining)
                    AmountRow(title: "테트리스 결산 · 원본", amount: data.tetrisRemaining)
                    Divider()
                    AmountRow(title: "차이", amount: data.remaining - data.tetrisRemaining, emphasized: true)
                    AmountRow(title: "원 단위 지출·적립", amount: data.outflow)
                    AmountRow(title: "테트리스 지출·적립", amount: data.tetrisOutflow)
                }
                Text("1칸 = 1,000원 · 원본 표시값").font(.caption2).foregroundStyle(Palette.muted)
                BudgetCard {
                    ForEach(data.dailyAmounts) { day in
                        HStack {
                            Text("\(model.month)월 \(day.day)일")
                            Spacer()
                            Text("\(won(day.actual)) → \(won(day.tetris))원").monospacedDigit()
                        }.font(.caption).frame(minHeight: 28)
                    }
                }
            }
        }
    }

    private var monthPicker: some View {
        VStack(alignment: .leading, spacing: 17) {
            Stepper("\(String(pickerYear))년", value: $pickerYear, in: 2000...2100)
            LazyVGrid(columns: Array(repeating: GridItem(.flexible()), count: 3), spacing: 10) {
                ForEach(1...12, id: \.self) { month in
                    Button {
                        model.selectPeriod(year: pickerYear, month: month); model.sheet = nil
                    } label: {
                        Text("\(month)월").font(.subheadline.weight(.semibold))
                            .frame(maxWidth: .infinity, minHeight: 50)
                            .background(model.month == month && model.year == pickerYear ? Palette.lavender : .white,
                                        in: RoundedRectangle(cornerRadius: 14))
                    }
                }
            }
            BudgetNote(text: "10월 이관 데이터부터 결산을 연결했습니다. 이전 RAW 원문은 백업에 보관하며, 과거 월 전체 결산의 이관은 후속 단계에서 진행합니다.")
        }.onAppear { pickerYear = model.year }
    }

    private var about: some View {
        VStack(alignment: .leading, spacing: 17) {
            Text("기존 엑셀을 기준으로 만든 개인용 가계부입니다.").font(.subheadline)
            BudgetCard {
                detailRow("이관 기준일", data.asOf)
                detailRow("앱 버전", "MVP 01")
            }
            BudgetNote(text: "거래는 iPhone 기기에 저장하고 수정·삭제 시 결산과 잔액을 함께 계산해요. 예비비 계획과 실제적립, 카드 사용과 대금 납부는 따로 관리합니다.")
            Text("전체 RAW 원문은 보존합니다. 현재 운영 이관은 10월 45건이며, 엑셀 파일 직접 추가 이관·전체 과거 결산·할부 자동 인식은 아직 지원하지 않습니다.")
                .font(.caption).foregroundStyle(Palette.muted).lineSpacing(5)
        }
    }

    private var living: some View {
        VStack(alignment: .leading, spacing: 17) {
            Text("\(won(data.living)) 원").font(.largeTitle.weight(.semibold)).monospacedDigit()
            BudgetCard {
                AmountRow(title: "생활비", amount: data.living)
                AmountRow(title: "목적자금", amount: data.reserve)
                Divider()
                AmountRow(title: "WRAP 총잔액", amount: data.wrap, emphasized: true)
            }
            Text("생활비 이동 내역").font(.subheadline.weight(.semibold))
            BudgetCard {
                ForEach(model.wrapTransactions.filter { model.entry($0.id)?.funding == "living" || model.entry($0.id)?.operation == "reserveReturn" }) { t in
                    Button { model.sheet = .transaction(t) } label: { TransactionRow(transaction: t) }.buttonStyle(.plain)
                }
            }
            BudgetNote(text: "일반 계좌 잔액과 카드 이용액은 WRAP 총잔액에 더하지 않아요.")
        }
    }

    private func detailRow(_ title: String, _ value: String) -> some View {
        HStack(alignment: .top, spacing: 14) {
            Text(title).font(.caption).foregroundStyle(Palette.muted).frame(width: 65, alignment: .leading)
            Spacer(minLength: 0)
            Text(value).font(.subheadline).multilineTextAlignment(.trailing)
        }.frame(minHeight: 28)
    }
}
