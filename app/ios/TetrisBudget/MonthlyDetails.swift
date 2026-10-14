import SwiftUI

@MainActor
struct MonthlyDetail: View {
    let route: BudgetSheet
    @ObservedObject var model: BudgetViewModel
    let data: BudgetSnapshot
    var body: some View {
        VStack(alignment: .leading, spacing: 17) {
            switch route {
            case .variable:
                VariableExpenseList(model: model, data: data)
            case .cash:
                Text("예비비 계획 밖의 계좌 현금성 지출").font(.caption).foregroundStyle(Palette.muted)
                largeAmount(data.monthlyFocus.unplannedCash)
                BudgetCard {
                    ForEach(data.cashItems) { item in
                        if let t = data.transactions.first(where: { $0.id == item.sourceRow }) {
                            Button { model.sheet = .transaction(t) } label: { TransactionRow(transaction: t) }.buttonStyle(.plain)
                        }
                    }
                }
                BudgetCard {
                    AmountRow(title: "발생한 비예비비", amount: data.monthlyFocus.unplannedCash)
                    AmountRow(title: "이미 고정·비고정비에 포함", amount: data.monthlyFocus.cashAlreadyCounted)
                    Divider()
                    AmountRow(title: "결산에 추가할 금액", amount: data.monthlyFocus.cashAdditional, emphasized: true)
                }
                BudgetNote(text: "카드 지출이나 미리 계획한 예비비 사용과 구분해요. 공동계좌 납입은 현금 소비와 따로 관리해요.")
            case .reserveUse:
                Text("\(model.month)월 목적자금 실제 사용").font(.caption).foregroundStyle(Palette.muted)
                largeAmount(data.monthlyFocus.reserveUse)
                BudgetCard {
                    ForEach(["가족행사비", "여행비", "경조사비", "상환비", "특별비"], id: \.self) { group in
                        AmountRow(title: group, amount: data.reserveUseItems.filter { t in data.reserves.first(where: { $0.id == t.reserveID })?.group == group }.reduce(0) { $0 - $1.amount })
                    }
                }
                if data.reserveUseItems.isEmpty { BudgetNote(text: "이번달 확정된 사용은 아직 없어요. 사용한 달의 결산에 반영합니다.") }
                ForEach(data.reserveUseItems) { t in
                    Button { model.sheet = .transaction(t) } label: { TransactionRow(transaction: t) }.buttonStyle(.plain)
                }
                Button(model.name("tab.reserve") + " 보기") { model.sheet = nil; model.selectedTab = .reserve }.buttonStyle(PrimaryButtonStyle())
            case .classification:
                BudgetNote(text: "고정비·비고정비는 지출 성격, 비예비비·예비비 사용은 자금 사용을 살펴보는 메뉴예요. 네 카드의 금액을 모두 더하지 않아요.")
                BudgetCard {
                    AmountRow(title: model.name("monthly.fixed"), amount: data.monthlyFocus.fixed)
                    AmountRow(title: model.name("monthly.variable"), amount: data.monthlyFocus.variable)
                    AmountRow(title: "공동계좌 적립", amount: data.monthlyFocus.jointAllocation)
                    AmountRow(title: "환급·할인 차감", amount: -data.monthlyFocus.refund)
                    AmountRow(title: "예비비 실제적립", amount: data.monthlyFocus.reserveAllocation)
                    AmountRow(title: "투자금 적립", amount: data.monthlyFocus.investmentAllocation)
                    Divider()
                    AmountRow(title: "지출·적립 합계", amount: data.outflow, emphasized: true)
                }
                Text("비예비비는 성격별 지출에 이미 포함되어 추가 합산하지 않아요. 예비비 사용도 같은 거래를 한 번 반영합니다.")
                    .font(.caption).foregroundStyle(Palette.muted).lineSpacing(5)
            default: EmptyView()
            }
        }
    }
    private func largeAmount(_ amount: Int) -> some View {
        Text("\(won(amount)) 원").font(.largeTitle.weight(.semibold)).monospacedDigit()
    }
}

@MainActor
private struct VariableExpenseList: View {
    @ObservedObject var model: BudgetViewModel
    let data: BudgetSnapshot
    private struct DayGroup: Identifiable {
        let id: String
        let transactions: [BudgetTransaction]
        var total: Int { transactions.reduce(0) { $0 - $1.amount } }
    }
    private var transactions: [BudgetTransaction] {
        let ids = Set(data.variableRecordIDs)
        return data.transactions.filter { ids.contains($0.id) }.sorted {
            if $0.date == $1.date { return $0.id > $1.id }
            return $0.date > $1.date
        }
    }
    private var groups: [DayGroup] {
        let grouped = Dictionary(grouping: transactions) { $0.date.contains("날짜 미확인") ? "" : String($0.date.prefix(10)) }
        return grouped.keys.sorted(by: >).map { DayGroup(id: $0, transactions: grouped[$0] ?? []) }
    }
    var body: some View {
        VStack(alignment: .leading, spacing: 16) {
            HStack {
                VStack(alignment: .leading, spacing: 6) {
                    Text("\(model.year)년 \(model.month)월 · 확정 지출").font(.caption).foregroundStyle(Palette.muted)
                    Text("\(won(data.monthlyFocus.variable)) 원").font(.title.weight(.semibold)).monospacedDigit()
                }
                Spacer()
                Text("\(transactions.count)건").font(.caption).foregroundStyle(Palette.muted)
            }
            if transactions.isEmpty {
                BudgetNote(text: "확정 기록이 없어요.")
            } else {
                LazyVStack(spacing: 7) {
                    ForEach(groups) { group in
                        VStack(spacing: 0) {
                            dayHeader(group)
                            ForEach(group.transactions) { t in
                                Button { model.sheet = .transaction(t) } label: { transactionRow(t) }.buttonStyle(.plain)
                                if t.id != group.transactions.last?.id { Divider().padding(.horizontal, 12) }
                            }
                        }.background(.white, in: RoundedRectangle(cornerRadius: 12))
                    }
                }
            }
            BudgetNote(text: "확정된 비고정비만 표시해요. 계좌 현금성 지출·예비비 사용이 포함된 경우 같은 거래를 다시 더하지 않아요. 거래를 누르면 전체 메모를 확인하고 수정할 수 있어요.")
        }
    }
    private func dayHeader(_ group: DayGroup) -> some View {
        HStack(spacing: 7) {
            if group.id.isEmpty {
                Text("날짜 미확인").font(.subheadline.weight(.semibold))
            } else {
                Text(String(Int(group.id.suffix(2)) ?? 0)).font(.title2.weight(.bold)).monospacedDigit()
                Text(weekday(group.id)).font(.caption2).foregroundStyle(Palette.muted)
                    .padding(.horizontal, 6).padding(.vertical, 3).background(Palette.line, in: RoundedRectangle(cornerRadius: 4))
            }
            Spacer(minLength: 5)
            Text("일 합계").font(.caption2).foregroundStyle(Palette.muted)
            Text("\(won(group.total))원").font(.subheadline.weight(.semibold)).monospacedDigit()
                .foregroundStyle(group.total < 0 ? Palette.blue : Palette.red)
        }.padding(.horizontal, 12).padding(.vertical, 11).background(Palette.lavender.opacity(0.45))
    }
    private func weekday(_ date: String) -> String {
        let pieces = date.split(separator: "-").compactMap { Int($0) }
        guard pieces.count == 3 else { return "" }
        var calendar = Calendar(identifier: .gregorian)
        calendar.timeZone = TimeZone(secondsFromGMT: 0)!
        guard let value = calendar.date(from: DateComponents(year: pieces[0], month: pieces[1], day: pieces[2])) else { return "" }
        return ["일요일", "월요일", "화요일", "수요일", "목요일", "금요일", "토요일"][calendar.component(.weekday, from: value) - 1]
    }
    private func transactionRow(_ t: BudgetTransaction) -> some View {
        let entry = model.entry(t.id)
        let raw = model.rawArchive.first { r in
            guard let key = entry?.sourceKey else { return false }
            return r.importKey == key || (r.importKey == nil && r.id == t.id && key.hasSuffix(":" + String(r.id)))
        }
        let useOriginalCategory = raw != nil && (["변동비", "비고정비", "고정비"].contains(t.category) || t.category == raw?.category)
        let category = useOriginalCategory ? (raw?.category ?? t.category) : t.category
        let subcategory = useOriginalCategory ? (raw?.subcategory ?? "") : ""
        let originalNote = [raw?.additional, raw?.note].compactMap { $0 }.filter { !$0.isEmpty }.joined(separator: " · ")
        let memo = raw != nil && t.note == originalNote ? (raw?.additional ?? "") : t.note
        return HStack(spacing: 8) {
            VStack(alignment: .leading, spacing: 2) {
                Text(category).font(.caption)
                if !subcategory.isEmpty { Text(subcategory).font(.caption2) }
            }.foregroundStyle(Palette.muted).lineLimit(1).frame(width: 68, alignment: .leading)
            VStack(alignment: .leading, spacing: 3) {
                (Text(t.title).font(.subheadline.weight(.medium)) + Text(memo.isEmpty ? "" : " " + memo).font(.caption).foregroundColor(Palette.muted)).lineLimit(1)
                Text(t.account).font(.caption).foregroundStyle(Palette.muted).lineLimit(1)
            }.frame(maxWidth: .infinity, alignment: .leading)
            Text("\(won(-t.amount))원").font(.subheadline).monospacedDigit().lineLimit(1).minimumScaleFactor(0.7)
                .foregroundStyle(t.amount > 0 ? Palette.blue : Palette.red).layoutPriority(1)
        }.padding(.horizontal, 12).padding(.vertical, 13).frame(minHeight: 64).contentShape(Rectangle())
            .accessibilityElement(children: .ignore)
            .accessibilityLabel([category, subcategory, t.title, memo, t.account, "\(won(-t.amount))원"].filter { !$0.isEmpty }.joined(separator: ", "))
    }
}
