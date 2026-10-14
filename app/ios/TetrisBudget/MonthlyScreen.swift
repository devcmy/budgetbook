import SwiftUI

@MainActor
struct HomeScreen: View {
    @ObservedObject var model: BudgetViewModel
    let data: BudgetSnapshot
    @State private var selectedDay = 5
    @Environment(\.dynamicTypeSize) private var typeSize
    private var columns: [GridItem] {
        Array(repeating: GridItem(.flexible(), spacing: 8), count: typeSize.isAccessibilitySize ? 1 : 2)
    }
    private var dayAmount: DayAmount? { data.dailyAmounts.first { $0.day == selectedDay } }
    private var dayRecords: [BudgetTransaction] {
        data.transactions.filter { data.variableRecordIDs.contains($0.id) && Int($0.date.dropFirst(8).prefix(2)) == selectedDay }
    }
    var body: some View {
        HStack {
            VStack(alignment: .leading, spacing: 4) {
                Text("\(model.name("tab.home")) · 저장된 기록 기준").font(.caption2).foregroundStyle(Palette.muted)
            }
            Spacer()
            Button { model.sheet = .input(nil) } label: {
                Image(systemName: "plus").font(.system(size: 20)).frame(width: 44, height: 44)
            }.modifier(GlassControl(reduced: model.personalization.reduceTransparency)).accessibilityLabel("거래 입력")
        }.padding(.bottom, 4)
        HStack(spacing: 10) {
            Button { moveMonth(-1) } label: {
                Image(systemName: "chevron.left").font(.subheadline.weight(.semibold)).frame(width: 44, height: 44)
            }.buttonStyle(.plain).modifier(GlassControl(reduced: model.personalization.reduceTransparency))
                .accessibilityLabel("이전 달").disabled(model.year == 1 && model.month == 1)
            Button { model.sheet = .month } label: {
                Text("\(String(model.year))년 \(model.month)월").font(.title3.weight(.bold))
                    .lineLimit(1).minimumScaleFactor(0.7).frame(maxWidth: .infinity, minHeight: 44)
            }.buttonStyle(.plain).accessibilityLabel("조회월 선택")
            Button { moveMonth(1) } label: {
                Image(systemName: "chevron.right").font(.subheadline.weight(.semibold)).frame(width: 44, height: 44)
            }.buttonStyle(.plain).modifier(GlassControl(reduced: model.personalization.reduceTransparency))
                .accessibilityLabel("다음 달").disabled(model.year == 9999 && model.month == 12)
        }.padding(.bottom, 6)
        if data.historicalPending {
            BudgetNote(text: "원본 RAW \(data.archivedCount)건 보관 중 · 과거 결산은 아직 이관하지 않았어요. 금액은 현재 앱에 등록한 운영 거래만 표시합니다.")
            Button("이 달 원문 보기") { model.sheet = .archive }.font(.caption).frame(minHeight: 44)
        }
        if data.hasRecords || data.hasForecast {
            Button { model.sheet = .settlement } label: { summary }.buttonStyle(.plain)
            LazyVGrid(columns: columns, spacing: 8) {
                focusCard("monthly.variable", amount: data.monthlyFocus.variable, hint: "달력에 기록된 변동지출", symbol: "cart", tint: Palette.violet, route: .variable)
                focusCard("monthly.fixed", amount: data.monthlyFocus.fixed, hint: "예상·미확인 포함", symbol: "list.bullet.rectangle", tint: Palette.blue, route: .fixed)
                focusCard("monthly.cash", amount: data.monthlyFocus.unplannedCash, hint: "계좌에서 쓴 계획 외 현금", symbol: "wallet.pass", tint: Palette.green, route: .cash)
                focusCard("monthly.reserveUse", amount: data.monthlyFocus.reserveUse, hint: "목적자금의 이번달 실제 사용", symbol: "gift", tint: Color(hex: 0xB68143), route: .reserveUse)
            }
            HStack(spacing: 10) {
                Text("공동계좌 적립 \(won(data.monthlyFocus.jointAllocation))")
                Text("환급·할인 −\(won(data.monthlyFocus.refund))")
                Spacer(minLength: 0)
                Button { model.sheet = .classification } label: {
                    Image(systemName: "info.circle").frame(width: 32, height: 44)
                }.accessibilityLabel("집계 기준 설명")
            }.font(.system(size: 10)).foregroundStyle(Palette.muted)
            HStack {
                Text(model.name("monthly.calendar")).font(.subheadline.weight(.semibold))
                Spacer()
                Text("\(model.name("monthly.variable")) · 원 단위").font(.caption2).foregroundStyle(Palette.muted)
            }
            calendar
            HStack {
                Text("\(model.month)월 \(selectedDay)일").font(.caption)
                Text(dayAmount == nil ? "미입력" : model.name("monthly.variable")).font(.caption2).foregroundStyle(Palette.muted)
                Spacer()
                Text(dayAmount.map { "\(won($0.actual))원" } ?? "—").font(.caption.weight(.semibold)).monospacedDigit()
            }.frame(minHeight: 30)
            if dayAmount != nil {
                BudgetCard {
                    ForEach(Array(dayRecords.prefix(3))) { t in
                        Button { model.sheet = .transaction(t) } label: { TransactionRow(transaction: t) }.buttonStyle(.plain)
                        if t.id != Array(dayRecords.prefix(3)).last?.id { Divider() }
                    }
                    if dayRecords.count > 3 {
                        Button("이 날 \(dayRecords.count)건 전체 보기") { model.selectedTab = .records }
                            .font(.caption).frame(minHeight: 44)
                    }
                }
            } else { BudgetNote(text: "아직 실제 기록이 없는 날이에요.") }
            Button { model.sheet = .tetris } label: {
                HStack(spacing: 9) {
                    Image(systemName: "square.grid.3x3.fill").foregroundStyle(Palette.violet)
                    Text(model.name("monthly.tetris")).font(.caption.weight(.medium))
                    Spacer()
                    VStack(alignment: .trailing, spacing: 4) {
                        Text(data.tetrisPending ? "검토 대기" : "\(won(data.tetrisRemaining))원").font(.caption.weight(.semibold)).monospacedDigit()
                        Text(data.tetrisPending ? "표시 규칙 확인 필요" : "예상 · 차이 \(won(data.remaining - data.tetrisRemaining))원").font(.system(size: 10)).foregroundStyle(Palette.muted)
                    }
                    Image(systemName: "chevron.right").font(.system(size: 9)).foregroundStyle(Palette.muted)
                }.padding(12).frame(minHeight: 50)
                    .background(Palette.lavender, in: RoundedRectangle(cornerRadius: 15))
            }.buttonStyle(.plain)
            HStack(spacing: 8) {
                balanceLink(model.name("tab.reserve") + " 잔액", value: data.reserve, tab: .reserve)
                balanceLink("WRAP 잔액", value: data.wrap, tab: .balance)
            }
        } else {
            BudgetCard {
                Text(data.historicalPending ? "운영 결산 미이관" : "아직 미입력").font(.title2.weight(.semibold))
                Text("예정 금액은 실제 거래와 따로 관리해요.").font(.caption).foregroundStyle(Palette.muted)
                AmountRow(title: "\(model.month)월 예비비 적립계획", amount: data.reserveMonthPlan)
            }
            Button("첫 기록 입력하기") { model.sheet = .input(nil) }.buttonStyle(PrimaryButtonStyle())
        }
    }

    private func moveMonth(_ delta: Int) {
        model.moveMonth(delta)
        let cal = Calendar(identifier: .gregorian)
        if let date = cal.date(from: DateComponents(year: model.year, month: model.month, day: 1)),
           let count = cal.range(of: .day, in: .month, for: date)?.count {
            selectedDay = min(selectedDay, count)
        }
    }
    private var summary: some View {
        HStack(alignment: .center, spacing: 15) {
            VStack(alignment: .leading, spacing: 9) {
                HStack(spacing: 7) {
                    Text(model.name("monthly.remaining")).font(.caption2).foregroundStyle(.white.opacity(0.8))
                    Text(data.hasForecast ? "예상 포함" : "확정 기록").font(.system(size: 9)).foregroundStyle(Color(hex: 0xF2C6BC))
                        .padding(4).background(.white.opacity(0.08), in: RoundedRectangle(cornerRadius: 5))
                }
                Text("\(won(data.remaining)) 원").font(.title.weight(.semibold)).monospacedDigit().foregroundStyle(.white)
                    .lineLimit(1).minimumScaleFactor(0.7)
            }.frame(maxWidth: .infinity, alignment: .leading)
            if !typeSize.isAccessibilitySize {
                Rectangle().fill(.white.opacity(0.12)).frame(width: 1, height: 69)
                VStack(alignment: .leading, spacing: 13) {
                    heroMetric(model.name("monthly.income"), value: data.income)
                    heroMetric(model.name("monthly.outflow"), value: data.outflow)
                }
            }
        }.padding(17).frame(maxWidth: .infinity, minHeight: 110)
            .background(Palette.navy, in: RoundedRectangle(cornerRadius: 21))
    }

    private func heroMetric(_ title: String, value: Int) -> some View {
        VStack(alignment: .leading, spacing: 4) {
            Text(title).font(.system(size: 10)).foregroundStyle(.white.opacity(0.7))
            Text("\(won(value))원").font(.caption.weight(.semibold)).monospacedDigit().foregroundStyle(.white)
        }
    }

    private func focusCard(_ key: String, amount: Int, hint: String, symbol: String, tint: Color, route: BudgetSheet) -> some View {
        Button { model.sheet = route } label: {
            VStack(alignment: .leading, spacing: 7) {
                HStack(spacing: 6) {
                    Image(systemName: symbol).foregroundStyle(tint)
                    Text(model.name(key)).lineLimit(1)
                    Spacer(minLength: 0)
                    if key == "monthly.fixed" && data.hasForecast { StatusBadge(text: "예상") }
                }.font(.caption.weight(.semibold))
                HStack(alignment: .firstTextBaseline) {
                    Text("\(won(amount))원").font(.title3.weight(.semibold)).monospacedDigit().lineLimit(1).minimumScaleFactor(0.7)
                    Spacer(minLength: 2)
                    Image(systemName: "chevron.right").font(.system(size: 9)).foregroundStyle(Palette.muted)
                }
                Text(hint).font(.system(size: 10)).foregroundStyle(Palette.muted).lineLimit(2)
            }.frame(maxWidth: .infinity, alignment: .leading).padding(model.personalization.compact ? 12 : 17)
                .background(.white, in: RoundedRectangle(cornerRadius: 17))
                .overlay(RoundedRectangle(cornerRadius: 17).stroke(Palette.line))
        }.buttonStyle(.plain)
    }

    private var calendar: some View {
        let cal = Calendar(identifier: .gregorian)
        let start = cal.date(from: DateComponents(year: model.year, month: model.month, day: 1)) ?? .now
        let offset = (cal.component(.weekday, from: start) + 5) % 7
        let count = cal.range(of: .day, in: .month, for: start)?.count ?? 31
        return BudgetCard {
            LazyVGrid(columns: Array(repeating: GridItem(.flexible(), spacing: 0), count: 7), spacing: 3) {
                ForEach(["월", "화", "수", "목", "금", "토", "일"], id: \.self) { day in
                    Text(day).font(.system(size: 10)).foregroundStyle(Palette.muted).frame(height: 18)
                }
                ForEach(0..<offset, id: \.self) { _ in Color.clear.frame(height: 44) }
                ForEach(1...count, id: \.self) { day in
                    let amount = data.dailyAmounts.first { $0.day == day }
                    Button { selectedDay = day } label: {
                        VStack(spacing: 4) {
                            Text("\(day)").font(.caption.weight(amount == nil ? .regular : .medium))
                            Text(amount.map { won($0.actual) } ?? "—").font(.system(size: 10)).monospacedDigit()
                        }.frame(maxWidth: .infinity, minHeight: 44)
                            .foregroundStyle(selectedDay == day ? .white : amount == nil ? Palette.muted.opacity(0.55) : Palette.ink)
                            .background(selectedDay == day ? Palette.violet : .clear, in: RoundedRectangle(cornerRadius: 10))
                    }.buttonStyle(.plain).accessibilityLabel("\(model.month)월 \(day)일, \(amount.map { "비고정비 \(won($0.actual))원" } ?? "미입력")")
                }
            }
        }
    }

    private func balanceLink(_ title: String, value: Int, tab: BudgetTab) -> some View {
        Button { model.selectedTab = tab } label: {
            VStack(alignment: .leading, spacing: 6) {
                Text(title).font(.caption2).foregroundStyle(Palette.muted)
                Text("\(won(value))원").font(.caption.weight(.semibold)).monospacedDigit()
            }.frame(maxWidth: .infinity, alignment: .leading).padding(12)
                .background(.white, in: RoundedRectangle(cornerRadius: 13))
        }.buttonStyle(.plain)
    }
}
