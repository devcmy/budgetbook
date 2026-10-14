import SwiftUI

@MainActor
struct ContentView: View {
    @StateObject private var model = BudgetViewModel()
    var body: some View {
        Group {
            if let data = model.snapshot {
                TabView(selection: $model.selectedTab) {
                    ForEach(BudgetTab.allCases, id: \.self) { tab in
                        NavigationStack {
                            ScrollView {
                                VStack(alignment: .leading, spacing: model.personalization.compact ? 8 : 14) {
                                    switch tab {
                                    case .home: HomeScreen(model: model, data: data)
                                    case .records: RecordsScreen(model: model)
                                    case .reserve: ReserveScreen(model: model, data: data)
                                    case .balance: BalanceScreen(model: model, data: data)
                                    case .settings: SettingsScreen(model: model)
                                    }
                                }.padding(.horizontal, model.personalization.compact ? 16 : 20)
                                    .padding(.top, 8).padding(.bottom, 20)
                            }.background(Palette.paper).toolbar(.hidden, for: .navigationBar)
                        }.tabItem { Label(model.name(tab.labelKey), systemImage: tab.symbol) }.tag(tab)
                    }
                }
                .modifier(TabTransparency(reduced: model.personalization.reduceTransparency))
                .overlay(alignment: .bottom) {
                    if let toast = model.toast {
                        Text(toast).font(.caption).foregroundStyle(.white).padding(14)
                            .frame(maxWidth: .infinity).background(Palette.navy, in: RoundedRectangle(cornerRadius: 14))
                            .padding(.horizontal, 20).padding(.bottom, 90).allowsHitTesting(false)
                    }
                }
                .sheet(item: $model.sheet) { route in
                    BudgetSheetView(route: route, model: model, data: data)
                        .presentationDetents([.medium, .large]).presentationDragIndicator(.visible)
                        .presentationCornerRadius(27)
                }
            } else {
                VStack(spacing: 18) {
                    ContentUnavailableView("가계부를 열 수 없어요", systemImage: "doc.badge.exclamationmark",
                                           description: Text("기존 파일을 자동 초기화하지 않습니다. 백업을 보관하거나 복원해주세요."))
                    BackupControls(model: model)
                }.padding()
            }
        }.tint(Palette.violet).foregroundStyle(Palette.ink).preferredColorScheme(.light)
            .environment(\.locale, Locale(identifier: "ko_KR"))
            .environment(\.timeZone, TimeZone(identifier: "Asia/Seoul")!)
            .environment(\.budgetCompact, model.personalization.compact)
            .alert("처리하지 못했어요", isPresented: Binding(get: { model.errorMessage != nil && model.sheet == nil }, set: { if !$0 { model.errorMessage = nil } })) {
                Button("확인") { model.errorMessage = nil }
            } message: { Text(model.errorMessage ?? "") }
    }
}

@MainActor
struct ScreenHeader: View {
    let title: String
    let subtitle: String
    @ObservedObject var model: BudgetViewModel
    var body: some View {
        HStack {
            VStack(alignment: .leading, spacing: 6) {
                Text(title).font(.title2.weight(.bold))
                Text(subtitle).font(.caption2).foregroundStyle(Palette.muted)
            }
            Spacer()
            Button { model.sheet = .input(nil) } label: {
                Image(systemName: "plus").font(.system(size: 20)).frame(width: 44, height: 44)
            }.modifier(GlassControl(reduced: model.personalization.reduceTransparency))
                .accessibilityLabel("거래 입력")
        }.padding(.bottom, 6)
    }
}

@MainActor
struct RecordsScreen: View {
    @ObservedObject var model: BudgetViewModel
    @State private var calendarMode = false
    @State private var query = ""
    @State private var selectedDay = 5
    private var filtered: [BudgetTransaction] {
        model.transactions.filter { t in
            (query.isEmpty || [t.title, t.category, t.account, t.note].contains { $0.localizedCaseInsensitiveContains(query) }) &&
            (!calendarMode || Int(t.date.dropFirst(8).prefix(2)) == selectedDay)
        }
    }
    private var days: [String] { Array(Set(filtered.map(\.day))).sorted { a, b in
        if a.contains("날짜 미확인") != b.contains("날짜 미확인") { return !a.contains("날짜 미확인") }
        return a > b
    } }

    var body: some View {
        ScreenHeader(title: model.name("tab.records"), subtitle: "\(String(model.year))년 \(model.month)월 · \(model.transactions.count)건", model: model)
        Button(model.name("common.archive") + " · 전체 이력") { model.sheet = .archive }.font(.caption).frame(minHeight: 44)
        Picker("보기", selection: $calendarMode) {
            Text(model.name("records.list")).tag(false)
            Text(model.name("records.calendar")).tag(true)
        }.pickerStyle(.segmented)
        if calendarMode { calendar }
        HStack(spacing: 9) {
            Image(systemName: "magnifyingglass").foregroundStyle(Palette.muted)
            TextField("내용, 분류, 결제수단 검색", text: $query).font(.caption).autocorrectionDisabled()
        }.padding(13).background(.white, in: RoundedRectangle(cornerRadius: 12))
        if filtered.isEmpty {
            Text("찾는 기록이 없어요.\n검색어나 날짜를 바꿔보세요.")
                .font(.subheadline).foregroundStyle(Palette.muted).multilineTextAlignment(.center)
                .frame(maxWidth: .infinity).padding(.vertical, 40)
        }
        Text("표시된 기록 기준 · 확정 개인 수입·소비의 일 합계입니다. 환급·적립·이체·카드 납부·공동 거래는 합계에서 제외하고 내역에 표시합니다.")
            .font(.caption2).foregroundStyle(Palette.muted)
        ForEach(days, id: \.self) { day in
            let entries = filtered.filter { $0.day == day }
            HStack {
                VStack(alignment: .leading, spacing: 3) {
                    Text(day.contains("날짜 미확인") ? "날짜 미확인" : "\(model.month)월 \(Int(day.suffix(2)) ?? 1)일").font(.subheadline.weight(.semibold))
                    Text("\(weekday(day)) · \(entries.count)건").font(.caption2).foregroundStyle(Palette.muted)
                }
                Spacer()
                VStack(alignment: .trailing, spacing: 3) {
                    Text("수입 \(won(entries.reduce(0) { $0 + $1.recordIncome }))원").foregroundStyle(Palette.blue)
                    Text("지출 \(won(entries.reduce(0) { $0 + $1.recordExpense }))원").foregroundStyle(Palette.red)
                }.font(.caption).monospacedDigit()
            }.padding(12).background(Palette.lavender.opacity(0.45), in: RoundedRectangle(cornerRadius: 12))
            BudgetCard {
                ForEach(entries) { t in
                    Button { model.sheet = .transaction(t) } label: { TransactionRow(transaction: t) }.buttonStyle(.plain)
                    if t.id != entries.last?.id { Divider() }
                }
            }
        }
    }

    private func weekday(_ day: String) -> String {
        guard !day.contains("날짜 미확인"), let number = Int(day.suffix(2)) else { return "월만 기록" }
        var cal = Calendar(identifier: .gregorian)
        cal.timeZone = TimeZone(secondsFromGMT: 0)!
        guard let date = cal.date(from: DateComponents(year: model.year, month: model.month, day: number)) else { return "" }
        return ["일요일", "월요일", "화요일", "수요일", "목요일", "금요일", "토요일"][cal.component(.weekday, from: date) - 1]
    }
    private var calendar: some View {
        let cal = Calendar(identifier: .gregorian)
        let date = cal.date(from: DateComponents(year: model.year, month: model.month, day: 1)) ?? .now
        let offset = cal.component(.weekday, from: date) - 1
        let count = cal.range(of: .day, in: .month, for: date)?.count ?? 31
        return BudgetCard {
            LazyVGrid(columns: Array(repeating: GridItem(.flexible(), spacing: 3), count: 7), spacing: 5) {
                ForEach(["일", "월", "화", "수", "목", "금", "토"], id: \.self) { day in
                    Text(day).font(.caption2).foregroundStyle(Palette.muted).frame(height: 25)
                }
                ForEach(0..<offset, id: \.self) { _ in Color.clear.frame(height: 44) }
                ForEach(1...count, id: \.self) { day in
                    Button { selectedDay = day } label: {
                        VStack(spacing: 5) {
                            Text("\(day)").font(.caption)
                            Circle().fill(model.transactions.contains { Int($0.date.dropFirst(8).prefix(2)) == day } ? Palette.violet.opacity(0.5) : .clear)
                                .frame(width: 3, height: 3)
                        }.frame(maxWidth: .infinity, minHeight: 44)
                            .foregroundStyle(selectedDay == day ? .white : Palette.ink)
                            .background(selectedDay == day ? Palette.violet : .clear, in: RoundedRectangle(cornerRadius: 11))
                    }.accessibilityLabel("\(day)일 기록")
                }
            }
        }
    }
}

@MainActor
struct ReserveScreen: View {
    @ObservedObject var model: BudgetViewModel
    let data: BudgetSnapshot
    @State private var showAll = false
    var body: some View {
        ScreenHeader(title: model.name("tab.reserve"), subtitle: "계획보다 중요한 건, 실제로 모인 돈.", model: model)
        BudgetCard {
            Text("남은 금액 · 전체 기록 기준").font(.caption).foregroundStyle(Palette.muted)
            Text("\(won(data.reserve)) 원").font(.title.weight(.semibold)).monospacedDigit()
            Divider().padding(.vertical, 5)
            HStack {
                metric("누적 실제적립", data.reserveAllocated)
                Spacer()
                metric("누적 사용", data.reserveUsed)
            }
        }
        BudgetCard {
            AmountRow(title: "\(model.month)월 적립계획", amount: data.reserveMonthPlan)
            Divider()
            AmountRow(title: "\(model.month)월 실제적립", amount: data.reserveMonthActual)
            Text("이번달 계획은 아직 잔액에 포함되지 않았어요.").font(.caption2).foregroundStyle(Palette.muted)
        }
        SectionHeading(title: model.name("reserve.purposes"))
        Picker("예비비 보기", selection: $showAll) {
            Text(model.name("reserve.remaining")).tag(false)
            Text(model.name("reserve.all")).tag(true)
        }.pickerStyle(.segmented)
        ForEach(data.reserves.filter { showAll || $0.balance > 0 }) { r in
            Button { model.sheet = .reserve(r) } label: {
                BudgetCard {
                    HStack {
                        Text(r.title).font(.subheadline.weight(.semibold))
                        Spacer()
                        Text("\(won(r.balance))원").font(.subheadline.weight(.semibold)).monospacedDigit()
                    }
                    GeometryReader { proxy in
                        ZStack(alignment: .leading) {
                            Capsule().fill(Palette.lavender)
                            Capsule().fill(Palette.violet.opacity(0.6))
                                .frame(width: proxy.size.width * min(1, max(0, r.allocated == 0 ? 0 : Double(r.balance) / Double(r.allocated))))
                        }
                    }.frame(height: 5)
                    HStack {
                        Text(r.group)
                        Spacer()
                        Text("\(model.month)월 계획 \(won(r.plan))원")
                    }.font(.caption2).foregroundStyle(Palette.muted)
                }
            }.buttonStyle(.plain)
        }
        BudgetNote(text: "남은 금액 = 실제적립 − 사용\n사용한 돈은 실제로 쓴 달의 결산에 반영해요.")
    }
    private func metric(_ title: String, _ value: Int) -> some View {
        VStack(alignment: .leading, spacing: 8) {
            Text(title).font(.caption2).foregroundStyle(Palette.muted)
            Text("\(won(value))원").font(.subheadline.weight(.medium)).monospacedDigit()
        }
    }
}

@MainActor
struct BalanceScreen: View {
    @ObservedObject var model: BudgetViewModel
    let data: BudgetSnapshot
    var body: some View {
        ScreenHeader(title: model.name("tab.balance"), subtitle: "미래에셋 WRAP · 전체 기록 기준", model: model)
        BudgetCard {
            Text("WRAP 총잔액").font(.caption).foregroundStyle(Palette.muted)
            Text("\(won(data.wrap)) 원").font(.title.weight(.semibold)).monospacedDigit()
            GeometryReader { proxy in
                ZStack(alignment: .leading) {
                    RoundedRectangle(cornerRadius: 5).fill(Palette.lavender)
                    RoundedRectangle(cornerRadius: 5).fill(Palette.violet)
                        .frame(width: proxy.size.width * (data.wrap == 0 ? 0 : min(1, max(0, Double(data.living) / Double(data.wrap)))))
                }
            }.frame(height: 12).padding(.vertical, 8)
            HStack { Text("생활비 \(data.wrap == 0 ? 0 : Int(Double(data.living) / Double(data.wrap) * 100))%"); Spacer(); Text("목적자금 \(data.wrap == 0 ? 0 : Int(Double(data.reserve) / Double(data.wrap) * 100))%") }.font(.caption2).foregroundStyle(Palette.muted)
        }
        SectionHeading(title: model.name("balance.places"), actionTitle: model.name("balance.movement")) {
            var entry = LedgerEntry()
            entry.month = model.monthKey; entry.date = model.monthKey + "-01T12:00:00"
            entry.operation = "reserveAllocate"; entry.funding = "living"; entry.reserveID = 22
            entry.account = "미래에셋_WRAP형CMA"; entry.category = "예비비 적립"
            model.sheet = .inputEntry(entry)
        }
        Text("같은 WRAP 계좌의 돈을 생활비와 예비비 목적자금으로 나눈 구성이에요. 별도 통장 잔액을 더한 값이 아닙니다.")
            .font(.caption2).foregroundStyle(Palette.muted)
        Button { model.sheet = .living } label: {
            BudgetCard {
                AmountRow(title: model.name("balance.living"), amount: data.living, emphasized: true)
                Text("목적을 정하지 않은 생활 자금").font(.caption2).foregroundStyle(Palette.muted)
            }
        }.buttonStyle(.plain)
        BudgetCard {
            AmountRow(title: model.name("balance.purposes"), amount: data.reserve, emphasized: true)
            ForEach(["가족행사비", "여행비", "상환비", "경조사비"], id: \.self) { group in
                Divider()
                Button { model.sheet = .purpose(group) } label: {
                    AmountRow(title: group, amount: data.reserves.filter { $0.group == group }.reduce(0) { $0 + $1.balance })
                }.buttonStyle(.plain)
            }
        }
        SectionHeading(title: model.name("balance.recent"))
        BudgetCard {
            ForEach(Array(model.wrapTransactions.prefix(5))) { t in
                Button { model.sheet = .transaction(t) } label: { TransactionRow(transaction: t) }.buttonStyle(.plain)
                if t.id != model.wrapTransactions.prefix(5).last?.id { Divider() }
            }
        }
        BudgetNote(text: "생활비와 목적자금 사이의 이동은 총잔액을 바꾸지 않아요. 외부로 보낸 돈은 총잔액에서 빠져요.")
    }
}

#Preview { ContentView() }
