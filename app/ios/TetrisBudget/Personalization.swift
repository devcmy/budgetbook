import Foundation
import SwiftUI

struct InterfaceLabel: Decodable, Identifiable {
    let id, defaultName, group: String
    let maxLength: Int
}

struct Personalization: Codable {
    var names: [String: String] = [:]
    var compact = true
    var reduceTransparency = false
    static let storageKey = "tetris-budget-ui-v2"
    static func load(catalog: [InterfaceLabel]) -> Personalization {
        guard let data = UserDefaults.standard.data(forKey: storageKey),
              var settings = try? JSONDecoder().decode(Self.self, from: data) else { return Self() }
        settings.names = settings.names.filter { key, value in
            guard let label = catalog.first(where: { $0.id == key }) else { return false }
            return validName(value, limit: label.maxLength) != nil
        }
        return settings
    }
    func save() {
        if let data = try? JSONEncoder().encode(self) { UserDefaults.standard.set(data, forKey: Self.storageKey) }
    }
    static func validName(_ raw: String, limit: Int) -> String? {
        let text = raw.trimmingCharacters(in: .whitespacesAndNewlines)
        guard !text.isEmpty, text.unicodeScalars.count <= limit,
              !text.unicodeScalars.contains(where: { CharacterSet.controlCharacters.contains($0) }) else { return nil }
        return text
    }
}

@MainActor
struct SettingsScreen: View {
    @ObservedObject var model: BudgetViewModel
    var menuGroups: [String] {
        model.labels.filter { !$0.id.hasPrefix("tab.") }.reduce(into: []) { result, label in
            if !result.contains(label.group) { result.append(label.group) }
        }
    }
    var body: some View {
        HStack {
            VStack(alignment: .leading, spacing: 5) {
                Text(model.name("tab.settings")).font(.title2.weight(.bold))
                Text("내 가계부는, 내 이름으로.").font(.caption2).foregroundStyle(Palette.muted)
            }
            Spacer()
            TileMark().scaleEffect(0.55).frame(width: 40, height: 40)
        }.padding(.bottom, 7)
        BudgetCard {
            Text("나의 테트리스 가계부").font(.subheadline.weight(.semibold))
            Text("iPhone 17 Pro Max · 개인용").font(.caption2).foregroundStyle(Palette.muted)
        }
        BackupControls(model: model)
        SectionHeading(title: model.name("settings.plans"), actionTitle: "추가") { model.sheet = .plan(nil) }
        MonthlyPlanOverview(model: model)
        SectionHeading(title: "탭 이름")
        BudgetCard {
            ForEach(BudgetTab.allCases, id: \.self) { tab in
                Button { model.sheet = .rename(tab.labelKey) } label: {
                    HStack(spacing: 12) {
                        Image(systemName: tab.symbol).foregroundStyle(Palette.violet).frame(width: 31, height: 31)
                            .background(Palette.lavender, in: RoundedRectangle(cornerRadius: 10))
                        VStack(alignment: .leading, spacing: 5) {
                            Text(model.name(tab.labelKey)).font(.subheadline.weight(.medium))
                            Text("기본 이름 · \(model.defaultName(tab.labelKey))").font(.caption2).foregroundStyle(Palette.muted)
                        }
                        Spacer()
                        Image(systemName: "chevron.right").font(.caption2).foregroundStyle(Palette.muted)
                    }.frame(minHeight: 44).contentShape(Rectangle())
                }.buttonStyle(.plain)
                if tab != .settings { Divider() }
            }
        }
        SectionHeading(title: "메뉴 이름")
        ForEach(menuGroups, id: \.self) { group in
            BudgetCard {
                DisclosureGroup(group) {
                    ForEach(model.labels.filter { $0.group == group }) { label in
                        Button { model.sheet = .rename(label.id) } label: {
                            HStack {
                                Text(label.defaultName).font(.caption).foregroundStyle(Palette.muted)
                                Spacer()
                                Text(model.name(label.id)).font(.caption.weight(.medium)).lineLimit(1)
                                Image(systemName: "chevron.right").font(.system(size: 9))
                            }.frame(minHeight: 44).contentShape(Rectangle())
                        }.buttonStyle(.plain)
                    }
                }.font(.subheadline.weight(.medium))
            }
        }
        SectionHeading(title: "화면")
        BudgetCard {
            Toggle(isOn: $model.personalization.compact) {
                VStack(alignment: .leading, spacing: 5) {
                    Text("촘촘하게 보기").font(.subheadline)
                    Text("여백을 줄이고 더 많은 내역을 봐요.").font(.caption2).foregroundStyle(Palette.muted)
                }
            }
            Divider()
            Toggle(isOn: $model.personalization.reduceTransparency) {
                VStack(alignment: .leading, spacing: 5) {
                    Text("투명 효과 줄이기").font(.subheadline)
                    Text("유리 배경을 더 선명하게 표시해요.").font(.caption2).foregroundStyle(Palette.muted)
                }
            }
        }
        HStack {
            Button("이름을 기본값으로 복원") { model.sheet = .resetNames }
            Spacer()
            Button(model.name("common.about")) { model.sheet = .about }
        }.font(.caption).frame(minHeight: 44)
        Text("이름을 바꿔도 원본 분류와 계산 기준은 유지돼요.\n이름과 화면 설정은 이 기기에 저장합니다.")
            .font(.caption2).foregroundStyle(Palette.muted).lineSpacing(5)
    }
}

@MainActor
struct RenameSheet: View {
    @ObservedObject var model: BudgetViewModel
    let labelID: String
    @State private var name: String
    @State private var error = ""
    init(model: BudgetViewModel, labelID: String) {
        self.model = model
        self.labelID = labelID
        _name = State(initialValue: model.name(labelID))
    }
    var body: some View {
        VStack(alignment: .leading, spacing: 17) {
            Text("기본 이름 · \(model.defaultName(labelID))").font(.caption).foregroundStyle(Palette.muted)
            TextField("표시할 이름", text: $name).font(.title3).padding(14)
                .background(.white, in: RoundedRectangle(cornerRadius: 13)).autocorrectionDisabled()
                .accessibilityLabel("표시할 이름")
            Text("\(model.labels.first(where: { $0.id == labelID })?.maxLength ?? 20)자 이내 · 화면의 이름만 바뀌어요.")
                .font(.caption2).foregroundStyle(Palette.muted)
            if !error.isEmpty { Text(error).font(.caption).foregroundStyle(Palette.red) }
            Button("이름 저장") {
                let limit = model.labels.first(where: { $0.id == labelID })?.maxLength ?? 20
                guard let text = Personalization.validName(name, limit: limit) else {
                    error = "1~\(limit)자 이름을 입력해주세요."; return
                }
                model.personalization.names[labelID] = text
                model.sheet = nil
                model.showToast("이름을 저장했어요.")
            }.buttonStyle(PrimaryButtonStyle())
            Button("이 이름만 기본값으로") {
                model.personalization.names.removeValue(forKey: labelID)
                model.sheet = nil
                model.showToast("기본 이름으로 되돌렸어요.")
            }.font(.caption).frame(minHeight: 44)
        }
    }
}

@MainActor
private struct MonthlyPlanOverview: View {
    @ObservedObject var model: BudgetViewModel
    private let types = [("income", "수입 예상"), ("fixed", "고정비 예상"), ("reserve", "예비비 적립계획"), ("familyContribution", "가족계좌 납입 · 중단")]
    var body: some View {
        Text("\(model.year)년 \(model.month)월 · 등록 계획액입니다. 실제 잔액은 바뀌지 않아요. 종류를 눌러 항목을 확인·수정하세요.")
            .font(.caption2).foregroundStyle(Palette.muted)
        ForEach(types, id: \.0) { type in
            let items = model.plans.filter { $0.month == model.monthKey && $0.type == type.0 }
            if type.0 != "familyContribution" || !items.isEmpty {
                BudgetCard {
                    DisclosureGroup {
                        ForEach(items) { plan in
                            Button { model.sheet = .plan(plan) } label: {
                                VStack(alignment: .leading, spacing: 4) {
                                    AmountRow(title: plan.title, amount: type.0 == "familyContribution" ? 0 : plan.amount)
                                    let linked = model.entries.filter { $0.planID == plan.id }
                                    let actual = linked.filter { $0.status == "confirmed" }.reduce(0) { $0 + $1.amount }
                                    Text(plan.account + " · " + plan.status + (linked.isEmpty ? "" : " · 실제 연결 \(won(actual))원"))
                                        .font(.caption2).foregroundStyle(Palette.muted)
                                }.padding(.vertical, 6)
                            }.buttonStyle(.plain)
                        }
                        if items.isEmpty { Text("등록한 항목이 없어요. 추가 버튼으로 등록하세요.").font(.caption).foregroundStyle(Palette.muted) }
                    } label: {
                        HStack {
                            VStack(alignment: .leading, spacing: 5) {
                                Text(type.1).font(.subheadline.weight(.semibold))
                                Text("\(items.count)개 항목 · 미확인 \(items.filter { $0.status == "미확인" }.count)개")
                                    .font(.caption2).foregroundStyle(Palette.muted)
                            }
                            Spacer()
                            Text("\(won(type.0 == "familyContribution" ? 0 : items.reduce(0) { $0 + $1.amount }))원")
                                .font(.title3.weight(.semibold)).foregroundStyle(Palette.violet).monospacedDigit()
                        }.padding(.vertical, 5)
                    }
                }
            }
        }
    }
}
