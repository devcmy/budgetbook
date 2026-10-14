import SwiftUI

enum Palette {
    static let paper = Color(hex: 0xF7F8FA)
    static let ink = Color(hex: 0x202B3A)
    static let muted = Color(hex: 0x6E7886)
    static let line = Color(hex: 0xEDF0F4)
    static let violet = Color(hex: 0x7662D8)
    static let lavender = Color(hex: 0xEEEAFB)
    static let navy = Color(hex: 0x253248)
    static let green = Color(hex: 0x147D69)
    static let blue = Color(hex: 0x3B6DAF)
    static let red = Color(hex: 0xAD4848)
}

extension Color {
    init(hex: UInt32) {
        self.init(.sRGB, red: Double((hex >> 16) & 255) / 255,
                  green: Double((hex >> 8) & 255) / 255, blue: Double(hex & 255) / 255, opacity: 1)
    }
}

struct BudgetCard<Content: View>: View {
    @Environment(\.budgetCompact) private var compact
    let content: Content
    init(@ViewBuilder content: () -> Content) { self.content = content() }
    var body: some View {
        VStack(alignment: .leading, spacing: compact ? 8 : 12) { content }
            .frame(maxWidth: .infinity, alignment: .leading).padding(compact ? 14 : 18)
            .background(.white, in: RoundedRectangle(cornerRadius: 20))
            .overlay(RoundedRectangle(cornerRadius: 20).stroke(Palette.line, lineWidth: 1))
    }
}

struct StatusBadge: View {
    let text: String
    var color: Color { text.hasPrefix("예상") ? Palette.red : text == "미확인" ? Palette.blue : text == "입력 시안" ? Palette.violet : Palette.green }
    var body: some View {
        Text(text).font(.caption2.weight(.semibold)).foregroundStyle(color)
            .padding(.horizontal, 7).padding(.vertical, 4)
            .background(color.opacity(0.08), in: RoundedRectangle(cornerRadius: 6))
    }
}

struct SectionHeading: View {
    let title: String
    var actionTitle: String? = nil
    var action: (() -> Void)? = nil
    var body: some View {
        HStack {
            Text(title).font(.subheadline.weight(.semibold))
            Spacer()
            if let actionTitle, let action {
                Button(action: action) {
                    HStack(spacing: 3) { Text(actionTitle); Image(systemName: "chevron.right").font(.system(size: 9)) }
                        .font(.caption).foregroundStyle(Palette.muted).frame(minHeight: 44)
                }
            }
        }.padding(.top, 5)
    }
}

struct AmountRow: View {
    let title: String
    let amount: Int
    var emphasized = false
    var body: some View {
        HStack(alignment: .firstTextBaseline, spacing: 12) {
            Text(title).font(.subheadline)
            Spacer(minLength: 8)
            Text("\(won(amount))원").font(.subheadline).monospacedDigit()
        }.fontWeight(emphasized ? .semibold : .regular).frame(minHeight: 32)
    }
}

struct BudgetNote: View {
    let text: String
    var body: some View {
        Text(text).font(.caption).foregroundStyle(Palette.muted).lineSpacing(5)
            .frame(maxWidth: .infinity, alignment: .leading).padding(13)
            .background(Palette.lavender.opacity(0.6), in: RoundedRectangle(cornerRadius: 13))
    }
}

struct TileMark: View {
    var body: some View {
        LazyVGrid(columns: Array(repeating: GridItem(.flexible(), spacing: 3), count: 4), spacing: 3) {
            ForEach(0..<16) { i in
                RoundedRectangle(cornerRadius: 3)
                    .fill(i % 3 == 0 ? Palette.violet.opacity(0.7) : Palette.violet.opacity(0.22))
                    .aspectRatio(1, contentMode: .fit)
            }
        }.frame(width: 65, height: 65).rotationEffect(.degrees(-6)).accessibilityHidden(true)
    }
}

struct TransactionRow: View {
    @Environment(\.budgetCompact) private var compact
    let transaction: BudgetTransaction
    var body: some View {
        HStack(spacing: 11) {
            Image(systemName: transaction.symbol).font(.system(size: 16))
                .foregroundStyle(transaction.kind == "지출" ? Palette.violet : transaction.tint)
                .frame(width: 37, height: 37)
                .background(Palette.lavender.opacity(0.5), in: RoundedRectangle(cornerRadius: 13))
            VStack(alignment: .leading, spacing: 5) {
                Text(transaction.title).font(.subheadline.weight(.medium)).lineLimit(1)
                Text("\(transaction.category) · \(transaction.account)\(transaction.isPreview ? " · 입력 시안" : "")")
                    .font(.system(size: 10)).foregroundStyle(Palette.muted).lineLimit(1)
            }
            Spacer(minLength: 4)
            Text("\(transaction.kind == "수입" && transaction.amount >= 0 ? "+" : "")\(won(transaction.amount))원")
                .font(.caption.weight(.semibold)).monospacedDigit().foregroundStyle(transaction.tint)
        }.frame(minHeight: compact ? 44 : 50).contentShape(Rectangle())
    }
}

struct PrimaryButtonStyle: ButtonStyle {
    func makeBody(configuration: Configuration) -> some View {
        configuration.label.font(.subheadline.weight(.semibold)).foregroundStyle(.white)
            .frame(maxWidth: .infinity, minHeight: 48)
            .background(Palette.violet, in: RoundedRectangle(cornerRadius: 14))
            .opacity(configuration.isPressed ? 0.7 : 1)
    }
}

private struct BudgetCompactKey: EnvironmentKey { static let defaultValue = true }
extension EnvironmentValues {
    var budgetCompact: Bool {
        get { self[BudgetCompactKey.self] }
        set { self[BudgetCompactKey.self] = newValue }
    }
}

// Build with Xcode 26+ SDK. iOS 27 uses the system's current Glass appearance.
struct GlassControl: ViewModifier {
    let reduced: Bool
    @Environment(\.accessibilityReduceTransparency) private var systemReduced
    @ViewBuilder func body(content: Content) -> some View {
        if reduced || systemReduced {
            content.background(Color.white, in: Capsule()).overlay(Capsule().stroke(Palette.line))
        } else if #available(iOS 26.0, *) {
            content.glassEffect(.regular.interactive(), in: Capsule())
        } else {
            content.background(.regularMaterial, in: Capsule())
        }
    }
}

struct TabTransparency: ViewModifier {
    let reduced: Bool
    @ViewBuilder func body(content: Content) -> some View {
        if reduced {
            content.toolbarBackground(Palette.paper, for: .tabBar).toolbarBackground(.visible, for: .tabBar)
        } else { content }
    }
}
