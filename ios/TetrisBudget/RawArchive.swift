import SwiftUI

struct RawRecord: Decodable, Identifiable {
    let id: Int
    let date, month, account, category, subcategory, title, kind, additional, classification, method, note: String?
    let amount: Int?
    let burden: Double?
    let importFile, importSheet: String?
    let importRow: Int?
    let importKey: String?
    enum CodingKeys: String, CodingKey {
        case id = "source_row", date = "일시", month = "월", account = "자산", category = "분류", subcategory = "소분류"
        case title = "내용", amount = "금액(KRW)", kind = "수입/지출", additional = "추가입력"
        case classification = "재분류", burden = "본인부담", method = "결제구분", note = "비고"
        case importFile = "import_file", importSheet = "import_sheet", importRow = "import_row"
        case importKey
    }
    var fields: [(String, String)] {
        var result: [(String, String)] = [(importRow == nil ? "원본 행" : "보관 ID", String(id)), ("일시", date ?? "날짜 미확인"), ("월", month ?? "—"),
         ("자산", account ?? "—"), ("분류", category ?? "—"), ("소분류", subcategory ?? "—"),
         ("내용", title ?? "—"), ("금액(KRW)", amount.map(won) ?? "—"), ("수입/지출", kind ?? "—"),
         ("추가입력", additional ?? "—"), ("재분류", classification ?? "—"),
         ("본인부담 · 원문", burden.map { $0.formatted(.number.locale(Locale(identifier: "ko_KR")).precision(.fractionLength(0...6))) } ?? "—"),
         ("결제구분", method ?? "—"), ("비고", note ?? "—")]
        if let importFile { result.append(("가져온 파일", importFile)) }
        if let importSheet { result.append(("엑셀 시트", importSheet)) }
        if let importRow { result.append(("엑셀 행", String(importRow))) }
        return result
    }
}

@MainActor
struct RawArchiveScreen: View {
    @ObservedObject var model: BudgetViewModel
    @State private var query = ""
    @State private var allMonths = false
    private var records: [RawRecord] {
        model.rawArchive.filter { r in
            (allMonths || (r.month ?? r.date ?? "").hasPrefix(model.monthKey)) &&
            (query.isEmpty || [r.title, r.account, r.classification, r.additional, r.note].compactMap { $0 }.contains { $0.localizedCaseInsensitiveContains(query) })
        }.sorted { ($0.date ?? "") > ($1.date ?? "") }
    }
    var body: some View {
        VStack(alignment: .leading, spacing: 12) {
            BudgetNote(text: "원본 RAW 보관함 · 읽기 전용\n당시의 분류·본인부담을 그대로 보관합니다. 과거 공동계좌 50% 분담 등 원문 규칙은 현재 앱 계산에 적용하지 않아요.")
            Toggle("전체 기간 보기", isOn: $allMonths)
            TextField("내용·계좌·분류 검색", text: $query).padding(12).background(.white, in: RoundedRectangle(cornerRadius: 12))
                .autocorrectionDisabled()
            Text("\(records.count)건 · 원문은 백업에도 포함돼요.").font(.caption2).foregroundStyle(Palette.muted)
            LazyVStack(spacing: 8) {
                ForEach(records) { record in
                    BudgetCard {
                        DisclosureGroup {
                            ForEach(record.fields, id: \.0) { field in
                                VStack(alignment: .leading, spacing: 3) {
                                    Text(field.0).font(.caption2).foregroundStyle(Palette.muted)
                                    Text(field.1).font(.caption).textSelection(.enabled)
                                }.frame(maxWidth: .infinity, alignment: .leading).padding(.vertical, 3)
                            }
                        } label: {
                            VStack(alignment: .leading, spacing: 5) {
                                Text(record.title ?? "내용 없음").font(.subheadline.weight(.medium))
                                Text("\(record.date ?? record.month ?? "날짜 미확인") · \(record.amount.map(won) ?? "—")원")
                                    .font(.caption2).foregroundStyle(Palette.muted)
                            }
                        }
                    }
                }
            }
        }
    }
}
