import Foundation
import JavaScriptCore

struct LedgerError: LocalizedError {
    let message: String
    var errorDescription: String? { message }
}

// Native SwiftUI uses exactly the engine exercised by tests/ledger-engine.test.js.
@MainActor
final class LedgerCoreBridge {
    private let context: JSContext
    init() throws {
        guard let context = JSContext() else {
            throw LedgerError(message: "계산 모듈을 찾을 수 없습니다.")
        }
        self.context = context
        for name in ["ledger-engine", "xlsx.full.min", "excel-importer"] {
            guard let url = Bundle.main.url(forResource: name, withExtension: "js") else {
                throw LedgerError(message: "계산·엑셀 모듈을 찾을 수 없습니다.")
            }
            context.evaluateScript(try String(contentsOf: url, encoding: .utf8))
            guard context.exception == nil else { throw LedgerError(message: "계산·엑셀 모듈을 시작할 수 없습니다.") }
        }
        guard context.objectForKeyedSubscript("ledgerDispatch")?.isObject == true,
              context.objectForKeyedSubscript("excelDispatch")?.isObject == true else {
            throw LedgerError(message: "계산 모듈을 시작할 수 없습니다.")
        }
    }
    func call(_ action: String, state: Data, month: String? = nil, change: [String: Any]? = nil) throws -> Data {
        let object = try JSONSerialization.jsonObject(with: state)
        var request: [String: Any] = ["action": action, "state": object]
        if let month { request["month"] = month }
        if let change { request["change"] = change }
        let json = try JSONSerialization.data(withJSONObject: request, options: [.sortedKeys])
        guard let input = String(data: json, encoding: .utf8) else { throw LedgerError(message: "JSON 인코딩 오류입니다.") }
        context.exception = nil
        guard let result = context.objectForKeyedSubscript("ledgerDispatch")?.call(withArguments: [input])?.toString(),
              context.exception == nil, let bytes = result.data(using: .utf8),
              let response = try JSONSerialization.jsonObject(with: bytes) as? [String: Any] else {
            throw LedgerError(message: "계산 결과를 읽을 수 없습니다.")
        }
        guard response["ok"] as? Bool == true, let value = response["value"] else {
            throw LedgerError(message: response["error"] as? String ?? "계산에 실패했습니다.")
        }
        return try JSONSerialization.data(withJSONObject: value, options: [.sortedKeys])
    }
    func spreadsheet(_ action: String, state: Data, payload: [String: Any]) throws -> Data {
        var request = payload
        request["action"] = action
        request["state"] = try JSONSerialization.jsonObject(with: state)
        let bytes = try JSONSerialization.data(withJSONObject: request, options: [.sortedKeys])
        guard let input = String(data: bytes, encoding: .utf8) else { throw LedgerError(message: "엑셀 요청 인코딩 오류입니다.") }
        context.exception = nil
        guard let result = context.objectForKeyedSubscript("excelDispatch")?.call(withArguments: [input])?.toString(),
              context.exception == nil, let data = result.data(using: .utf8),
              let response = try JSONSerialization.jsonObject(with: data) as? [String: Any] else {
            throw LedgerError(message: "엑셀 결과를 읽을 수 없습니다.")
        }
        guard response["ok"] as? Bool == true, let value = response["value"] else {
            throw LedgerError(message: response["error"] as? String ?? "엑셀을 가져오지 못했습니다.")
        }
        return try JSONSerialization.data(withJSONObject: value, options: [.sortedKeys])
    }
}

struct LedgerEntry: Codable, Identifiable {
    var id: Int = 0
    var date: String? = nil
    var month = "2026-10"
    var title = ""
    var account = "개인 계좌"
    var category = "식비"
    var amount: Int = 0
    var operation = "expense"
    var status = "confirmed"
    var funding = "bank"
    var destination = "bank"
    var expenseClass = "variable"
    var reserveID: Int? = nil
    var linkedID: Int? = nil
    var attributedMonth: String? = nil
    var alreadyNet = false
    var displayAmount: Int? = nil
    var planID: String? = nil
    var note = ""
    var sourceKey: String? = nil
    var sourceKind: String? = nil
    func dictionary() throws -> [String: Any] {
        var result = try JSONSerialization.jsonObject(with: JSONEncoder().encode(self)) as? [String: Any] ?? [:]
        for key in ["date", "reserveID", "linkedID", "attributedMonth", "displayAmount", "planID", "sourceKey", "sourceKind"] where result[key] == nil {
            result[key] = NSNull()
        }
        return result
    }
}

struct LedgerPlan: Codable, Identifiable {
    var id, month, type, title, account, status: String
    var amount: Int
}

struct LedgerMetadata: Decodable {
    let transactions: [LedgerEntry]
    let sourceArchive: [RawRecord]
    let plans: [LedgerPlan]
    let personalization: Personalization?
}
