Attribute VB_Name = "BBFixed"
Option Explicit

Private Const META As String = "_RAW자동연동"
Private Const FIRST_META As Long = 32002
Private Const RAW_LAST As Long = 4001
Private Const SNAPSHOT As Long = 33000
Private busy As Boolean

Public Sub BB_SnapshotRaw()
    Dim raw As Worksheet, metaSheet As Worksheet, data As Variant, saved As Variant
    Dim r As Long, ident As Long
    Set raw = ThisWorkbook.Worksheets("RAW")
    Set metaSheet = ThisWorkbook.Worksheets(META)
    raw.Calculate
    data = raw.Range("A2:AQ4001").Value2
    ReDim saved(1 To 4000, 1 To 2)
    For r = 1 To 4000
        ident = CLng(data(r, 42))
        saved(ident - 1, 1) = Number(data(r, 7))
        saved(ident - 1, 2) = Number(data(r, 11))
    Next r
    metaSheet.Range("CG33002:CH37001").Value2 = saved
End Sub

Private Sub SnapshotOne(ByVal raw As Worksheet, ByVal r As Long)
    Dim ident As Long
    ident = CLng(raw.Cells(r, 42).Value2)
    ThisWorkbook.Worksheets(META).Cells(SNAPSHOT + ident, 85).Value2 = Number(raw.Cells(r, 7).Value2)
    ThisWorkbook.Worksheets(META).Cells(SNAPSHOT + ident, 86).Value2 = Number(raw.Cells(r, 11).Value2)
End Sub

Private Sub RawFixedChange(ByVal raw As Worksheet, ByVal target As Range)
    Dim hit As Range, r As Long, ident As Long, oldGross As Double, oldBurden As Double
    Dim ledger As Worksheet, ledgerRow As Long, ledgerOld As Variant, failed As String
    Dim savedEvents As Boolean, savedCalc As XlCalculation, ownAmount As Variant
    Set hit = Intersect(target, raw.Range("G2:G4001"))
    If hit Is Nothing Then Exit Sub
    If hit.CountLarge <> 1 Then Exit Sub
    r = hit.Row: ident = CLng(raw.Cells(r, 42).Value2)
    oldGross = Number(ThisWorkbook.Worksheets(META).Cells(SNAPSHOT + ident, 85).Value2)
    oldBurden = Number(ThisWorkbook.Worksheets(META).Cells(SNAPSHOT + ident, 86).Value2)
    savedEvents = Application.EnableEvents: savedCalc = Application.Calculation
    On Error GoTo Failed
    busy = True: Application.EnableEvents = False: Application.Calculation = xlCalculationManual
    raw.Calculate
    If CStr(raw.Cells(r, 20).Value2) <> "고정비" Then GoTo CleanUp
    If Not IsNumeric(raw.Cells(r, 7).Value2) Then Err.Raise vbObjectError + 119, , "고정비 금액은 숫자로 입력해 주세요."
    Set ledger = ThisWorkbook.Worksheets("잔액")
    ledgerRow = LinkedLedgerRow(ledger, ident)
    If ledgerRow > 0 Then ledgerOld = ledger.Range("A" & ledgerRow & ":Y" & ledgerRow).Formula2
    If Intersect(target, raw.Cells(r, 11)) Is Nothing Then
        If Not raw.Cells(r, 11).HasFormula And Abs(oldGross - oldBurden) < 0.0001 And _
           Abs(Number(raw.Cells(r, 11).Value2) - oldBurden) < 0.0001 Then
            raw.Cells(r, 11).Value2 = Number(raw.Cells(r, 7).Value2)
        End If
    End If
    raw.Calculate
    SyncCash raw, r, ledger, ledgerRow, ledgerOld, oldGross
    ValidateCash ledger, ledgerRow
    Application.Calculate
CleanUp:
    SnapshotOne raw, r
    Application.Calculation = savedCalc: Application.EnableEvents = savedEvents: busy = False
    Exit Sub
Failed:
    failed = Err.Description
    On Error Resume Next
    raw.Cells(r, 7).Value2 = oldGross
    If Not raw.Cells(r, 11).HasFormula Then raw.Cells(r, 11).Value2 = oldBurden
    If ledgerRow > 0 Then ledger.Range("A" & ledgerRow & ":Y" & ledgerRow).Formula2 = ledgerOld
    Application.Calculate
    MsgBox failed, vbExclamation, "RAW 고정비 수정"
    Resume CleanUp
End Sub

Public Sub BB_FixedEditable(ByVal enabled As Boolean)
    Dim sh As Worksheet, m As Long, slot As Long
    For m = 1 To 12
        Set sh = ThisWorkbook.Worksheets(CStr(m) & "월")
        sh.Unprotect
        For slot = 4 To 26
            If slot <> 25 Then
                sh.Cells(slot, 15).Locked = Not enabled
                sh.Range(sh.Cells(slot, 22), sh.Cells(slot, 27)).Locked = Not enabled
            End If
        Next slot
        sh.Protect UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True, _
                   AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
    Next m
End Sub

Public Sub BB_FixedChange(ByVal sh As Object, ByVal target As Range)
    Dim m As Long, slot As Long, col As Long, metaRow As Long, hit As Range, cell As Range
    Dim typed As Variant, oldTotal As Double, desired As Double, count As Long
    Dim raw As Worksheet, ledger As Worksheet, candidates As Collection, chosen As Variant
    Dim rawRow As Long, ident As Long, isNew As Boolean, choice As VbMsgBoxResult
    Dim oldRaw As Variant, ledgerOld As Variant, ledgerRow As Long, oldGross As Double
    Dim newDate As Variant, newAsset As Variant, newTitle As Variant, proposed As Variant
    Dim savedEvents As Boolean, savedCalc As XlCalculation, changed As Boolean
    Dim failed As String, savedState As Boolean, fullBurden As Boolean, portion As Double
    If busy Then Exit Sub
    If Not TypeOf sh Is Worksheet Then Exit Sub
    If sh.Name = "RAW" Then RawFixedChange sh, target: Exit Sub
    m = Val(sh.Name)
    If m < 1 Or m > 12 Then Exit Sub
    If sh.Name <> CStr(m) & "월" Then Exit Sub
    Set hit = Intersect(target, Union(sh.Range("O4:O24"), sh.Range("O26"), _
                                     sh.Range("V4:V24"), sh.Range("V26")))
    If hit Is Nothing Then Exit Sub
    savedEvents = Application.EnableEvents
    savedCalc = Application.Calculation
    savedState = True
    On Error GoTo Failed
    busy = True
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    If hit.CountLarge <> 1 Or target.CountLarge > hit.Cells(1).MergeArea.CountLarge Then
        For Each cell In hit.Cells
            RestoreCell sh, cell.Row, cell.Column
        Next cell
        MsgBox "청구액과 메모는 한 칸씩 수정해 주세요. 여러 칸 붙여넣기는 RAW에서 할 수 있습니다.", vbInformation
        GoTo CleanUp
    End If
    Set cell = hit.Cells(1)
    slot = cell.Row: col = cell.Column
    typed = cell.Value2
    RestoreCell sh, slot, col
    cell.Calculate
    Set raw = ThisWorkbook.Worksheets("RAW")
    Set ledger = ThisWorkbook.Worksheets("잔액")
    raw.Calculate
    sh.Cells(slot, 15).Calculate
    oldTotal = Number(sh.Cells(slot, 15).Value2)
    If col = 15 Then
        If IsEmpty(typed) Or Len(CStr(typed)) = 0 Then typed = 0
        If Not IsNumeric(typed) Then Err.Raise vbObjectError + 120, , "청구액은 숫자로 입력해 주세요."
        desired = CDbl(typed)
        If desired < 0 Or desired > 1000000000000# Or desired <> Fix(desired) Then _
            Err.Raise vbObjectError + 121, , "청구액은 0원 이상의 정수로 입력해 주세요."
    Else
        If IsError(typed) Then Err.Raise vbObjectError + 122, , "메모에 오류 값을 입력할 수 없습니다."
    End If
    Set candidates = FixedRows(raw, m, slot)
    count = candidates.Count
    If count = 1 Then
        ident = CLng(candidates(1))
    ElseIf count > 1 Then
        ident = ChooseID(raw, candidates, "여러 거래가 합산된 항목입니다. 수정할 거래ID를 선택하세요.")
        If ident = 0 Then GoTo CleanUp
    Else
        If col = 22 Then
            MsgBox "연결된 실제 거래가 없습니다. 청구액을 입력해 RAW를 연결·등록한 뒤 메모를 수정해 주세요.", vbInformation
            GoTo CleanUp
        End If
        choice = MsgBox("연결된 실제 RAW 거래가 없습니다." & vbCrLf & _
                       "예: 실제 거래 연결·등록" & vbCrLf & _
                       "아니오: 예정액만 수정" & vbCrLf & "취소: 입력 취소", vbYesNoCancel + vbQuestion)
        If choice = vbCancel Then GoTo CleanUp
        If choice = vbNo Then
            If m < 11 Then
                MsgBox "이전 월에는 예정 입력칸이 없습니다. 실제 거래를 연결하거나 RAW에서 확인해 주세요.", vbInformation
            Else
                sh.Cells(slot, 28).Value2 = desired
            End If
            GoTo CleanUp
        End If
        chosen = Application.InputBox("연결할 RAW 거래ID를 입력하세요. 새 실제 거래를 등록하려면 0을 입력하세요.", _
                                      "RAW 연결", Type:=1)
        If IsCancelled(chosen) Then GoTo CleanUp
        If chosen <> Fix(chosen) Or chosen < 0 Then Err.Raise vbObjectError + 123, , "거래ID를 확인해 주세요."
        ident = CLng(chosen)
        isNew = (ident = 0)
    End If
    If isNew Then
        rawRow = EmptyRawRow(raw)
        If rawRow = 0 Then Err.Raise vbObjectError + 124, , "RAW 빈 입력행이 없습니다."
        ident = CLng(raw.Cells(rawRow, 42).Value2)
        newDate = Application.InputBox("실제 거래일을 입력하세요. 예: 2026-11-16", "실제 거래일", Type:=2)
        If IsCancelled(newDate) Then GoTo CleanUp
        If Not IsDate(newDate) Then Err.Raise vbObjectError + 125, , "날짜를 확인해 주세요."
        newDate = CDate(newDate)
        If Year(newDate) <> 2026 Or Month(newDate) <> m Then _
            Err.Raise vbObjectError + 126, , "선택한 월의 실제 거래일을 입력해 주세요."
        newAsset = Application.InputBox("실제 결제한 카드·계좌명을 입력하세요. CASH 대신 실제 계좌명을 적어 주세요.", _
                      "결제수단", LastField(raw, slot, 3), Type:=2)
        If IsCancelled(newAsset) Then GoTo CleanUp
        If Len(Trim$(CStr(newAsset))) = 0 Or UCase$(Trim$(CStr(newAsset))) = "CASH" Then _
            Err.Raise vbObjectError + 127, , "실제 카드·계좌명이 필요합니다."
        newTitle = Application.InputBox("RAW에 기록할 실제 거래내용을 입력하세요.", "거래내용", _
                      LastField(raw, slot, 6), Type:=2)
        If IsCancelled(newTitle) Then GoTo CleanUp
        If Len(Trim$(CStr(newTitle))) = 0 Then Err.Raise vbObjectError + 128, , "거래내용이 필요합니다."
    Else
        rawRow = RowForID(raw, ident)
        If rawRow = 0 Then Err.Raise vbObjectError + 129, , "거래ID를 찾을 수 없습니다."
        If Not IsDate(raw.Cells(rawRow, 1).Value) Then Err.Raise vbObjectError + 130, , "RAW 날짜를 확인해 주세요."
        If Year(raw.Cells(rawRow, 1).Value) <> 2026 Or Month(raw.Cells(rawRow, 1).Value) <> m Then _
            Err.Raise vbObjectError + 131, , "다른 월의 거래를 연결할 수 없습니다."
        If CStr(raw.Cells(rawRow, 21).Value2) <> "본인" Or Number(raw.Cells(rawRow, 17).Value2) <> 1 Then _
            Err.Raise vbObjectError + 132, , "본인 거래이며 중복이 아닌 내역만 연결할 수 있습니다."
        If CStr(raw.Cells(rawRow, 8).Value2) <> "지출" And _
           Not (CStr(raw.Cells(rawRow, 20).Value2) = "고정비" And Number(raw.Cells(rawRow, 27).Value2) = slot) Then _
            Err.Raise vbObjectError + 133, , "지출 거래만 고정비로 연결할 수 있습니다."
        If count = 0 And CStr(raw.Cells(rawRow, 20).Value2) = "고정비" And _
           Number(raw.Cells(rawRow, 27).Value2) <> slot Then _
            Err.Raise vbObjectError + 136, , "이미 다른 고정비 항목에 연결된 거래입니다. 기존 연결 항목에서 수정하거나 RAW 연결코드를 먼저 확인해 주세요."
    End If
    oldRaw = raw.Range("A" & rawRow & ":AQ" & rawRow).Formula2
    oldGross = Number(raw.Cells(rawRow, 7).Value2)
    ledgerRow = LinkedLedgerRow(ledger, ident)
    If ledgerRow > 0 Then ledgerOld = ledger.Range("A" & ledgerRow & ":Y" & ledgerRow).Formula2
    If col = 15 Then
        If count > 0 Then
            portion = Number(raw.Cells(rawRow, 23).Value2) + desired - oldTotal
        Else
            portion = desired
        End If
        If portion < 0 Then Err.Raise vbObjectError + 134, , "선택 거래의 금액보다 감소액이 큽니다. 다른 거래를 선택하거나 RAW를 수정해 주세요."
        changed = True
        If isNew Then
            raw.Cells(rawRow, 1).Value2 = CDbl(newDate)
            raw.Cells(rawRow, 3).Value2 = CStr(newAsset)
            raw.Cells(rawRow, 4).Value2 = "고정비"
            raw.Cells(rawRow, 6).Value2 = CStr(newTitle)
            raw.Cells(rawRow, 7).Value2 = portion
            raw.Cells(rawRow, 8).Value2 = "지출"
            raw.Cells(rawRow, 10).Value2 = "고정비"
            raw.Cells(rawRow, 11).Value2 = portion
            raw.Cells(rawRow, 12).Value2 = "본인"
            raw.Cells(rawRow, 43).Value2 = slot
        Else
            fullBurden = Abs(Number(raw.Cells(rawRow, 23).Value2) - oldGross) < 0.0001
            If fullBurden Then
                raw.Cells(rawRow, 7).Value2 = portion
                If Not raw.Cells(rawRow, 11).HasFormula Then raw.Cells(rawRow, 11).Value2 = portion
            Else
                raw.Cells(rawRow, 11).Value2 = portion
            End If
            raw.Cells(rawRow, 10).Value2 = "고정비"
            raw.Cells(rawRow, 43).Value2 = slot
        End If
        raw.Calculate
        SyncCash raw, rawRow, ledger, ledgerRow, ledgerOld, oldGross
        ValidateCash ledger, ledgerRow
    Else
        changed = True
        If IsEmpty(typed) Or Len(CStr(typed)) = 0 Then
            raw.Cells(rawRow, 9).ClearContents
        Else
            raw.Cells(rawRow, 9).NumberFormat = "@"
            raw.Cells(rawRow, 9).Value2 = CStr(typed)
        End If
    End If
    Application.Calculate
    If col = 15 Then
        If Abs(Number(sh.Cells(slot, 15).Value2) - desired) > 0.0001 Then _
            Err.Raise vbObjectError + 135, , "연결 후 합계가 입력액과 다릅니다. 변경을 취소했습니다."
    End If
    SnapshotOne raw, rawRow
CleanUp:
    If savedState Then
        Application.Calculation = savedCalc
        Application.EnableEvents = savedEvents
    End If
    busy = False
    Exit Sub
Failed:
    failed = Err.Description
    On Error Resume Next
    If changed Then
        raw.Range("A" & rawRow & ":AQ" & rawRow).Formula2 = oldRaw
        If ledgerRow > 0 Then ledger.Range("A" & ledgerRow & ":Y" & ledgerRow).Formula2 = ledgerOld
        Application.Calculate
    End If
    MsgBox failed, vbExclamation, "고정비 수정"
    Resume CleanUp
End Sub

Private Sub RestoreCell(ByVal sh As Worksheet, ByVal slot As Long, ByVal column As Long)
    Dim r As Long, formula As String
    r = FIRST_META + (Val(sh.Name) - 1) * 23 + slot - 4
    formula = CStr(ThisWorkbook.Worksheets(META).Cells(r, IIf(column = 15, 88, 89)).Value2)
    If Left$(formula, 1) <> "=" Then Err.Raise vbObjectError + 140, , "고정비 연결 수식이 없습니다."
    sh.Cells(slot, column).Formula2 = Replace(Replace(formula, "_xlfn.", ""), "_xlws.", "")
End Sub

Private Function FixedRows(ByVal raw As Worksheet, ByVal m As Long, ByVal slot As Long) As Collection
    Dim rows As New Collection, r As Long, stamp As Variant
    For r = 2 To RAW_LAST
        stamp = raw.Cells(r, 1).Value2
        If IsNumeric(stamp) And Not IsEmpty(stamp) Then
            If CDbl(stamp) > 0 Then
                If Year(CDate(stamp)) = 2026 And Month(CDate(stamp)) = m And _
                   CStr(raw.Cells(r, 20).Value2) = "고정비" And _
                   CStr(raw.Cells(r, 21).Value2) = "본인" And _
                   Number(raw.Cells(r, 17).Value2) = 1 And Number(raw.Cells(r, 27).Value2) = slot Then
                    rows.Add CLng(raw.Cells(r, 42).Value2)
                End If
            End If
        End If
    Next r
    Set FixedRows = rows
End Function

Private Function ChooseID(ByVal raw As Worksheet, ByVal ids As Collection, ByVal heading As String) As Long
    Dim i As Long, r As Long, answer As Variant, prompt As String
    prompt = heading & vbCrLf
    For i = 1 To ids.Count
        r = RowForID(raw, CLng(ids(i)))
        If i <= 12 Then prompt = prompt & CStr(ids(i)) & " | " & Format$(raw.Cells(r, 1).Value, "m/d") & _
                    " | " & Left$(CStr(raw.Cells(r, 6).Value2), 24) & " | " & _
                    Format$(Number(raw.Cells(r, 23).Value2), "#,##0") & "원" & vbCrLf
    Next i
    answer = InputBox(prompt, "수정할 거래 선택")
    If Len(CStr(answer)) = 0 Then Exit Function
    If Not IsNumeric(answer) Then Err.Raise vbObjectError + 141, , "거래ID는 숫자로 입력해 주세요."
    For i = 1 To ids.Count
        If CDbl(ids(i)) = CDbl(answer) Then ChooseID = CLng(answer): Exit Function
    Next i
    Err.Raise vbObjectError + 141, , "표시된 거래ID 중에서 선택해 주세요."
End Function

Private Function RowForID(ByVal raw As Worksheet, ByVal ident As Long) As Long
    Dim found As Variant
    found = Application.Match(ident, raw.Range("AP2:AP4001"), 0)
    If Not IsError(found) Then RowForID = CLng(found) + 1
End Function

Private Function EmptyRawRow(ByVal raw As Worksheet) As Long
    Dim r As Long
    For r = 2 To RAW_LAST
        If IsEmpty(raw.Cells(r, 1).Value2) And Application.CountA(raw.Range("C" & r & ":I" & r)) = 0 Then
            EmptyRawRow = r: Exit Function
        End If
    Next r
End Function

Private Function LastField(ByVal raw As Worksheet, ByVal slot As Long, ByVal column As Long) As String
    Dim r As Long, latest As Double
    For r = 2 To RAW_LAST
        If Number(raw.Cells(r, 27).Value2) = slot And Number(raw.Cells(r, 15).Value2) > latest Then
            latest = Number(raw.Cells(r, 15).Value2)
            LastField = CStr(raw.Cells(r, column).Value2)
        End If
    Next r
End Function

Private Function LinkedLedgerRow(ByVal ledger As Worksheet, ByVal ident As Long) As Long
    Dim r As Long
    For r = 64 To 500
        If Number(ledger.Cells(r, 23).Value2) = ident Or Number(ledger.Cells(r, 24).Value2) = ident Then
            If LinkedLedgerRow > 0 Then Err.Raise vbObjectError + 142, , "현금 원장이 여러 건 연결돼 있습니다. RAW와 원장을 확인해 주세요."
            LinkedLedgerRow = r
        End If
    Next r
End Function

Private Sub SyncCash(ByVal raw As Worksheet, ByVal r As Long, ByVal ledger As Worksheet, _
                     ByRef ledgerRow As Long, ByRef old As Variant, ByVal oldGross As Double)
    Dim asset As String, amount As Double, i As Long, stamp As Date, last As Long
    asset = CStr(raw.Cells(r, 3).Value2): amount = Number(raw.Cells(r, 7).Value2)
    stamp = CDate(raw.Cells(r, 1).Value2)
    If ledgerRow > 0 Then
        If Number(ledger.Cells(ledgerRow, 24).Value2) <> 0 Or _
           Abs(Number(ledger.Cells(ledgerRow, 5).Value2) - oldGross) > 0.0001 Then _
            Err.Raise vbObjectError + 143, , "분할·합산된 현금 원장은 RAW에서 확인해 주세요."
        If amount = 0 Then
            ledger.Range("A" & ledgerRow & ":E" & ledgerRow).ClearContents
        Else
            If IsEmpty(ledger.Cells(ledgerRow, 1).Value2) Then
                ledger.Cells(ledgerRow, 1).Value2 = CDbl(stamp)
                ledger.Cells(ledgerRow, 2).Value2 = "출금"
                ledger.Cells(ledgerRow, 3).Value2 = "생활비"
                ledger.Cells(ledgerRow, 4).Value2 = "지출"
                If Left$(CStr(raw.Cells(r, 8).Value2), 2) = "이체" Then
                    ledger.Cells(ledgerRow, 2).Value2 = "계좌이체"
                    ledger.Cells(ledgerRow, 4).Value2 = "외부"
                End If
            End If
            ledger.Cells(ledgerRow, 5).Value2 = amount
        End If
        Exit Sub
    End If
    If asset <> "미래에셋_WRAP형CMA" And Not (asset = "네이버미래에셋" And stamp >= DateSerial(2026, 11, 12)) Then Exit Sub
    If amount <= 0 Then Exit Sub
    For i = 64 To 500
        If Not IsEmpty(ledger.Cells(i, 1).Value2) Or Not IsEmpty(ledger.Cells(i, 23).Value2) Then last = i
    Next i
    ledgerRow = last + 1
    If ledgerRow > 500 Then ledgerRow = 0: Err.Raise vbObjectError + 144, , "현금 원장에 빈 입력행이 없습니다."
    old = ledger.Range("A" & ledgerRow & ":Y" & ledgerRow).Formula2
    ledger.Cells(ledgerRow, 1).Value2 = CDbl(stamp)
    ledger.Cells(ledgerRow, 2).Value2 = "출금"
    ledger.Cells(ledgerRow, 3).Value2 = "생활비"
    ledger.Cells(ledgerRow, 4).Value2 = "지출"
    ledger.Cells(ledgerRow, 5).Value2 = amount
    ledger.Cells(ledgerRow, 6).Value2 = CStr(raw.Cells(r, 6).Value2)
    ledger.Cells(ledgerRow, 10).Value2 = asset
    ledger.Cells(ledgerRow, 11).Value2 = "외부"
    ledger.Cells(ledgerRow, 12).Value2 = "WRAP입출금"
    ledger.Cells(ledgerRow, 23).Value2 = CLng(raw.Cells(r, 42).Value2)
End Sub

Private Function Number(ByVal value As Variant) As Double
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then Exit Function
    If IsNumeric(value) Then Number = CDbl(value)
End Function

Private Sub ValidateCash(ByVal ledger As Worksheet, ByVal r As Long)
    If r = 0 Then Exit Sub
    ledger.Cells(r, 9).Calculate
    If Not IsEmpty(ledger.Cells(r, 1).Value2) Then
        If CStr(ledger.Cells(r, 9).Value2) <> "OK" Then _
            Err.Raise vbObjectError + 145, , "현금 원장의 금액·연결을 확인해 주세요. 변경을 취소했습니다."
    End If
End Sub

Private Function IsCancelled(ByVal value As Variant) As Boolean
    If VarType(value) = vbBoolean Then IsCancelled = (value = False)
End Function
