Attribute VB_Name = "BBSearch"
Option Explicit

' Read RAW only; write search result formulas only. Never post transactions.
Public Sub BB_Search()
    Dim book As Workbook, view As Worksheet, raw As Worksheet
    Dim data As Variant, query As String, monthText As String, needle As String
    Dim row As Long, hit As Long, wantedMonth As Long, lastRow As Long, oldCount As Long
    Dim stamp As Variant, ident As Variant, position As String, sourceRow As String
    Dim savedScreen As Boolean, savedEvents As Boolean, savedCalc As XlCalculation
    Dim oldStatus As Variant, wasSaved As Boolean
    On Error GoTo Failed
    Set book = ThisWorkbook
    Set view = book.Names("BB_Query").RefersToRange.Worksheet
    Set raw = book.Worksheets("RAW")
    query = Trim$(CStr(book.Names("BB_Query").RefersToRange.Value2))
    monthText = CStr(book.Names("BB_Month").RefersToRange.Value2)
    wantedMonth = Val(monthText)
    If wantedMonth < 0 Or wantedMonth > 12 Then wantedMonth = 0
    oldCount = Val(CStr(view.Range("J1").Value2))
    If oldCount < 1 Or oldCount > 4000 Then oldCount = 1
    savedScreen = Application.ScreenUpdating
    savedEvents = Application.EnableEvents
    savedCalc = Application.Calculation
    oldStatus = Application.StatusBar
    wasSaved = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Searching RAW..."
    raw.Calculate
    data = raw.Range("A2:AQ4001").Value2
    view.Range("A8:H" & CStr(oldCount + 7)).ClearContents
    lastRow = UBound(data, 1)
    For row = 1 To lastRow
        stamp = data(row, 1)
        ident = data(row, 42)
        If IsNumeric(stamp) And Not IsEmpty(stamp) And IsNumeric(ident) Then
            If CDbl(stamp) > 0 Then
                If wantedMonth = 0 Or Month(CDate(stamp)) = wantedMonth Then
                    needle = TextValue(data(row, 3)) & " " & TextValue(data(row, 6)) _
                             & " " & TextValue(data(row, 9)) & " " & TextValue(data(row, 7))
                    If Len(query) = 0 Or InStr(1, needle, query, vbTextCompare) > 0 Then
                        hit = hit + 1
                        view.Cells(hit + 7, 8).Value2 = CLng(ident)
                        position = "IFERROR(MATCH($H" & CStr(hit + 7) & ",RAW!$AP$2:$AP$4001,0),0)"
                        sourceRow = "(" & position & "+1)"
                        view.Cells(hit + 7, 1).Formula = "=IFERROR(INDEX(RAW!$A$2:$A$4001," & position & "),"""")"
                        view.Cells(hit + 7, 2).Formula = "=IFERROR(INDEX(RAW!$C$2:$C$4001," & position & "),"""")"
                        view.Cells(hit + 7, 3).Formula = "=IFERROR(INDEX(RAW!$F$2:$F$4001," & position & "),"""")"
                        view.Cells(hit + 7, 4).Formula = "=IFERROR(INDEX(RAW!$G$2:$G$4001," & position & "),"""")"
                        view.Cells(hit + 7, 5).Formula = "=IFERROR(IF(INDEX(RAW!$I$2:$I$4001," & position & ")="""","""",INDEX(RAW!$I$2:$I$4001," & position & ")),"""")"
                        view.Cells(hit + 7, 6).Formula = "=IF(" & position & "=0,"""",HYPERLINK(""#'RAW'!A""&" & sourceRow & ",""" & view.Range("J3").Value2 & """))"
                        view.Cells(hit + 7, 7).Formula = "=IF(" & position & "=0,"""",HYPERLINK(""#'RAW'!I""&" & sourceRow & ",""" & view.Range("J4").Value2 & """))"
                    End If
                End If
            End If
        End If
    Next row
    view.Range("J1").Value2 = hit
    If hit = 0 Then
        view.Range("A8").Value2 = view.Range("J5").Value2
    End If
    view.Range("A4").Value2 = Replace(CStr(view.Range("J2").Value2), "{n}", CStr(hit))
    view.Range("A4").Value2 = view.Range("A4").Value2 & "  " & Format$(Now, "mm/dd hh:nn")
    ' Clearing display filters lets result links show the target RAW row.
    If raw.FilterMode Then raw.ShowAllData
    If hit > 0 Then view.Range("A8:G" & CStr(hit + 7)).Calculate
    Application.Goto view.Range("A7"), True
CleanUp:
    If wasSaved Then
        Application.Calculation = savedCalc
        Application.EnableEvents = savedEvents
        Application.ScreenUpdating = savedScreen
        Application.StatusBar = oldStatus
    End If
    Exit Sub
Failed:
    On Error Resume Next
    If Not view Is Nothing Then
        view.Range("A4").Value2 = "Search failed: " & Err.Description
    End If
    Resume CleanUp
End Sub

Private Function TextValue(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        TextValue = ""
    Else
        TextValue = CStr(value)
    End If
End Function
