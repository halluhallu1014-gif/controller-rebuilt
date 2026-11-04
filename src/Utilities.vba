Option Explicit

Private calendarTarget As Range
Private calendarMonth As Date

Public Function IsDigits(ByVal value As Variant, ByVal size As Long) As Boolean
    Dim index As Long
    Dim text As String
    If VarType(value) <> vbString Then Exit Function
    text = CStr(value)
    If Len(text) <> size Then Exit Function
    For index = 1 To size
        If Mid$(text, index, 1) < "0" Or Mid$(text, index, 1) > "9" Then Exit Function
    Next
    IsDigits = True
End Function

Public Function MaskNumber(ByVal number As String) As String
    Dim index As Long
    Dim body As String
    If Not IsDigits(number, 13) Then Exit Function
    For index = 13 To 1 Step -1
        body = body & CStr((CLng(Mid$(number, index, 1)) + 7) Mod 10)
    Next
    MaskNumber = "MN1:" & body & ":" & NumberChecksum(number)
End Function

Public Function RestoreNumber(ByVal masked As Variant, ByRef number As String) As Boolean
    Dim text As String
    Dim body As String
    Dim index As Long
    number = ""
    If VarType(masked) <> vbString Then Exit Function
    text = CStr(masked)
    If Len(text) <> 20 Then Exit Function
    If Left$(text, 4) <> "MN1:" Or Mid$(text, 18, 1) <> ":" Then Exit Function
    body = Mid$(text, 5, 13)
    If Not IsDigits(body, 13) Then Exit Function
    For index = 13 To 1 Step -1
        number = number & CStr((CLng(Mid$(body, index, 1)) + 3) Mod 10)
    Next
    If Right$(text, 2) <> NumberChecksum(number) Then
        number = ""
        Exit Function
    End If
    RestoreNumber = True
End Function

Private Function NumberChecksum(ByVal number As String) As String
    Dim index As Long
    Dim total As Long
    For index = 1 To 13
        total = total + index * CLng(Mid$(number, index, 1))
    Next
    NumberChecksum = Format$(total Mod 97, "00")
End Function

Public Function NormalizeValue(ByVal fieldId As String, ByVal value As Variant, ByRef normalized As Variant) As Boolean
    Dim text As String
    normalized = value
    If IsError(value) Or IsNull(value) Then Exit Function
    If Len(CStr(value)) = 0 Then
        normalized = ""
    ElseIf fields(fieldId)(1) = "Date" Then
        If VarType(value) = vbString Or VarType(value) = vbBoolean Then Exit Function
        If Not IsNumeric(value) Then Exit Function
        If CDbl(value) < 1 Or CDbl(value) > 2958465# Then Exit Function
        If CDbl(value) <> Fix(CDbl(value)) Then Exit Function
        normalized = CDbl(value)
    Else
        '文字列項目を数値から推測復元しない。失われた先頭0は補完しない。
        If VarType(value) <> vbString Then Exit Function
        text = CStr(value)
        If fieldId = "PostalCode" Or fieldId = "PhoneNumber" Then
            text = Replace(text, "-", "")
            If fieldId = "PostalCode" Then
                If Not IsDigits(text, 7) Then Exit Function
            Else
                If Len(text) < 9 Or Len(text) > 11 Then Exit Function
                If Not IsDigits(text, Len(text)) Then Exit Function
            End If
        End If
        normalized = text
    End If
    NormalizeValue = True
End Function

Public Function DisplayValue(ByVal fieldId As String, ByVal value As Variant) As Variant
    Dim text As String
    DisplayValue = value
    If fieldId = "MaskedNumber" Then
        DisplayValue = "*************"
    ElseIf fieldId = "PostalCode" Then
        text = CStr(value)
        If IsDigits(text, 7) Then DisplayValue = Left$(text, 3) & "-" & Right$(text, 4)
    ElseIf fieldId = "PhoneNumber" Then
        DisplayValue = FormatPhone(CStr(value))
    End If
End Function

Private Function FormatPhone(ByVal text As String) As String
    Dim prefix As String
    Dim rowIndex As Long
    Dim areaLength As Long
    Dim exchangeLength As Long
    Dim longest As Long
    Dim ws As Worksheet
    FormatPhone = text
    If Len(text) = 11 Then
        Select Case Left$(text, 3)
            Case "070", "080", "090", "050"
                FormatPhone = Left$(text, 3) & "-" & Mid$(text, 4, 4) & "-" & Right$(text, 4)
                Exit Function
        End Select
    End If
    If Len(text) <> 10 Then Exit Function
    Set ws = Sheet("Settings")
    For rowIndex = 2 To LastRow(ws)
        prefix = CStr(ws.Cells(rowIndex, 10).Value2)
        If Len(prefix) > longest And Len(prefix) > 0 Then
            If Left$(text, Len(prefix)) = prefix Then
                If IsNumeric(ws.Cells(rowIndex, 11).Value2) And IsNumeric(ws.Cells(rowIndex, 12).Value2) Then
                    areaLength = CLng(ws.Cells(rowIndex, 11).Value2)
                    exchangeLength = CLng(ws.Cells(rowIndex, 12).Value2)
                    If areaLength >= 2 And areaLength <= 5 And exchangeLength >= 1 Then
                        If areaLength + exchangeLength + 4 = 10 Then
                            FormatPhone = Left$(text, areaLength) & "-" & Mid$(text, areaLength + 1, exchangeLength) & "-" & Right$(text, 4)
                            longest = Len(prefix)
                        End If
                    End If
                End If
            End If
        End If
    Next
End Function

Public Function WriteClipboard(ByVal text As String) As Boolean
    Dim clipboard As Object
    On Error GoTo Failed
    Set clipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    clipboard.SetText text
    clipboard.PutInClipboard
    WriteClipboard = True
Failed:
End Function

Public Function ReadClipboard(ByRef text As String) As Boolean
    Dim clipboard As Object
    On Error GoTo Failed
    Set clipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    clipboard.GetFromClipboard
    text = clipboard.GetText(1)
    ReadClipboard = True
Failed:
End Function

Public Sub OpenCalendar()
    Dim id As Variant
    Dim target As Range
    If isBusy Then Exit Sub
    If fields Is Nothing Then Exit Sub
    If Not ActiveWorkbook Is ThisWorkbook Then Exit Sub
    For Each id In fields.Keys
        If fields(id)(1) = "Date" Then
            Set target = FieldCell(CStr(id))
            If ActiveSheet Is target.Parent Then
                If Not Intersect(ActiveCell, target) Is Nothing Then
                    Set calendarTarget = target
                    calendarMonth = DateSerial(Year(Date), Month(Date), 1)
                    If IsDate(target.Value) Then calendarMonth = DateSerial(Year(target.Value), Month(target.Value), 1)
                    DrawCalendar
                    Exit Sub
                End If
            End If
        End If
    Next
    Notify "FieldsでType=Dateに設定された入力セルを選択してください。"
End Sub

Private Sub DrawCalendar()
    Dim ws As Worksheet
    Dim firstDate As Date
    Dim index As Long
    Dim dayValue As Date
    Set ws = Sheet("Calendar")
    ws.Range("B2").Value = calendarMonth
    ws.Range("B2").NumberFormat = "yyyy年m月"
    firstDate = calendarMonth - Weekday(calendarMonth, vbSunday) + 1
    ws.Range("B5:H10").NumberFormat = "d"
    For index = 0 To 41
        dayValue = firstDate + index
        ws.Cells(5 + index \ 7, 2 + index Mod 7).Value = dayValue
    Next
    ws.Activate
End Sub

Public Sub PreviousMonth()
    If isBusy Then Exit Sub
    If calendarTarget Is Nothing Then Exit Sub
    calendarMonth = DateAdd("m", -1, calendarMonth)
    DrawCalendar
End Sub

Public Sub NextMonth()
    If isBusy Then Exit Sub
    If calendarTarget Is Nothing Then Exit Sub
    calendarMonth = DateAdd("m", 1, calendarMonth)
    DrawCalendar
End Sub

Public Sub PickCalendarDate()
    If isBusy Then Exit Sub
    If calendarTarget Is Nothing Then Exit Sub
    If Not ActiveSheet Is Sheet("Calendar") Then Exit Sub
    If Intersect(ActiveCell, Sheet("Calendar").Range("B5:H10")) Is Nothing Then Exit Sub
    PutValue calendarTarget, CDbl(ActiveCell.Value2), True
    Application.Goto calendarTarget
End Sub

Public Sub ClearCalendarDate()
    If isBusy Then Exit Sub
    If calendarTarget Is Nothing Then Exit Sub
    calendarTarget.MergeArea.ClearContents
    Application.Goto calendarTarget
End Sub

Public Function ParseImport(ByVal text As String, ByVal values As Object, ByRef message As String) As Boolean
    '実サンプル未確認。明示的な「ラベル:値」または「ラベル<TAB>値」だけを扱う。
    Dim line As Variant
    Dim separator As Long
    Dim label As String
    Dim value As String
    Dim id As Variant
    Dim found As Object
    Set found = NewDictionary()
    text = Replace(text, vbCrLf, vbLf)
    For Each line In Split(text, vbLf)
        separator = InStr(CStr(line), vbTab)
        If separator = 0 Then separator = InStr(CStr(line), "：")
        If separator = 0 Then separator = InStr(CStr(line), ":")
        If separator > 0 Then
            label = Trim$(Left$(CStr(line), separator - 1))
            value = Trim$(Mid$(CStr(line), separator + 1))
            For Each id In fields.Keys
                If label = fields(id)(0) Or label = CStr(id) Then
                    If id <> "MaskedNumber" And id <> "Branch" Then
                        If found.Exists(id) Then
                            message = "ラベルが重複しています。" & CStr(id) & "。貼付内容を確認してください。"
                            Exit Function
                        End If
                        found.Add id, value
                    End If
                End If
            Next
            If label = "識別番号" Then
                If Not IsDigits(value, 18) Then
                    message = "識別番号は半角数字18桁の文字列で貼り付けてください。"
                    Exit Function
                End If
                If found.Exists("MaskedNumber") Then
                    message = "識別番号ラベルが重複しています。"
                    Exit Function
                End If
                found.Add "MaskedNumber", MaskNumber(Left$(value, 13))
                found.Add "Branch", Right$(value, 5)
            End If
        End If
    Next
    If found.Count = 0 Then
        message = "対応するラベルがありません。実サンプルの確認後に解析規則を調整してください。"
        Exit Function
    End If
    For Each id In found.Keys
        values(id) = found(id)
    Next
    message = CStr(found.Count) & "項目を画面へ反映しました。内容を確認してください。管理簿は保存していません。"
    ParseImport = True
End Function
