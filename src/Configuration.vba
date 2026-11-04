Option Explicit
Option Private Module

Public Function Setting(ByVal settingName As String) As String
    Dim rowIndex As Long
    For rowIndex = 2 To LastRow(Sheet("Settings"))
        If CStr(Sheet("Settings").Cells(rowIndex, 1).Value2) = settingName Then
            Setting = CStr(Sheet("Settings").Cells(rowIndex, 2).Value2)
            Exit Function
        End If
    Next
End Function

Public Function LoadConfiguration(ByRef message As String) As Boolean
    Dim fieldDefinitions As Object
    Dim yearDefinitions As Object
    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim fieldId As String
    Dim definition As Variant
    Dim target As Range
    Dim usedCells As Object
    Dim year As String
    Dim address As String
    Dim id As Variant
    Dim yearProblem As String
    Dim previous As Variant
    On Error GoTo Failed
    If ThisWorkbook.Date1904 Then
        message = "CONTROLLERが1904年日付システムです。日付のずれを防ぐため処理を中止します。標準の1900年日付システムの別ブックで構築してください。"
        Exit Function
    End If
    Set fieldDefinitions = NewDictionary()
    Set yearDefinitions = NewDictionary()
    Set usedCells = NewDictionary()
    Set ws = Sheet("Fields")
    For rowIndex = 2 To LastRow(ws)
        fieldId = CStr(ws.Cells(rowIndex, 1).Value2)
        If Len(fieldId) > 0 Then
            message = "Fields 行" & rowIndex & "：項目ID重複、型、Header、Sheet、Cellを確認してください。"
            If fieldDefinitions.Exists(fieldId) Then Exit Function
            definition = Array(CStr(ws.Cells(rowIndex, 2).Value2), CStr(ws.Cells(rowIndex, 3).Value2), _
                CStr(ws.Cells(rowIndex, 4).Value2), CStr(ws.Cells(rowIndex, 5).Value2), CStr(ws.Cells(rowIndex, 6).Value2))
            If Len(definition(0)) = 0 Then Exit Function
            If definition(1) <> "Text" And definition(1) <> "Date" Then Exit Function
            If definition(4) <> "Yes" And definition(4) <> "No" Then Exit Function
            Set target = Sheet(definition(2)).Range(definition(3))
            If target.CountLarge <> 1 Then Exit Function
            If target.HasFormula Then Exit Function
            If target.MergeCells Then
                If target.Address <> target.MergeArea.Cells(1, 1).Address Then
                    message = "Fields 行" & rowIndex & "：結合入力欄は左上セルをCellに設定してください。"
                    Exit Function
                End If
            End If
            If definition(2) = "Settings" Or definition(2) = "Fields" Or definition(2) = "Recovery" Or definition(2) = "MoveRecovery" Or definition(2) = "Calendar" Then Exit Function
            If IsReservedUiCell(target) Then
                message = "Fields 行" & rowIndex & "：" & fieldId & "の入力セルが検索・関連事案一覧・操作欄と重複しています。別のセルを設定してください。"
                Exit Function
            End If
            If usedCells.Exists(target.Parent.Name & "!" & target.Address) Then Exit Function
            usedCells.Add target.Parent.Name & "!" & target.Address, True
            fieldDefinitions.Add fieldId, definition
        End If
    Next
    For Each id In Array("MaskedNumber", "Branch", "PersonName")
        message = "Fields：必須定義がありません。" & CStr(id)
        If Not fieldDefinitions.Exists(id) Then Exit Function
    Next
    If fieldDefinitions("MaskedNumber")(1) <> "Text" Or fieldDefinitions("Branch")(1) <> "Text" Then Exit Function
    Set ws = Sheet("Settings")
    For rowIndex = 2 To LastRow(ws)
        year = CStr(ws.Cells(rowIndex, 4).Value2)
        If Len(year) > 0 Then
            message = "Settings 年度表 行" & rowIndex & "：年度重複、Enabled、直接アドレス、シート名を確認してください。"
            If ws.Cells(rowIndex, 7).Value2 <> "No" Then
                yearProblem = ""
                If CStr(ws.Cells(rowIndex, 7).Value2) <> "Yes" Then yearProblem = "EnabledはYesまたはNoにしてください。"
                address = CStr(ws.Cells(rowIndex, 5).Value2)
                If Len(address) = 0 Or Len(CStr(ws.Cells(rowIndex, 6).Value2)) = 0 Or Len(CStr(ws.Cells(rowIndex, 9).Value2)) = 0 Then yearProblem = "常勤用・非常勤用アドレスとシート名を設定してください。"
                If ws.Cells(rowIndex, 8).Value2 <> "Yes" And ws.Cells(rowIndex, 8).Value2 <> "No" Then yearProblem = "AllowWriteはYesまたはNoにしてください。"
                If yearDefinitions.Exists(year) Then
                    previous = yearDefinitions(year)
                    previous(5) = "年度が重複しています。Settingsの年度表を確認してください。"
                    yearDefinitions(year) = previous
                Else
                    yearDefinitions.Add year, Array(address, CStr(ws.Cells(rowIndex, 6).Value2), CStr(ws.Cells(rowIndex, 8).Value2), address, CStr(ws.Cells(rowIndex, 9).Value2), yearProblem)
                End If
            End If
        End If
    Next
    message = "Settings：有効な年度がありません。"
    If yearDefinitions.Count = 0 Then Exit Function
    message = ""
    Set fields = fieldDefinitions
    Set years = yearDefinitions
    LoadConfiguration = True
    Exit Function
Failed:
    message = message & " 定義したセルが存在し、単一・非結合・数式なしであることを確認してください。"
End Function

Public Function FieldCell(ByVal fieldId As String) As Range
    Dim definition As Variant
    definition = fields(fieldId)
    Set FieldCell = Sheet(definition(2)).Range(definition(3))
End Function

Public Function HeaderFor(ByVal fieldId As String, ByVal year As String, ByRef message As String) As String
    Dim rowIndex As Long
    Dim count As Long
    Dim ws As Worksheet
    HeaderFor = fields(fieldId)(0)
    Set ws = Sheet("Fields")
    For rowIndex = 2 To LastRow(ws)
        If CStr(ws.Cells(rowIndex, 8).Value2) = year And CStr(ws.Cells(rowIndex, 9).Value2) = fieldId Then
            count = count + 1
            HeaderFor = CStr(ws.Cells(rowIndex, 10).Value2)
        End If
    Next
    If count > 1 Then message = year & "年度：Fieldsの年度別見出しが重複しています。" & fieldId
End Function

Public Function ColumnMap(ByVal ws As Worksheet, ByVal year As String, ByRef message As String) As Object
    Dim headers As Object
    Dim result As Object
    Dim usedColumns As Object
    Dim data As Variant
    Dim colIndex As Long
    Dim id As Variant
    Dim heading As String
    Set headers = NewDictionary()
    Set result = NewDictionary()
    Set usedColumns = NewDictionary()
    data = ws.Range(ws.Cells(1, 1), ws.Cells(1, LastColumn(ws))).Value2
    For colIndex = 1 To UBound(data, 2)
        If IsError(data(1, colIndex)) Then
            message = year & "年度：見出しにExcelエラーがあります。"
            Exit Function
        End If
        heading = CStr(data(1, colIndex))
        If Len(heading) > 0 Then
            If headers.Exists(heading) Then
                message = year & "年度：見出しが重複しています。列" & colIndex
                Exit Function
            End If
            headers.Add heading, colIndex
        End If
    Next
    For Each id In fields.Keys
        heading = HeaderFor(CStr(id), year, message)
        If Len(message) > 0 Then Exit Function
        If heading <> "[Absent]" Then
            If Not headers.Exists(heading) Then
                message = year & "年度：" & CStr(id) & "の見出しがありません。管理簿1行目とFieldsのHeaderを確認してください。"
                Exit Function
            End If
            If usedColumns.Exists(CStr(headers(heading))) Then
                message = year & "年度：複数項目が同じ列を参照しています。" & CStr(id)
                Exit Function
            End If
            usedColumns.Add CStr(headers(heading)), True
            result.Add id, headers(heading)
        ElseIf id = "MaskedNumber" Or id = "Branch" Or fields(id)(4) = "Yes" Then
            message = year & "年度：識別項目と必須項目は[Absent]にできません。" & CStr(id)
            Exit Function
        End If
    Next
    If result("MaskedNumber") <> 1 Or result("Branch") <> 2 Then
        message = year & "年度：A列MaskedNumberとB列枝番を確認してください。"
        Exit Function
    End If
    Set ColumnMap = result
End Function

Public Function ValidateValues(ByVal values As Object, ByRef message As String) As Boolean
    Dim id As Variant
    Dim value As Variant
    Dim number As String
    Dim normalized As Variant
    For Each id In fields.Keys
        message = "入力項目 " & CStr(id) & "："
        If Not values.Exists(id) Then
            message = message & "入力がありません。"
            Exit Function
        End If
        value = values(id)
        If IsError(value) Or IsNull(value) Then
            message = message & "Excelエラーは保存できません。"
            Exit Function
        End If
        If fields(id)(4) = "Yes" And Len(CStr(value)) = 0 Then
            message = message & "必須項目です。"
            Exit Function
        End If
        If Not NormalizeValue(CStr(id), value, normalized) Then
            message = message & "形式が不正です。日付はExcel日付、郵便番号は7桁、電話番号は数字で入力してください。"
            Exit Function
        End If
        values(id) = normalized
    Next
    message = "Masked Numberまたは枝番の形式が不正です。更新・コピーを中止します。"
    If Not RestoreNumber(values("MaskedNumber"), number) Then Exit Function
    If Not IsDigits(values("Branch"), 5) Then Exit Function
    message = ""
    ValidateValues = True
End Function
