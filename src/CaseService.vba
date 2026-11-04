Option Explicit
Option Private Module

Public Sub ReloadCases(ByRef message As String)
    Dim year As Variant
    Dim loaded As Object
    Dim key As Variant
    Dim detail As String
    Dim invalidCount As Long
    Set Cases = NewDictionary()
    Set yearErrors = NewDictionary()
    For Each year In years.Keys
        detail = ""
        Set loaded = ReadYear(CStr(year), detail)
        If loaded Is Nothing Then
            yearErrors.Add year, detail
            message = message & CStr(year) & "年度：" & detail & vbCrLf
        Else
            For Each key In loaded.Keys
                Cases.Add key, loaded(key)
                If Len(loaded(key).invalidReason) > 0 Then invalidCount = invalidCount + 1
            Next
        End If
    Next
    message = "読込件数 " & Cases.Count & "。正常年度の検索を利用できます。" & vbCrLf & message
    If invalidCount > 0 Then message = message & vbCrLf & "番号・型・重複などの要確認事案 " & invalidCount & "件は更新できません。詳細画面で理由を確認してください。"
End Sub

Public Function SearchPeople(ByVal query As String, ByVal year As String, ByVal fieldId As String, _
    ByVal fieldQuery As String, ByRef message As String) As Object
    Dim groups As Object
    Dim matched As Object
    Dim result As Object
    Dim key As Variant
    Dim groupKey As String
    Dim item As Variant
    Dim record As CaseRecord
    Dim matches As Boolean
    Set groups = NewDictionary()
    Set matched = NewDictionary()
    Set result = NewDictionary()
    If Len(fieldId) > 0 Then
        If Not fields.Exists(fieldId) Then
            message = "追加検索条件のFieldIdがFieldsにありません。"
            Set SearchPeople = result
            Exit Function
        End If
        If fieldId = "MaskedNumber" Then
            message = "MaskedNumberは検索条件に使用しません。"
            Set SearchPeople = result
            Exit Function
        End If
    End If
    For Each key In Cases.Keys
        Set record = Cases(key)
        groupKey = record.PersonKey
        If Len(groupKey) = 0 Then groupKey = "Invalid|" & CStr(key)
        If Not groups.Exists(groupKey) Then groups.Add groupKey, Array(CStr(record.values("PersonName")), CStr(record.values("Branch")), 0, "")
        item = groups(groupKey)
        item(2) = CLng(item(2)) + 1
        If Len(record.invalidReason) > 0 Then item(3) = "要確認事案あり"
        groups(groupKey) = item
        matches = (InStr(1, CStr(record.values("PersonName")), query, vbTextCompare) > 0)
        If Len(year) > 0 Then
            If record.sourceYear <> year Then matches = False
        End If
        If Len(fieldId) > 0 Then
            If InStr(1, CStr(record.values(fieldId)), fieldQuery, vbTextCompare) = 0 Then matches = False
        End If
        If matches Then matched(groupKey) = True
    Next
    For Each key In matched.Keys
        result.Add key, groups(key)
    Next
    Set SearchPeople = result
End Function

Public Function ScreenValues(ByVal masked As String, ByRef message As String) As Object
    Dim result As Object
    Dim id As Variant
    Dim value As Variant
    Dim normalized As Variant
    Set result = NewDictionary()
    For Each id In fields.Keys
        If id = "MaskedNumber" Then
            result.Add id, masked
        Else
            value = FieldCell(CStr(id)).Value2
            If Not NormalizeValue(CStr(id), value, normalized) Then
                message = "入力項目 " & CStr(id) & "の型または形式が不正です。FieldsのTypeと入力値を確認してください。"
                Exit Function
            End If
            result.Add id, normalized
        End If
    Next
    Set ScreenValues = result
End Function

Public Sub DisplayValues(ByVal values As Object)
    Dim id As Variant
    For Each id In fields.Keys
        PutValue FieldCell(CStr(id)), DisplayValue(CStr(id), values(id)), (fields(id)(1) = "Date")
    Next
End Sub

Public Sub RefreshOneYear(ByVal year As String, ByRef message As String)
    Dim loaded As Object
    Dim key As Variant
    Dim detail As String
    Set loaded = ReadYear(year, detail)
    For Each key In Cases.Keys
        If Cases(key).sourceYear = year Then Cases.Remove key
    Next
    If loaded Is Nothing Then
        yearErrors(year) = detail
        message = message & vbCrLf & "保存は照合済みですが、年度の再読込に失敗しました。" & detail
    Else
        If yearErrors.Exists(year) Then yearErrors.Remove year
        For Each key In loaded.Keys
            Cases.Add key, loaded(key)
        Next
    End If
End Sub
