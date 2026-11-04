Option Explicit
Option Private Module

Public Cases As Object
Public fields As Object
Public years As Object
Public yearErrors As Object
Public currentCase As CaseRecord
Public isBusy As Boolean
Public lastMessage As String

Public Function NewDictionary() As Object
    Set NewDictionary = CreateObject("Scripting.Dictionary")
End Function

Public Function Sheet(ByVal sheetName As String) As Worksheet
    Set Sheet = ThisWorkbook.Worksheets(sheetName)
End Function

Public Function LastRow(ByVal ws As Worksheet) As Long
    Dim found As Range
    Set found = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    LastRow = 1
    If Not found Is Nothing Then LastRow = found.Row
End Function

Public Function LastColumn(ByVal ws As Worksheet) As Long
    Dim found As Range
    Set found = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    LastColumn = 1
    If Not found Is Nothing Then LastColumn = found.Column
End Function

Public Function SameValue(ByVal first As Variant, ByVal second As Variant) As Boolean
    If IsError(first) Or IsError(second) Then Exit Function
    If IsNull(first) Or IsNull(second) Then Exit Function
    If Len(CStr(first)) = 0 And Len(CStr(second)) = 0 Then
        SameValue = True
    ElseIf VarType(first) = vbString Or VarType(second) = vbString Then
        SameValue = (VarType(first) = VarType(second) And CStr(first) = CStr(second))
    ElseIf IsNumeric(first) And IsNumeric(second) Then
        SameValue = (CDbl(first) = CDbl(second))
    Else
        SameValue = (CStr(first) = CStr(second))
    End If
End Function

Public Sub PutValue(ByVal target As Range, ByVal value As Variant, Optional ByVal isDate As Boolean = False)
    If isDate Then
        target.NumberFormat = "yyyy/mm/dd"
    Else
        target.NumberFormat = "@"
    End If
    target.Value2 = value
End Sub

Public Sub Notify(ByVal message As String)
    lastMessage = message
    MsgBox message, vbInformation, "CONTROLLER"
End Sub

Public Function BeginAction() As Boolean
    If isBusy Then Exit Function
    isBusy = True
    BeginAction = True
End Function

Public Function PendingSave() As Boolean
    PendingSave = (Len(CStr(Sheet("Recovery").Range("B1").Value2)) > 0)
End Function

Public Function SafeFailure(ByVal actionName As String, ByVal errorNumber As Long) As String
    'Excel/COMの生の例外本文には値やURLが含まれ得るため表示しない。
    SafeFailure = actionName & "を中止しました。エラー番号 " & CStr(errorNumber) & _
        "。設定と対象管理簿を確認してください。画面の入力は保持しています。"
End Function
