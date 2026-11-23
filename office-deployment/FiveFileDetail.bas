Attribute VB_Name = "FiveFileDetail"
Option Explicit

' 現在の詳細画面に表示している事案を保持する。
Private displayedRecord As CaseRecord

' 検索結果の選択事案を、Sheet2の表示先設定に従って詳細画面へ表示する。
Public Sub ShowSelectedCaseDetail()

    Dim selectedRecord As CaseRecord
    Dim detailSheet As Worksheet
    Dim definitionSheet As Worksheet
    Dim targetCells As Object
    Dim previousScreenUpdating As Boolean
    Dim fieldName As String
    Dim targetAddress As String
    Dim targetCell As Range
    Dim fieldValue As Variant
    Dim settingRow As Long
    Dim lastSettingRow As Long

    previousScreenUpdating = Application.ScreenUpdating
    On Error GoTo DisplayFailed

    Set selectedRecord = GetSelectedSearchRecord()
    If selectedRecord Is Nothing Then
        MsgBox "検索結果の事案行を選択してください。", vbInformation, "詳細表示"
        Exit Sub
    End If

    Set definitionSheet = ThisWorkbook.Worksheets("Sheet2")
    Set detailSheet = ThisWorkbook.Worksheets("詳細")
    Set targetCells = CreateObject("Scripting.Dictionary")
    lastSettingRow = definitionSheet.Cells(definitionSheet.Rows.Count, "A").End(xlUp).Row

    ' 書き込む前に、設定セルが正しいかを一通り検証する。
    For settingRow = 2 To lastSettingRow
        fieldName = Trim$(CStr(definitionSheet.Cells(settingRow, "A").Value2))
        targetAddress = Trim$(CStr(definitionSheet.Cells(settingRow, "C").Value2))
        If Len(targetAddress) > 0 Then
            If Len(fieldName) = 0 Then
                Err.Raise vbObjectError + 651, "FiveFileDetail", _
                    "Sheet2のA" & settingRow & "に項目名がありません。"
            End If

            On Error Resume Next
            Set targetCell = detailSheet.Range(targetAddress)
            If Err.Number <> 0 Or targetCell Is Nothing Then
                Err.Clear
                On Error GoTo DisplayFailed
                Err.Raise vbObjectError + 652, "FiveFileDetail", _
                    "Sheet2のC" & settingRow & "のセル番地が不正です: " & targetAddress
            End If
            On Error GoTo DisplayFailed

            If targetCell.Cells.CountLarge <> 1 Then
                Err.Raise vbObjectError + 653, "FiveFileDetail", _
                    "Sheet2のC" & settingRow & "にはセルを1個指定してください。"
            End If
            targetAddress = targetCell.MergeArea.Cells(1, 1).Address(False, False)
            If targetCells.Exists(targetAddress) Then
                Err.Raise vbObjectError + 654, "FiveFileDetail", _
                    "表示先の重複があります: " & targetAddress
            End If
            targetCells.Add targetAddress, fieldName
            Set targetCell = Nothing
        End If
    Next settingRow

    Application.ScreenUpdating = False

    ' 現在の表示先だけを消去する。見出し・結合・罫線は変更しない。
    Dim address As Variant
    For Each address In targetCells.Keys
        detailSheet.Range(CStr(address)).MergeArea.ClearContents
    Next address

    ' 項目名に対応する値を、設定されたセルへ表示する。
    For Each address In targetCells.Keys
        fieldName = CStr(targetCells(address))
        fieldValue = selectedRecord.Value(fieldName)
        If Not IsError(fieldValue) Then
            If Not IsNull(fieldValue) Then
                If Not IsEmpty(fieldValue) Then
                    Set targetCell = detailSheet.Range(CStr(address))
                    If IsDetailDateField(fieldName) And IsNumeric(fieldValue) Then
                        targetCell.Value2 = Format$(CDate(fieldValue), "yyyy/mm/dd")
                    Else
                        targetCell.Value2 = fieldValue
                    End If
                End If
            End If
        End If
    Next address

    ' 事案の案内表示は項目リストとは別の画面固定表示として扱う。
    detailSheet.Range("Q9").Value2 = _
        CStr(selectedRecord.Value("署名")) & " / " & _
        CStr(selectedRecord.Value("氏名法人名"))

    Set displayedRecord = selectedRecord
    detailSheet.Activate
    detailSheet.Range("C12").Select
    Application.ScreenUpdating = previousScreenUpdating
    Debug.Print "DETAIL=PASS : " & targetCells.Count & "項目"
    Exit Sub

DisplayFailed:
    Application.ScreenUpdating = previousScreenUpdating
    MsgBox "詳細表示に失敗しました。" & vbCrLf & Err.Description, vbExclamation, "詳細表示"
    Debug.Print "DETAIL=ERROR : " & Err.Description
End Sub

' 現在表示中の事案を取得する。
Public Function GetDisplayedCaseRecord() As CaseRecord
    Set GetDisplayedCaseRecord = displayedRecord
End Function

' 日付表示が必要な項目を判定する。配置先の変更には影響しない。
Private Function IsDetailDateField(ByVal fieldName As String) As Boolean
    Select Case fieldName
        Case "申請日", "処理期限", "署引継日", "取下書収受日", _
             "処理年月日", "施行年月日", "仮装経理に係る最終処理年月日"
            IsDetailDateField = True
    End Select
End Function