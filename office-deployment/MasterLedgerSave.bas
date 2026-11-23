Attribute VB_Name = "MasterLedgerSave"
Option Explicit

' 通常更新は進捗整理簿【全体】だけに保存する。
' 5署への反映はTransferToFiveOfficesで行う。
Public Sub SaveCurrentCase()
    Dim record As CaseRecord
    Dim definitions As Worksheet
    Dim detailSheet As Worksheet
    Dim targetBook As Workbook
    Dim targetSheet As Worksheet
    Dim changes As Object
    Dim fieldColumns As Object
    Dim fieldName As String
    Dim targetAddress As String
    Dim originalValue As Variant
    Dim inputValue As Variant
    Dim latestValue As Variant
    Dim item As Variant
    Dim settingRow As Long
    Dim lastSettingRow As Long
    Dim recordRow As Long
    Dim wasSaved As Boolean
    Dim previousEvents As Boolean
    Dim eventsChanged As Boolean
    Dim errorText As String

    On Error GoTo SaveFailed
    Set record = GetDisplayedCaseRecord()
    If record Is Nothing Then Err.Raise vbObjectError + 911, , "先に検索結果から詳細画面を表示してください。"
    Set definitions = ThisWorkbook.Worksheets("Sheet2")
    Set detailSheet = ThisWorkbook.Worksheets("詳細")
    Set changes = CreateObject("Scripting.Dictionary")
    Set fieldColumns = CreateObject("Scripting.Dictionary")
    lastSettingRow = definitions.Cells(definitions.Rows.Count, "A").End(xlUp).Row

    ' Sheet2のC列に指定されたセルだけを編集対象にする。
    For settingRow = 2 To lastSettingRow
        fieldName = Trim$(CStr(definitions.Cells(settingRow, "A").Value2))
        targetAddress = Trim$(CStr(definitions.Cells(settingRow, "C").Value2))
        If Len(targetAddress) > 0 Then
            If Len(fieldName) = 0 Then Err.Raise vbObjectError + 912, , "Sheet2に項目名がありません: " & settingRow
            If fieldColumns.Exists(fieldName) Then Err.Raise vbObjectError + 913, , "編集項目が重複しています: " & fieldName
            fieldColumns.Add fieldName, ColumnFromSettings(definitions.Cells(settingRow, "B").Value2)
            inputValue = detailSheet.Range(targetAddress).MergeArea.Cells(1, 1).Value2
            originalValue = record.Value(fieldName)
            If Not AreMasterValuesEqual(fieldName, originalValue, inputValue, record) Then
                If fieldName = "共通番号" Or fieldName = "枝番号" Or fieldName = "番号" Or fieldName = "署名" Then
                    Err.Raise vbObjectError + 914, , "識別項目・署名の変更は通常更新できません: " & fieldName
                End If
                changes.Add fieldName, inputValue
            End If
        End If
    Next settingRow

    If changes.Count = 0 Then
        Debug.Print "MASTER_UPDATE=NO_CHANGE"
        MsgBox "変更された項目はありません。", vbInformation, "更新"
        Exit Sub
    End If

    previousEvents = Application.EnableEvents
    Application.EnableEvents = False
    eventsChanged = True
    Set targetBook = Workbooks.Open(Filename:=GetMasterLedgerPath(), ReadOnly:=False, UpdateLinks:=0)
    If targetBook.ReadOnly Then Err.Raise vbObjectError + 915, , "全体ファイルが読み取り専用です。"
    Set targetSheet = targetBook.Worksheets("管理簿")
    ValidateMasterHeaders targetSheet
    recordRow = FindMasterRecordRow(targetSheet, record)
    If recordRow = 0 Then Err.Raise vbObjectError + 916, , "対象事案が見つからないか、識別番号が重複しています。"

    ' 保存直前の値を照合し、他の編集との競合を防ぐ。
    For Each item In changes.Keys
        fieldName = CStr(item)
        latestValue = targetSheet.Cells(recordRow, CLng(fieldColumns(fieldName))).Value2
        originalValue = record.Value(fieldName)
        If Not AreMasterValuesEqual(fieldName, latestValue, originalValue, record) Then
            Err.Raise vbObjectError + 917, , "ほかの職員による変更を検出しました: " & fieldName & vbCrLf & "再検索してください。"
        End If
    Next item

    ' 変更項目だけを保存。ほかのセルは書き換えない。
    For Each item In changes.Keys
        fieldName = CStr(item)
        targetSheet.Cells(recordRow, CLng(fieldColumns(fieldName))).Value2 = changes(item)
    Next item
    targetBook.Save
    wasSaved = True

    ' 保存済みの値をCaseRecordへ再読込し、次回更新時の比較元にする。
    Dim rowData As Variant
    Dim fieldNames As Variant
    rowData = targetSheet.Range(targetSheet.Cells(recordRow, 1), targetSheet.Cells(recordRow, 48)).Value2
    fieldNames = GetMasterFieldNames()
    record.LoadData rowData, 1, fieldNames
    targetBook.Close SaveChanges:=False
    Set targetBook = Nothing
    Application.EnableEvents = previousEvents
    eventsChanged = False
    Debug.Print "MASTER_UPDATE=PASS : " & changes.Count & "項目"
    MsgBox changes.Count & "項目を全体ファイルへ保存しました。", vbInformation, "更新完了"
    Exit Sub
SaveFailed:
    errorText = Err.Description
    On Error Resume Next
    If Not targetBook Is Nothing Then targetBook.Close SaveChanges:=False
    If eventsChanged Then Application.EnableEvents = previousEvents
    On Error GoTo 0
    If wasSaved Then errorText = errorText & vbCrLf & "全体ファイルの保存は完了しています。内容を確認してください。"
    Debug.Print "MASTER_UPDATE=ERROR : " & errorText
    MsgBox "更新できませんでした。" & vbCrLf & errorText, vbExclamation, "更新"
End Sub

' 共通番号＋枝番号で全体ファイル内の事案を一意に特定する。
Private Function FindMasterRecordRow(ByVal targetSheet As Worksheet, ByVal record As CaseRecord) As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim matchCount As Long
    Dim commonNumber As String
    Dim branchNumber As String
    commonNumber = CStr(record.Value("共通番号"))
    branchNumber = CStr(record.Value("枝番号"))
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, ColumnFromSettings(ThisWorkbook.Worksheets("Sheet2").Cells(2, "B").Value2)).End(xlUp).Row
    For rowIndex = 2 To lastRow
        If CStr(targetSheet.Cells(rowIndex, GetFieldColumn("共通番号")).Value2) = commonNumber And _
           CStr(targetSheet.Cells(rowIndex, GetFieldColumn("枝番号")).Value2) = branchNumber Then
            matchCount = matchCount + 1
            FindMasterRecordRow = rowIndex
        End If
    Next rowIndex
    If matchCount <> 1 Then FindMasterRecordRow = 0
End Function

' 日付項目はExcelの日付シリアル値と画面の日付表記を同一基準で比較する。
Private Function AreMasterValuesEqual(ByVal fieldName As String, ByVal firstValue As Variant, _
                                      ByVal secondValue As Variant, ByVal record As CaseRecord) As Boolean
    Dim firstDate As Long
    Dim secondDate As Long

    If record.SameValue(firstValue, secondValue) Then
        AreMasterValuesEqual = True
        Exit Function
    End If
    If IsError(firstValue) Or IsError(secondValue) Then Exit Function
    If IsNull(firstValue) Or IsNull(secondValue) Then Exit Function

    Select Case fieldName
        Case "申請日", "処理期限", "署引継日", "取下書収受日", _
             "処理年月日", "施行年月日", "仮装経理に係る最終処理年月日"
            If TryGetMasterDate(firstValue, firstDate) And _
               TryGetMasterDate(secondValue, secondDate) Then
                AreMasterValuesEqual = (firstDate = secondDate)
            End If
    End Select
End Function

' Excel内部の数値・日付型・画面上の日付文字列を、日単位の数値へ揃える。
Private Function TryGetMasterDate(ByVal sourceValue As Variant, ByRef dateNumber As Long) As Boolean
    Dim numericValue As Double
    Dim textValue As String

    If IsError(sourceValue) Or IsNull(sourceValue) Or IsEmpty(sourceValue) Then Exit Function

    If VarType(sourceValue) = vbDate Then
        dateNumber = CLng(Fix(CDbl(sourceValue)))
        TryGetMasterDate = True
        Exit Function
    End If

    textValue = Trim$(CStr(sourceValue))
    If Len(textValue) = 0 Then Exit Function

    ' Value2で取得したExcel日付は数値として渡される。
    If IsNumeric(sourceValue) Then
        numericValue = CDbl(sourceValue)
        If numericValue < 1 Or numericValue > 2958465 Then Exit Function
        dateNumber = CLng(Fix(numericValue))
        TryGetMasterDate = True
        Exit Function
    End If

    If IsDate(textValue) Then
        dateNumber = CLng(Fix(CDbl(CDate(textValue))))
        TryGetMasterDate = True
    End If
End Function

' Sheet2のB列に指定された管理簿の列を取得する。
Private Function GetFieldColumn(ByVal fieldName As String) As Long
    Dim definitions As Worksheet
    Dim settingRow As Long
    Set definitions = ThisWorkbook.Worksheets("Sheet2")
    For settingRow = 2 To definitions.Cells(definitions.Rows.Count, "A").End(xlUp).Row
        If CStr(definitions.Cells(settingRow, "A").Value2) = fieldName Then
            GetFieldColumn = ColumnFromSettings(definitions.Cells(settingRow, "B").Value2)
            Exit Function
        End If
    Next settingRow
    Err.Raise vbObjectError + 918, "GetFieldColumn", "列設定が見つかりません: " & fieldName
End Function

Private Function ColumnFromSettings(ByVal columnText As Variant) As Long
    Dim columnName As String
    columnName = UCase$(Trim$(CStr(columnText)))
    If Len(columnName) = 0 Then Err.Raise vbObjectError + 919, "ColumnFromSettings", "Sheet2のB列が空欄です。"
    ColumnFromSettings = ThisWorkbook.Worksheets("Sheet2").Range(columnName & "1").Column
End Function