Attribute VB_Name = "TransferEngine"
Option Explicit

' 週次一般事務整理簿 → 全体 → 5署の転記を担当する独立エンジン。
' Address: A=識別名 B=フォルダパス C=ファイル名 D=シート名
' Sheet2 : A=項目名 B=管理簿列 C=詳細表示セル D=5署へ転記（○/×）
' 進捗整理簿は5行目が見出し、6行目からデータ。転記基準は1行目が見出し。

Private Const ADDRESS_SHEET As String = "Address"
Private Const FIELD_SHEET As String = "Sheet2"
Private Const HISTORY_SHEET As String = "転記基準"
Private Const FIELD_COUNT As Long = 48
Private Const LEDGER_HEADER_ROW As Long = 5
Private Const LEDGER_FIRST_DATA_ROW As Long = LEDGER_HEADER_ROW + 1
Private Const FIRST_OFFICE_ROW As Long = 4
Private Const LAST_OFFICE_ROW As Long = 8

' 毎週の一般事務整理簿を選び、全体へ重複のない事案だけ追加する。
Public Sub ImportWeeklyLedger()
    Dim selectedPath As Variant
    Dim sourcePath As String
    Dim sourceBook As Workbook
    Dim masterBook As Workbook
    Dim sourceSheet As Worksheet
    Dim masterSheet As Worksheet
    Dim fieldNames As Variant
    Dim sourceMap As Object
    Dim masterMap As Object
    Dim existing As Object
    Dim sourceData As Variant
    Dim masterData As Variant
    Dim outputData() As Variant
    Dim newRows As Collection
    Dim values As Variant
    Dim fieldIndex As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim lastRow As Long
    Dim nextRow As Long
    Dim recordKey As String
    Dim fileName As String
    Dim addedCount As Long
    Dim duplicateCount As Long
    Dim sourceColumn As Long
    Dim masterColumn As Long
    Dim saved As Boolean
    Dim errorText As String

    On Error GoTo Failed

    ' ExcelのFileDialogを使用せず、標準のファイル選択画面を開く。
    ' 初期フォルダ指定は行わず、選択画面の起動を妨げないようにする。
    Debug.Print "IMPORT=START : ファイル選択画面を表示します"
    selectedPath = Application.GetOpenFilename( _
        FileFilter:="Excelブック (*.xlsx),*.xlsx", _
        Title:="取り込む一般事務整理簿を選択してください")
    If VarType(selectedPath) = vbBoolean Then
        Debug.Print "IMPORT=CANCEL"
        Exit Sub
    End If
    sourcePath = CStr(selectedPath)
    fileName = Mid$(sourcePath, InStrRev(Replace(sourcePath, "\", "/"), "/") + 1)
    Debug.Print "IMPORT=SELECTED : " & sourcePath

    fieldNames = ReadMasterFields()
    Set sourceBook = Workbooks.Open(Filename:=sourcePath, ReadOnly:=True, UpdateLinks:=0)
    ' D2を省略した一般事務整理簿は先頭シートを使う。
    If Len(Trim$(SafeText(ThisWorkbook.Worksheets(ADDRESS_SHEET).Range("D2").Value2))) > 0 Then
        Set sourceSheet = sourceBook.Worksheets(Trim$(SafeText(ThisWorkbook.Worksheets(ADDRESS_SHEET).Range("D2").Value2)))
    Else
        Set sourceSheet = sourceBook.Worksheets(1)
    End If
    Set sourceMap = ReadHeaderMap(sourceSheet, 1)
    RequireField sourceMap, "共通番号", sourceBook.Name
    RequireField sourceMap, "枝番号", sourceBook.Name
    RequireField sourceMap, "署名", sourceBook.Name
    RequireField sourceMap, "氏名法人名", sourceBook.Name
    sourceData = ReadTable(sourceSheet, CLng(sourceMap("共通番号")), 1)

    Set masterBook = Workbooks.Open(Filename:=GetConfiguredPath(3), ReadOnly:=False, UpdateLinks:=0)
    If masterBook.ReadOnly Then Err.Raise vbObjectError + 801, , "全体ファイルが読み取り専用です。"
    Set masterSheet = masterBook.Worksheets(GetConfiguredSheetName(3))
    Set masterMap = ReadHeaderMap(masterSheet, LEDGER_HEADER_ROW)
    ValidateMasterFields masterMap, fieldNames
    masterData = ReadTable(masterSheet, CLng(masterMap("共通番号")), LEDGER_HEADER_ROW)
    Set existing = BuildKeyIndex(masterData, masterMap, masterBook.Name)
    Set newRows = New Collection

    For rowIndex = 2 To UBound(sourceData, 1)
        If Len(Trim$(SafeText(sourceData(rowIndex, sourceMap("共通番号"))))) > 0 Then
            recordKey = MakeKey(sourceData(rowIndex, sourceMap("共通番号")), _
                                sourceData(rowIndex, sourceMap("枝番号")))
            If existing.Exists(recordKey) Then
                duplicateCount = duplicateCount + 1
            Else
                ' 読み込みファイルにない列は空欄にし、元ファイル名を記録する。
                ReDim values(1 To FIELD_COUNT)
                For fieldIndex = 1 To FIELD_COUNT
                    If sourceMap.Exists(CStr(fieldNames(fieldIndex))) Then
                        sourceColumn = CLng(sourceMap(CStr(fieldNames(fieldIndex))))
                        values(fieldIndex) = sourceData(rowIndex, sourceColumn)
                    Else
                        values(fieldIndex) = Empty
                    End If
                Next fieldIndex
                values(CLng(GetFieldPosition(fieldNames, "元ファイル"))) = fileName
                newRows.Add values
                existing.Add recordKey, True
            End If
        End If
    Next rowIndex

    If newRows.Count > 0 Then
        lastRow = masterSheet.Cells(masterSheet.Rows.Count, CLng(masterMap("共通番号"))).End(xlUp).Row
        nextRow = Application.Max(LEDGER_FIRST_DATA_ROW, lastRow + 1)
        ReDim outputData(1 To newRows.Count, 1 To FIELD_COUNT)
        For rowIndex = 1 To newRows.Count
            values = newRows(rowIndex)
            For fieldIndex = 1 To FIELD_COUNT
                outputData(rowIndex, fieldIndex) = values(fieldIndex)
            Next fieldIndex
        Next rowIndex

        ' 5桁枝番と13桁番号の先頭ゼロを維持する。
        masterSheet.Range(masterSheet.Cells(nextRow, CLng(masterMap("共通番号"))), _
                          masterSheet.Cells(nextRow + newRows.Count - 1, CLng(masterMap("共通番号")))).NumberFormat = "@"
        masterSheet.Range(masterSheet.Cells(nextRow, CLng(masterMap("枝番号"))), _
                          masterSheet.Cells(nextRow + newRows.Count - 1, CLng(masterMap("枝番号")))).NumberFormat = "@"

        ' 列順が異なっても、項目名で対応させて一括転記する。
        For fieldIndex = 1 To FIELD_COUNT
            masterColumn = CLng(masterMap(CStr(fieldNames(fieldIndex))))
            ReDim values(1 To newRows.Count, 1 To 1)
            For rowIndex = 1 To newRows.Count
                values(rowIndex, 1) = outputData(rowIndex, fieldIndex)
            Next rowIndex
            masterSheet.Cells(nextRow, masterColumn).Resize(newRows.Count, 1).Value2 = values
        Next fieldIndex
        masterBook.Save
        saved = True
    End If

    addedCount = newRows.Count
    sourceBook.Close SaveChanges:=False
    Set sourceBook = Nothing
    masterBook.Close SaveChanges:=False
    Set masterBook = Nothing
    Debug.Print "IMPORT=PASS : 追加=" & addedCount & " 重複=" & duplicateCount
    MsgBox "取込完了" & vbCrLf & "新規追加: " & addedCount & "件" & vbCrLf & _
           "重複除外: " & duplicateCount & "件", vbInformation, "一般事務整理簿の取込"
    Exit Sub

Failed:
    errorText = Err.Description
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    If Not masterBook Is Nothing Then masterBook.Close SaveChanges:=False
    On Error GoTo 0
    Debug.Print "IMPORT=ERROR : " & errorText
    If saved Then errorText = errorText & vbCrLf & "全体ファイルへの保存は完了しています。"
    MsgBox "取込を中止しました。" & vbCrLf & errorText, vbExclamation
End Sub

' 全体の変更を、設定された項目だけ5署へ追加・差分転記する。
' 前回転記値は全体ブックの「転記基準」に保存して次回の照合に使う。
Public Sub TransferToFiveOffices()
    Dim masterBook As Workbook
    Dim masterSheet As Worksheet
    Dim historySheet As Worksheet
    Dim officeBook As Workbook
    Dim officeSheet As Worksheet
    Dim fieldNames As Variant
    Dim masterMap As Object
    Dim officeMap As Object
    Dim masterData As Variant
    Dim officeData As Variant
    Dim historyData As Variant
    Dim officeIndexes As Object
    Dim historyIndexes As Object
    Dim enabledFields As Object
    Dim currentRow As Long
    Dim officeRow As Long
    Dim historyRow As Long
    Dim nextOfficeRow As Long
    Dim nextHistoryRow As Long
    Dim settingRow As Long
    Dim fieldIndex As Long
    Dim officeName As String
    Dim officeKey As String
    Dim fullKey As String
    Dim fieldName As String
    Dim sourceValue As Variant
    Dim officeValue As Variant
    Dim previousValue As Variant
    Dim targetData() As Variant
    Dim historyRecord() As Variant
    Dim addedCount As Long
    Dim changedCount As Long
    Dim isInitialHistory As Boolean
    Dim didWrite As Boolean
    Dim didSaveOffice As Boolean
    Dim errorText As String

    On Error GoTo Failed
    fieldNames = ReadMasterFields()
    Set enabledFields = ReadTransferFlags(fieldNames)
    Set masterBook = Workbooks.Open(Filename:=GetConfiguredPath(3), ReadOnly:=False, UpdateLinks:=0)
    If masterBook.ReadOnly Then Err.Raise vbObjectError + 811, , "全体ファイルが読み取り専用です。"
    Set masterSheet = masterBook.Worksheets(GetConfiguredSheetName(3))
    Set masterMap = ReadHeaderMap(masterSheet, LEDGER_HEADER_ROW)
    ValidateMasterFields masterMap, fieldNames
    masterData = ReadTable(masterSheet, CLng(masterMap("共通番号")), LEDGER_HEADER_ROW)
    Set officeIndexes = BuildKeyIndex(masterData, masterMap, masterBook.Name)

    Set historySheet = GetHistorySheet(masterBook, fieldNames)
    historyData = ReadTable(historySheet, 2, 1)
    Set historyIndexes = CreateObject("Scripting.Dictionary")
    For currentRow = 2 To UBound(historyData, 1)
        If Len(SafeText(historyData(currentRow, 1))) > 0 Then
            fullKey = SafeText(historyData(currentRow, 1)) & "|" & _
                      MakeKey(historyData(currentRow, 2), historyData(currentRow, 3))
            If historyIndexes.Exists(fullKey) Then Err.Raise vbObjectError + 812, , "転記基準の識別キーが重複しています: " & fullKey
            historyIndexes.Add fullKey, currentRow
        End If
    Next currentRow

    For settingRow = FIRST_OFFICE_ROW To LAST_OFFICE_ROW
        officeKey = Trim$(SafeText(ThisWorkbook.Worksheets(ADDRESS_SHEET).Cells(settingRow, 1).Value2))
        officeName = GetOfficeName(settingRow)
        Set officeBook = Workbooks.Open(Filename:=GetConfiguredPath(settingRow), ReadOnly:=False, UpdateLinks:=0)
        If officeBook.ReadOnly Then Err.Raise vbObjectError + 813, , "署ファイルが読み取り専用です: " & officeName
        Set officeSheet = officeBook.Worksheets(GetConfiguredSheetName(settingRow))
        Set officeMap = ReadHeaderMap(officeSheet, LEDGER_HEADER_ROW)
        RequireField officeMap, "共通番号", officeName
        RequireField officeMap, "枝番号", officeName

        For fieldIndex = 1 To FIELD_COUNT
            fieldName = CStr(fieldNames(fieldIndex))
            If enabledFields.Exists(fieldName) Then RequireField officeMap, fieldName, officeName
        Next fieldIndex

        officeData = ReadTable(officeSheet, CLng(officeMap("共通番号")), LEDGER_HEADER_ROW)
        Set officeIndexes = BuildKeyIndex(officeData, officeMap, officeName)
        nextOfficeRow = Application.Max(LEDGER_FIRST_DATA_ROW, officeSheet.Cells(officeSheet.Rows.Count, _
                                               CLng(officeMap("共通番号"))).End(xlUp).Row + 1)
        nextHistoryRow = Application.Max(2, historySheet.Cells(historySheet.Rows.Count, 1).End(xlUp).Row + 1)
        didWrite = False

        For currentRow = LEDGER_FIRST_DATA_ROW To UBound(masterData, 1)
            If SafeText(masterData(currentRow, CLng(masterMap("署名")))) = officeName Then
                If Len(Trim$(SafeText(masterData(currentRow, CLng(masterMap("共通番号")))))) > 0 Then
                    fullKey = officeKey & "|" & MakeKey( _
                        masterData(currentRow, CLng(masterMap("共通番号"))), _
                        masterData(currentRow, CLng(masterMap("枝番号"))))
                    officeKey = Trim$(SafeText(ThisWorkbook.Worksheets(ADDRESS_SHEET).Cells(settingRow, 1).Value2))
                    If officeIndexes.Exists(Mid$(fullKey, Len(officeKey) + 2)) Then
                        officeRow = CLng(officeIndexes(Mid$(fullKey, Len(officeKey) + 2)))
                        isInitialHistory = Not historyIndexes.Exists(fullKey)
                        If isInitialHistory Then
                            ' 既存事案は、全体と署が一致している場合だけ基準を初期登録する。
                            For fieldIndex = 1 To FIELD_COUNT
                                fieldName = CStr(fieldNames(fieldIndex))
                                If enabledFields.Exists(fieldName) Then
                                    sourceValue = masterData(currentRow, CLng(masterMap(fieldName)))
                                    officeValue = officeData(officeRow, CLng(officeMap(fieldName)))
                                    If Not ValuesEqual(sourceValue, officeValue) Then
                                        Err.Raise vbObjectError + 814, , "初回転記の照合不一致: " & officeName & " / " & fullKey & " / " & fieldName
                                    End If
                                End If
                            Next fieldIndex
                            historyRow = nextHistoryRow
                            nextHistoryRow = nextHistoryRow + 1
                            historyIndexes.Add fullKey, historyRow
                        Else
                            historyRow = CLng(historyIndexes(fullKey))
                        End If

                        ' 書込前にこの事案の全変更箇所を照合する。
                        For fieldIndex = 1 To FIELD_COUNT
                            fieldName = CStr(fieldNames(fieldIndex))
                            If enabledFields.Exists(fieldName) Then
                                sourceValue = masterData(currentRow, CLng(masterMap(fieldName)))
                                If isInitialHistory Then
                                    previousValue = sourceValue
                                Else
                                    previousValue = historyData(historyRow, fieldIndex + 3)
                                End If
                                officeValue = officeData(officeRow, CLng(officeMap(fieldName)))
                                If Not ValuesEqual(sourceValue, previousValue) Then
                                    If Not ValuesEqual(officeValue, previousValue) And Not ValuesEqual(officeValue, sourceValue) Then
                                        Err.Raise vbObjectError + 815, , "署側の更新と競合: " & officeName & " / " & fullKey & " / " & fieldName
                                    End If
                                End If
                            End If
                        Next fieldIndex

                        For fieldIndex = 1 To FIELD_COUNT
                            fieldName = CStr(fieldNames(fieldIndex))
                            If enabledFields.Exists(fieldName) Then
                                sourceValue = masterData(currentRow, CLng(masterMap(fieldName)))
                                If isInitialHistory Then
                                    previousValue = sourceValue
                                Else
                                    previousValue = historyData(historyRow, fieldIndex + 3)
                                End If
                                If Not ValuesEqual(sourceValue, previousValue) Then
                                    officeValue = officeData(officeRow, CLng(officeMap(fieldName)))
                                    If Not ValuesEqual(officeValue, sourceValue) Then
                                        officeSheet.Cells(officeRow, CLng(officeMap(fieldName))).Value2 = sourceValue
                                        didWrite = True
                                        changedCount = changedCount + 1
                                    End If
                                End If
                            End If
                        Next fieldIndex
                    Else
                        isInitialHistory = True
                        ' 新規事案は、署ファイルの実際の見出し順に合わせて追加する。
                        ReDim targetData(1 To 1, 1 To officeMap.Count)
                        For fieldIndex = 1 To FIELD_COUNT
                            fieldName = CStr(fieldNames(fieldIndex))
                            If enabledFields.Exists(fieldName) Then
                                targetData(1, CLng(officeMap(fieldName))) = masterData(currentRow, CLng(masterMap(fieldName)))
                            End If
                        Next fieldIndex
                        officeSheet.Cells(nextOfficeRow, CLng(officeMap("共通番号"))).NumberFormat = "@"
                        officeSheet.Cells(nextOfficeRow, CLng(officeMap("枝番号"))).NumberFormat = "@"
                        officeSheet.Cells(nextOfficeRow, 1).Resize(1, officeMap.Count).Value2 = targetData
                        officeIndexes.Add Mid$(fullKey, Len(officeKey) + 2), nextOfficeRow
                        nextOfficeRow = nextOfficeRow + 1
                        addedCount = addedCount + 1
                        didWrite = True
                        historyRow = nextHistoryRow
                        nextHistoryRow = nextHistoryRow + 1
                        historyIndexes.Add fullKey, historyRow
                    End If

                    ' 基準は署ファイルの保存成功後に更新する（後でまとめて保存）。
                    ReDim historyRecord(1 To 1, 1 To FIELD_COUNT + 3)
                    historyRecord(1, 1) = officeKey
                    historyRecord(1, 2) = masterData(currentRow, CLng(masterMap("共通番号")))
                    historyRecord(1, 3) = masterData(currentRow, CLng(masterMap("枝番号")))
                    For fieldIndex = 1 To FIELD_COUNT
                        fieldName = CStr(fieldNames(fieldIndex))
                        If enabledFields.Exists(fieldName) Then
                            historyRecord(1, fieldIndex + 3) = masterData(currentRow, CLng(masterMap(fieldName)))
                        ElseIf Not isInitialHistory Then
                            historyRecord(1, fieldIndex + 3) = historyData(historyRow, fieldIndex + 3)
                        End If
                    Next fieldIndex
                    historySheet.Range(historySheet.Cells(historyRow, 2), historySheet.Cells(historyRow, 3)).NumberFormat = "@"
                    historySheet.Cells(historyRow, 1).Resize(1, FIELD_COUNT + 3).Value2 = historyRecord
                End If
            End If
        Next currentRow

        If didWrite Then
            officeBook.Save
            didSaveOffice = True
        End If
        officeBook.Close SaveChanges:=False
        Set officeBook = Nothing
        ' 署ごとに基準を保存。途中で止まった場合も成功済みの署を再転記しない。
        masterBook.Save
        Debug.Print officeName & " : 完了"
    Next settingRow

    historySheet.Visible = xlSheetVeryHidden
    masterBook.Save
    masterBook.Close SaveChanges:=False
    Set masterBook = Nothing
    Debug.Print "TRANSFER=PASS : 追加=" & addedCount & " 更新セル=" & changedCount
    MsgBox "5署への転記完了" & vbCrLf & "新規追加: " & addedCount & "件" & _
           vbCrLf & "更新セル: " & changedCount & "件", vbInformation
    Exit Sub

Failed:
    errorText = Err.Description
    On Error Resume Next
    If Not officeBook Is Nothing Then officeBook.Close SaveChanges:=False
    If Not masterBook Is Nothing Then masterBook.Close SaveChanges:=False
    On Error GoTo 0
    Debug.Print "TRANSFER=ERROR : " & errorText
    If didSaveOffice Then errorText = errorText & vbCrLf & "一部の署は保存済みの可能性があります。内容を確認してください。"
    MsgBox "転記を停止しました。" & vbCrLf & errorText, vbExclamation
End Sub

' Sheet2に5署への転記可否を設定する。既存の設定値は変更しない。
Public Sub InitializeTransferSettings()
    Dim sheet As Worksheet
    Dim rowIndex As Long
    Dim fieldName As String

    Set sheet = ThisWorkbook.Worksheets(FIELD_SHEET)
    sheet.Cells(1, 4).Value2 = "5署へ転記"
    For rowIndex = 2 To FIELD_COUNT + 1
        fieldName = Trim$(SafeText(sheet.Cells(rowIndex, 1).Value2))
        If Len(Trim$(SafeText(sheet.Cells(rowIndex, 4).Value2))) = 0 Then
            If Left$(fieldName, 2) = "メモ" Or Left$(fieldName, 2) = "発送" Then
                sheet.Cells(rowIndex, 4).Value2 = "×"
            Else
                sheet.Cells(rowIndex, 4).Value2 = "○"
            End If
        End If
    Next rowIndex
    Debug.Print "TRANSFER_SETTINGS=PASS"
End Sub

' -------- 以下は共通の小さな内部処理 --------

Private Function ReadMasterFields() As Variant
    Dim sheet As Worksheet
    Dim fields() As Variant
    Dim counts As Object
    Dim index As Long
    Dim fieldName As String
    Dim occurrence As Long

    Set sheet = ThisWorkbook.Worksheets(FIELD_SHEET)
    Set counts = CreateObject("Scripting.Dictionary")
    ReDim fields(1 To FIELD_COUNT)
    For index = 1 To FIELD_COUNT
        fieldName = Trim$(SafeText(sheet.Cells(index + 1, 1).Value2))
        If Len(fieldName) = 0 Then Err.Raise vbObjectError + 821, , "Sheet2の項目名が空欄です: A" & index + 1
        If counts.Exists(fieldName) Then
            occurrence = counts(fieldName) + 1
            counts(fieldName) = occurrence
            fields(index) = fieldName & "_" & occurrence
        Else
            counts.Add fieldName, 1
            fields(index) = fieldName
        End If
    Next index
    ReadMasterFields = fields
End Function

Private Function ReadHeaderMap(ByVal sheet As Worksheet, ByVal headerRow As Long) As Object
    Dim mapping As Object
    Dim counts As Object
    Dim columnIndex As Long
    Dim lastColumn As Long
    Dim name As String
    Dim occurrence As Long

    Set mapping = CreateObject("Scripting.Dictionary")
    Set counts = CreateObject("Scripting.Dictionary")
    lastColumn = sheet.Cells(headerRow, sheet.Columns.Count).End(xlToLeft).Column
    For columnIndex = 1 To lastColumn
        name = Trim$(SafeText(sheet.Cells(headerRow, columnIndex).Value2))
        If Len(name) > 0 Then
            If counts.Exists(name) Then
                occurrence = counts(name) + 1
                counts(name) = occurrence
                name = name & "_" & occurrence
            Else
                counts.Add name, 1
            End If
            mapping.Add name, columnIndex
        End If
    Next columnIndex
    Set ReadHeaderMap = mapping
End Function

Private Function ReadTransferFlags(ByRef fields As Variant) As Object
    Dim sheet As Worksheet
    Dim flags As Object
    Dim index As Long
    Dim mark As String
    Set sheet = ThisWorkbook.Worksheets(FIELD_SHEET)
    Set flags = CreateObject("Scripting.Dictionary")
    For index = 1 To FIELD_COUNT
        mark = Trim$(SafeText(sheet.Cells(index + 1, 4).Value2))
        If mark <> "○" And mark <> "×" Then Err.Raise vbObjectError + 822, , _
            "Sheet2のD" & index + 1 & "は○または×にしてください。最初にInitializeTransferSettingsを実行してください。"
        If mark = "○" Then flags.Add CStr(fields(index)), True
    Next index
    Set ReadTransferFlags = flags
End Function

Private Sub ValidateMasterFields(ByVal mapping As Object, ByRef fields As Variant)
    Dim index As Long
    For index = 1 To FIELD_COUNT
        RequireField mapping, CStr(fields(index)), "進捗整理簿【全体】"
    Next index
End Sub

Private Sub RequireField(ByVal mapping As Object, ByVal fieldName As String, ByVal name As String)
    If Not mapping.Exists(fieldName) Then Err.Raise vbObjectError + 823, , _
        "必要な列がありません: " & name & " / " & fieldName
End Sub

Private Function GetFieldPosition(ByRef fields As Variant, ByVal fieldName As String) As Long
    Dim index As Long
    For index = LBound(fields) To UBound(fields)
        If CStr(fields(index)) = fieldName Then
            GetFieldPosition = index
            Exit Function
        End If
    Next index
    Err.Raise vbObjectError + 824, , "Sheet2に項目がありません: " & fieldName
End Function

Private Function ReadTable(ByVal sheet As Worksheet, ByVal keyColumn As Long, ByVal headerRow As Long) As Variant
    Dim lastRow As Long
    Dim lastColumn As Long
    lastRow = sheet.Cells(sheet.Rows.Count, keyColumn).End(xlUp).Row
    lastColumn = sheet.Cells(headerRow, sheet.Columns.Count).End(xlToLeft).Column
    If lastRow < headerRow + 1 Then lastRow = headerRow + 1
    ReadTable = sheet.Range(sheet.Cells(1, 1), sheet.Cells(lastRow, lastColumn)).Value2
End Function

Private Function BuildKeyIndex(ByRef data As Variant, ByVal mapping As Object, _
                               ByVal fileName As String) As Object
    Dim keys As Object
    Dim rowIndex As Long
    Dim key As String
    Set keys = CreateObject("Scripting.Dictionary")
    For rowIndex = LEDGER_FIRST_DATA_ROW To UBound(data, 1)
        If Len(Trim$(SafeText(data(rowIndex, CLng(mapping("共通番号")))))) > 0 Then
            key = MakeKey(data(rowIndex, CLng(mapping("共通番号"))), _
                          data(rowIndex, CLng(mapping("枝番号"))))
            If keys.Exists(key) Then Err.Raise vbObjectError + 825, , _
                "同じ識別番号が複数あります: " & fileName & " / " & key
            keys.Add key, rowIndex
        End If
    Next rowIndex
    Set BuildKeyIndex = keys
End Function

Private Function MakeKey(ByVal number As Variant, ByVal branch As Variant) As String
    Dim numberText As String
    Dim branchText As String
    numberText = Trim$(SafeText(number))
    branchText = Trim$(SafeText(branch))
    If Len(numberText) <> 13 Or Len(branchText) <> 5 Then Err.Raise vbObjectError + 826, , _
        "共通番号13桁・枝番号5桁を確認してください: " & numberText & "-" & branchText
    If numberText Like "*[!0-9]*" Or branchText Like "*[!0-9]*" Then Err.Raise vbObjectError + 827, , _
        "識別番号に数字以外が含まれています: " & numberText & "-" & branchText
    MakeKey = numberText & "|" & branchText
End Function

Private Function SafeText(ByVal value As Variant) As String
    If IsError(value) Then Err.Raise vbObjectError + 828, , "セルにExcelのエラー値があります。"
    If IsNull(value) Or IsEmpty(value) Then
        SafeText = vbNullString
    Else
        SafeText = CStr(value)
    End If
End Function

Private Function ValuesEqual(ByVal first As Variant, ByVal second As Variant) As Boolean
    ValuesEqual = (SafeText(first) = SafeText(second))
End Function

Private Function GetConfiguredPath(ByVal settingRow As Long) As String
    Dim sheet As Worksheet
    Dim folderPath As String
    Dim fileName As String
    Set sheet = ThisWorkbook.Worksheets(ADDRESS_SHEET)
    folderPath = Trim$(SafeText(sheet.Cells(settingRow, 2).Value2))
    fileName = Trim$(SafeText(sheet.Cells(settingRow, 3).Value2))
    If Len(folderPath) = 0 Or Len(fileName) = 0 Then Err.Raise vbObjectError + 829, , _
        "Addressのパスまたはファイル名が空欄です: " & settingRow & "行目"
    ' 拡張子付きはそのまま使い、未指定時のみxlsxを補完する。
    If InStrRev(fileName, ".") = 0 Then fileName = fileName & ".xlsx"
    GetConfiguredPath = AddPathSeparator(folderPath) & fileName
End Function

' 各ファイル内の実際のシート名はAddressのD列から取得する。
Private Function GetConfiguredSheetName(ByVal settingRow As Long) As String
    Dim sheetName As String
    sheetName = Trim$(SafeText(ThisWorkbook.Worksheets(ADDRESS_SHEET).Cells(settingRow, 4).Value2))
    If Len(sheetName) = 0 Then
        Err.Raise vbObjectError + 831, "GetConfiguredSheetName", _
            "AddressのD" & settingRow & "にシート名を入力してください。"
    End If
    GetConfiguredSheetName = sheetName
End Function

Private Function AddPathSeparator(ByVal folderPath As String) As String
    Dim separator As String
    folderPath = Trim$(folderPath)
    separator = "\"
    If LCase$(Left$(folderPath, 4)) = "http" Then separator = "/"
    If Right$(folderPath, 1) <> "\" And Right$(folderPath, 1) <> "/" Then
        folderPath = folderPath & separator
    End If
    AddPathSeparator = folderPath
End Function

Private Function GetOfficeName(ByVal settingRow As Long) As String
    Dim fileName As String
    Dim startIndex As Long
    Dim endIndex As Long
    fileName = SafeText(ThisWorkbook.Worksheets(ADDRESS_SHEET).Cells(settingRow, 3).Value2)
    startIndex = InStr(fileName, "【")
    endIndex = InStr(fileName, "】")
    If startIndex = 0 Or endIndex <= startIndex Then Err.Raise vbObjectError + 830, , _
        "Addressのファイル名から署名を特定できません: " & fileName
    GetOfficeName = Mid$(fileName, startIndex + 1, endIndex - startIndex - 1)
End Function

Private Function GetHistorySheet(ByVal masterBook As Workbook, ByRef fields As Variant) As Worksheet
    Dim sheet As Worksheet
    Dim fieldIndex As Long
    On Error Resume Next
    Set sheet = masterBook.Worksheets(HISTORY_SHEET)
    On Error GoTo 0
    If sheet Is Nothing Then
        Set sheet = masterBook.Worksheets.Add(After:=masterBook.Worksheets(masterBook.Worksheets.Count))
        sheet.Name = HISTORY_SHEET
        sheet.Cells(1, 1).Value2 = "署識別名"
        sheet.Cells(1, 2).Value2 = "共通番号"
        sheet.Cells(1, 3).Value2 = "枝番号"
        For fieldIndex = 1 To FIELD_COUNT
            sheet.Cells(1, fieldIndex + 3).Value2 = fields(fieldIndex)
        Next fieldIndex
        sheet.Columns("B:C").NumberFormat = "@"
    Else
        For fieldIndex = 1 To FIELD_COUNT
            If SafeText(sheet.Cells(1, fieldIndex + 3).Value2) <> CStr(fields(fieldIndex)) Then
                Err.Raise vbObjectError + 831, , "転記基準の列構成が現在のSheet2と一致しません。"
            End If
        Next fieldIndex
    End If
    Set GetHistorySheet = sheet
End Function