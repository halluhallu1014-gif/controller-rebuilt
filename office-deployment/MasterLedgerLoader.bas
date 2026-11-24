Attribute VB_Name = "MasterLedgerLoader"
Option Explicit

Private Const FIELD_COUNT As Long = 48
Private Const HEADER_ROW As Long = 5
Private Const FIRST_DATA_ROW As Long = HEADER_ROW + 1

' Addressのzentai行を探し、全体ファイルのフルパスを返す。
Public Function GetMasterLedgerPath() As String
    Dim settings As Worksheet
    Dim settingRow As Long
    Dim lastRow As Long
    Dim folderPath As String
    Dim fileName As String
    Dim separator As String

    Set settings = ThisWorkbook.Worksheets("Address")
    lastRow = settings.Cells(settings.Rows.Count, "A").End(xlUp).Row
    For settingRow = 2 To lastRow
        If StrComp(Trim$(CStr(settings.Cells(settingRow, "A").Value2)), "zentai", vbTextCompare) = 0 Then
            folderPath = Trim$(CStr(settings.Cells(settingRow, "B").Value2))
            fileName = Trim$(CStr(settings.Cells(settingRow, "C").Value2))
            If Len(folderPath) = 0 Or Len(fileName) = 0 Then
                Err.Raise vbObjectError + 901, "GetMasterLedgerPath", "Addressのzentai行にパスまたはファイル名がありません。"
            End If
            ' 拡張子付きはそのまま使い、未指定時のみxlsxを補完する。
            If InStrRev(fileName, ".") = 0 Then fileName = fileName & ".xlsx"
            separator = "\"
            If LCase$(Left$(folderPath, 4)) = "http" Then separator = "/"
            If Right$(folderPath, 1) <> "\" And Right$(folderPath, 1) <> "/" Then
                folderPath = folderPath & separator
            End If
            GetMasterLedgerPath = folderPath & fileName
            Exit Function
        End If
    Next settingRow
    Err.Raise vbObjectError + 902, "GetMasterLedgerPath", "Addressにzentai行がありません。"
End Function

' Addressのzentai行のD列から全体管理簿のシート名を取得する。
Public Function GetMasterLedgerSheetName() As String
    Dim settings As Worksheet
    Dim settingRow As Long
    Dim lastRow As Long
    Dim sheetName As String

    Set settings = ThisWorkbook.Worksheets("Address")
    lastRow = settings.Cells(settings.Rows.Count, "A").End(xlUp).Row
    For settingRow = 2 To lastRow
        If StrComp(Trim$(CStr(settings.Cells(settingRow, "A").Value2)), "zentai", vbTextCompare) = 0 Then
            sheetName = Trim$(CStr(settings.Cells(settingRow, "D").Value2))
            If Len(sheetName) = 0 Then
                Err.Raise vbObjectError + 906, "GetMasterLedgerSheetName", _
                    "AddressのD" & settingRow & "に全体管理簿のシート名を入力してください。"
            End If
            GetMasterLedgerSheetName = sheetName
            Exit Function
        End If
    Next settingRow
    Err.Raise vbObjectError + 907, "GetMasterLedgerSheetName", "Addressにzentai行がありません。"
End Function

' Sheet2の項目名を、管理簿と同じ順番で取得する。
' 重複する項目名には2件目以降に連番を付ける。
Public Function GetMasterFieldNames() As Variant
    Dim definitions As Worksheet
    Dim names(1 To FIELD_COUNT) As Variant
    Dim counts As Object
    Dim fieldName As String
    Dim columnIndex As Long
    Dim duplicateCount As Long

    Set definitions = ThisWorkbook.Worksheets("Sheet2")
    Set counts = CreateObject("Scripting.Dictionary")
    For columnIndex = 1 To FIELD_COUNT
        fieldName = Trim$(CStr(definitions.Cells(columnIndex + 1, "A").Value2))
        If Len(fieldName) = 0 Then
            Err.Raise vbObjectError + 903, "GetMasterFieldNames", "Sheet2の項目が空欄です: A" & (columnIndex + 1)
        End If
        If counts.Exists(fieldName) Then
            duplicateCount = CLng(counts(fieldName)) + 1
            counts(fieldName) = duplicateCount
            names(columnIndex) = fieldName & "_" & duplicateCount
        Else
            counts.Add fieldName, 1
            names(columnIndex) = fieldName
        End If
    Next columnIndex
    GetMasterFieldNames = names
End Function

' 全体ファイルの見出しがSheet2の48項目と一致するか確認する。
Public Sub ValidateMasterHeaders(ByVal targetSheet As Worksheet)
    Dim definitions As Worksheet
    Dim columnIndex As Long
    Dim expectedName As String
    Dim actualName As String
    Set definitions = ThisWorkbook.Worksheets("Sheet2")
    For columnIndex = 1 To FIELD_COUNT
        expectedName = CStr(definitions.Cells(columnIndex + 1, "A").Value2)
        actualName = CStr(targetSheet.Cells(HEADER_ROW, columnIndex).Value2)
        If expectedName <> actualName Then
            Err.Raise vbObjectError + 904, "ValidateMasterHeaders", _
                "全体ファイルの見出しが一致しません: " & columnIndex & "列（" & actualName & "）"
        End If
    Next columnIndex
End Sub

' 全体ファイルから事案を読み込み、Collection(CaseRecord)として返す。
Public Function LoadMasterRecords() As Collection
    Dim targetBook As Workbook
    Dim targetSheet As Worksheet
    Dim records As Collection
    Dim data As Variant
    Dim fieldNames As Variant
    Dim record As CaseRecord
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim errorText As String

    On Error GoTo LoadFailed
    Set records = New Collection
    fieldNames = GetMasterFieldNames()
    Set targetBook = Workbooks.Open(Filename:=GetMasterLedgerPath(), ReadOnly:=True, UpdateLinks:=0)
    Set targetSheet = targetBook.Worksheets(GetMasterLedgerSheetName())
    ValidateMasterHeaders targetSheet
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    If lastRow >= FIRST_DATA_ROW Then
        data = targetSheet.Range(targetSheet.Cells(FIRST_DATA_ROW, 1), targetSheet.Cells(lastRow, FIELD_COUNT)).Value2
        For rowIndex = LBound(data, 1) To UBound(data, 1)
            If Len(Trim$(CStr(data(rowIndex, 1)))) > 0 Then
                Set record = New CaseRecord
                record.LoadData data, rowIndex, fieldNames
                record.SourceUses1904 = targetBook.Date1904
                records.Add record
            End If
        Next rowIndex
    End If
    targetBook.Close SaveChanges:=False
    Set targetBook = Nothing
    Set LoadMasterRecords = records
    Exit Function
LoadFailed:
    errorText = Err.Description
    On Error Resume Next
    If Not targetBook Is Nothing Then targetBook.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 905, "LoadMasterRecords", "全体ファイルの読込失敗: " & errorText
End Function

' 単体確認用。既存の5ファイル読込テストとは独立している。
Public Sub TestLoadMasterRecords()
    Dim records As Collection
    On Error GoTo Failed
    Set records = LoadMasterRecords()
    Debug.Print "全体 : " & records.Count & "件"
    Debug.Print "MASTER_LOAD=PASS"
    Exit Sub
Failed:
    Debug.Print "MASTER_LOAD=ERROR : " & Err.Description
    MsgBox Err.Description, vbExclamation, "全体読み込み"
End Sub

' 日付表示が必要な項目を判定する。配置先の変更には影響しない。
Public Function IsLedgerDateField(ByVal fieldName As String) As Boolean
    Select Case fieldName
        Case "申請日", "処理期限", "署引継日", "取下書収受日", _
             "処理年月日", "施行年月日", "仮装経理に係る最終処理年月日"
            IsLedgerDateField = True
    End Select
End Function
