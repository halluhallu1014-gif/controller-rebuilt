Attribute VB_Name = "FiveFileSearch"
Option Explicit

' 既存の「検索」シートに合わせた画面配置。
Private Const SEARCH_SHEET As String = "検索"
Private Const INPUT_CELL As String = "E9"
Private Const COUNT_CELL As String = "C17"
Private Const PREVIEW_CELL As String = "C35"
Private Const FIRST_RESULT_ROW As Long = 20
Private Const PAGE_SIZE As Long = 14

' 検索結果を記憶し、ページ切替時には全体ファイルを開き直さない。
Private searchMatches As Collection
Private searchPage As Long

' 入力された氏名を部分一致検索し、1ページ目を表示する。
Public Sub SearchCases()
    Dim records As Collection
    Dim record As CaseRecord
    Dim queryText As String
    Dim index As Long
    Dim previousScreenUpdating As Boolean

    On Error GoTo SearchFailed
    queryText = Trim$(CStr(ThisWorkbook.Worksheets(SEARCH_SHEET).Range(INPUT_CELL).Value2))
    previousScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Set searchMatches = New Collection
    searchPage = 1

    ' 6ファイル方式では全体ファイルを正本として検索する。
    Set records = LoadMasterRecords()
    For index = 1 To records.Count
        Set record = records(index)
        If Len(queryText) = 0 Or _
           InStr(1, CStr(record.Value("氏名法人名")), queryText, vbTextCompare) > 0 Then
            searchMatches.Add record
        End If
    Next index

    RenderSearchResults
    Application.ScreenUpdating = previousScreenUpdating
    Debug.Print "検索結果 : " & searchMatches.Count & "件"
    Debug.Print "RESULT=PASS"
    Exit Sub

SearchFailed:
    Application.ScreenUpdating = previousScreenUpdating
    MsgBox "検索に失敗しました。" & vbCrLf & Err.Description, vbExclamation, "検索"
    Debug.Print "RESULT=ERROR : " & Err.Description
End Sub

' 検索結果の次ページを表示する。再読み込みはしない。
Public Sub SearchNextPage()
    If searchMatches Is Nothing Then Exit Sub
    If searchPage * PAGE_SIZE < searchMatches.Count Then
        searchPage = searchPage + 1
        RenderSearchResults
    End If
End Sub

' 検索結果の前ページを表示する。再読み込みはしない。
Public Sub SearchPreviousPage()
    If searchMatches Is Nothing Then Exit Sub
    If searchPage > 1 Then
        searchPage = searchPage - 1
        RenderSearchResults
    End If
End Sub

' 検索結果一覧を、既存の結合セルと書式を保って更新する。
Private Sub RenderSearchResults()
    Dim searchSheet As Worksheet
    Dim record As CaseRecord
    Dim resultRow As Long
    Dim itemIndex As Long
    Dim pageCount As Long
    Dim displayPage As Long

    Set searchSheet = ThisWorkbook.Worksheets(SEARCH_SHEET)

    ' 古い検索結果だけを消す。見出し・罫線・書式は残す。
    For resultRow = FIRST_RESULT_ROW To FIRST_RESULT_ROW + PAGE_SIZE - 1
        searchSheet.Cells(resultRow, "C").MergeArea.ClearContents
        searchSheet.Cells(resultRow, "F").MergeArea.ClearContents
        searchSheet.Cells(resultRow, "K").MergeArea.ClearContents
        searchSheet.Cells(resultRow, "P").MergeArea.ClearContents
        searchSheet.Cells(resultRow, "T").MergeArea.ClearContents
    Next resultRow
    searchSheet.Range(PREVIEW_CELL).MergeArea.ClearContents

    For resultRow = FIRST_RESULT_ROW To FIRST_RESULT_ROW + PAGE_SIZE - 1
        itemIndex = (searchPage - 1) * PAGE_SIZE + resultRow - FIRST_RESULT_ROW + 1
        If itemIndex > searchMatches.Count Then Exit For

        Set record = searchMatches(itemIndex)
        searchSheet.Cells(resultRow, "C").Value2 = record.Value("署名")
        searchSheet.Cells(resultRow, "F").Value2 = record.Value("氏名法人名")
        ' 「ふりがな」は現行の48項目にないため、K列は空欄にする。
        searchSheet.Cells(resultRow, "P").Value2 = record.Value("申請書種別")
        searchSheet.Cells(resultRow, "T").Value2 = record.Value("処理区分")
    Next resultRow

    pageCount = (searchMatches.Count + PAGE_SIZE - 1) \ PAGE_SIZE
    displayPage = searchPage
    If searchMatches.Count = 0 Then displayPage = 0
    searchSheet.Range(COUNT_CELL).Value2 = "検索結果 " & searchMatches.Count & _
        "件  " & displayPage & " / " & pageCount & "ページ"
End Sub

' 一覧内のセルを選択して実行すると、その行の事案を返す。
' 詳細画面への連携に使うための共通入口。
Public Function GetSelectedSearchRecord() As CaseRecord
    Dim resultRow As Long
    Dim itemIndex As Long

    If searchMatches Is Nothing Then Exit Function
    If Not ActiveSheet Is ThisWorkbook.Worksheets(SEARCH_SHEET) Then Exit Function

    resultRow = ActiveCell.Row
    If resultRow < FIRST_RESULT_ROW Then Exit Function
    If resultRow >= FIRST_RESULT_ROW + PAGE_SIZE Then Exit Function

    itemIndex = (searchPage - 1) * PAGE_SIZE + resultRow - FIRST_RESULT_ROW + 1
    If itemIndex > searchMatches.Count Then Exit Function
    Set GetSelectedSearchRecord = searchMatches(itemIndex)
End Function

' 選択した結果の概要を画面下段に表示するテスト用処理。
Public Sub PreviewSelectedSearchRecord()
    Dim record As CaseRecord

    Set record = GetSelectedSearchRecord()
    If record Is Nothing Then
        MsgBox "検索結果の行（20～33行）を選択してください。", vbInformation, "検索"
        Exit Sub
    End If

    ThisWorkbook.Worksheets(SEARCH_SHEET).Range(PREVIEW_CELL).Value2 = _
        CStr(record.Value("氏名法人名")) & " / " & _
        CStr(record.Value("年分・事業年度等")) & " / " & _
        CStr(record.Value("申請書種別"))
End Sub
