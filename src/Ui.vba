Option Explicit

Private people As Object
Private personTokens As Object
Private caseTokens As Object
Private baseline As Object
Private pageIndex As Long
Private relatedCount As Long
Private relatedPageIndex As Long
Private draftMasked As String
Private isNewMode As Boolean
Private hasDetail As Boolean
Private baselineYear As String

Public Sub SetupController()
    Dim sheetName As Variant
    Dim ws As Worksheet
    Dim candidate As Worksheet
    Dim index As Long
    Dim message As String
    If MsgBox("空の新規xlsmに操作画面と設定例を作成します。既存シートと同名の場合は中止します。", vbOKCancel, "初期構築") <> vbOK Then Exit Sub
    On Error GoTo Failed
    For Each sheetName In Array("Settings", "Fields", "検索", "詳細", "Calendar", "Recovery", "MoveRecovery")
        For Each candidate In ThisWorkbook.Worksheets
            If candidate.Name = CStr(sheetName) Then
                Notify "既存シート " & CStr(sheetName) & " があるため中止しました。初期構築は空の別ブックで実行してください。既存内容は削除しません。"
                Exit Sub
            End If
        Next
    Next
    For Each sheetName In Array("Settings", "Fields", "検索", "詳細", "Calendar", "Recovery", "MoveRecovery")
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = CStr(sheetName)
        ws.Cells.Font.Name = "Yu Gothic"
        ws.Cells.Font.Size = 11
        ws.Columns("A:P").ColumnWidth = 16
    Next
    Set ws = Sheet("Settings")
    ws.Range("A1:B1").Value2 = Array("Setting", "Value")
    ws.Range("D1:I1").Value2 = Array("Year", "StaffAddress", "Sheet", "Enabled", "AllowWrite", "GeneralAddress")
    ws.Range("J1:L1").Value2 = Array("PhonePrefix", "AreaLength", "ExchangeLength")
    ws.Range("A2:B8").NumberFormat = "@"
    ws.Range("A2:B2").Value2 = Array("ControllerRole", "General")
    ws.Range("A3:B3").Value2 = Array("SharePointSite", "")
    ws.Range("A4:B4").Value2 = Array("SharePointExclusive", "No")
    ws.Range("A5:B5").Value2 = Array("BackupFolder", "")
    ws.Range("A6:B6").Value2 = Array("BusinessState", "")
    ws.Range("A7:B7").Value2 = Array("Transport", "Local")
    ws.Range("D2:I2").NumberFormat = "@"
    ws.Range("D2:I2").Value2 = Array("2026", "", "管理簿", "No", "No", "")
    ws.Range("J2:L3").NumberFormat = "@"
    ws.Range("J2:L2").Value2 = Array("03", "2", "4")
    ws.Range("J3:L3").Value2 = Array("06", "2", "4")
    ws.Range("A11").Value2 = "Staff=常勤 / General=非常勤。管理機能はStaffのみ。年度表は行追加で拡張。"
    ws.Range("A12").Value2 = "URLは直接ファイルURL。実URL・個人情報は公開しない。移動候補は同一サイト・同一ドライブ内。"
    Set ws = Sheet("Fields")
    ws.Range("A1:F1").Value2 = Array("FieldId", "Header", "Type", "Sheet", "Cell", "Required")
    ws.Range("H1:J1").Value2 = Array("Year", "FieldId", "HeaderOverride")
    ws.Range("A2:J100").NumberFormat = "@"
    ws.Range("A2:F2").Value2 = Array("MaskedNumber", "MaskedNumber", "Text", "詳細", "C20", "Yes")
    ws.Range("A3:F3").Value2 = Array("Branch", "枝番", "Text", "詳細", "C21", "Yes")
    ws.Range("A4:F4").Value2 = Array("Office", "署", "Text", "詳細", "C22", "No")
    ws.Range("A5:F5").Value2 = Array("PersonName", "氏名", "Text", "詳細", "C23", "Yes")
    ws.Range("A6:F6").Value2 = Array("Furigana", "氏名(フリガナ)", "Text", "詳細", "C24", "No")
    ws.Range("A7:F7").Value2 = Array("DocumentName", "文書名", "Text", "詳細", "C25", "No")
    ws.Range("A8:F8").Value2 = Array("TaxType", "税目", "Text", "詳細", "C26", "No")
    ws.Range("A9:F9").Value2 = Array("TaxYear", "年分", "Text", "詳細", "C27", "No")
    ws.Range("A10:F10").Value2 = Array("Status", "状態", "Text", "詳細", "C28", "No")
    ws.Range("A11:F11").Value2 = Array("ReceivedDate", "収受日", "Date", "詳細", "C29", "No")
    ws.Range("A12:F12").Value2 = Array("DraftDate", "起案日", "Date", "詳細", "C30", "No")
    ws.Range("A13:F13").Value2 = Array("ApprovedDate", "決裁日", "Date", "詳細", "C31", "No")
    ws.Range("A14:F14").Value2 = Array("IssuedDate", "施行日", "Date", "詳細", "C32", "No")
    ws.Range("A15:F15").Value2 = Array("Memo1", "メモ1", "Text", "詳細", "C33", "No")
    ws.Range("A16:F16").Value2 = Array("Memo2", "メモ2", "Text", "詳細", "C34", "No")
    ws.Range("A17:F17").Value2 = Array("Memo3", "メモ3", "Text", "詳細", "C35", "No")
    ws.Range("A18:F18").Value2 = Array("ApprovalNumber", "決裁番号", "Text", "詳細", "C36", "No")
    ws.Range("A19:F19").Value2 = Array("InquiryNumber", "お問い合わせ番号", "Text", "詳細", "C37", "No")
    ws.Range("A20:F20").Value2 = Array("PostalCode", "郵便番号", "Text", "詳細", "C38", "No")
    ws.Range("A21:F21").Value2 = Array("Address", "住所", "Text", "詳細", "C39", "No")
    ws.Range("A22:F22").Value2 = Array("PhoneNumber", "電話番号", "Text", "詳細", "C40", "No")
    Set ws = Sheet("検索")
    ws.Range("A1").Value2 = "CONTROLLER 人物検索"
    ws.Range("A2").Value2 = "氏名部分一致"
    ws.Range("D2").Value2 = "年度（空欄=全年度）"
    ws.Range("G2").Value2 = "追加FieldId"
    ws.Range("J2").Value2 = "追加条件（部分一致）"
    ws.Range("A7:E7").Value2 = Array("選択ID", "氏名", "枝番", "全年度事案件数", "確認状態")
    AddUiName "SearchName", ws.Range("B3")
    AddUiName "SearchYear", ws.Range("E3")
    AddUiName "SearchField", ws.Range("H3")
    AddUiName "SearchValue", ws.Range("K3")
    AddUiName "SearchResults", ws.Range("A8:E19")
    AddUiName "SearchState", ws.Range("A5")
    ws.Range("B3:K3").NumberFormat = "@"
    AddButton ws, "最新を再読込", "StartController", 22, 1
    AddButton ws, "検索", "FindPeople", 22, 3
    AddButton ws, "前へ", "PreviousPage", 22, 5
    AddButton ws, "次へ", "NextPage", 22, 7
    AddButton ws, "選択人物を開く", "OpenPerson", 24, 1
    AddButton ws, "新規登録", "NewCase", 24, 4
    AddButton ws, "業務開始（常勤）", "StartBusiness", 26, 1
    AddButton ws, "業務終了（常勤）", "FinishBusiness", 26, 4
    AddButton ws, "移動結果を確認", "ConfirmMoveResult", 28, 1
    AddButton ws, "バックアップを復元", "RestoreBackup", 28, 4
    Set ws = Sheet("詳細")
    ws.Range("A1").Value2 = "CONTROLLER 事案詳細"
    ws.Range("A3").Value2 = "対象事務年度"
    AddUiName "DetailYear", ws.Range("C3")
    ws.Range("C3").NumberFormat = "@"
    AddUiName "DetailState", ws.Range("A5")
    ws.Range("A7").Value2 = "18桁番号貼付欄（新規専用）"
    AddUiName "NumberPaste", ws.Range("C8")
    ws.Range("C8").NumberFormat = "@"
    ws.Range("A10").Value2 = "外部テキスト貼付欄（新規専用）"
    AddUiName "ImportPaste", ws.Range("C11")
    ws.Range("C11").NumberFormat = "@"
    ws.Range("C11").WrapText = True
    ws.Rows(11).RowHeight = 45
    ws.Range("J7:P7").Value2 = Array("選択ID", "事務年度", "年分", "税目", "文書名", "状態", "確認事項")
    AddUiName "RelatedStart", ws.Range("J8")
    ws.Range("A18").Value2 = "選択事案の入力欄"
    For index = 2 To 22
        ws.Cells(index + 18, 1).Value2 = Sheet("Fields").Cells(index, 2).Value2
        ws.Cells(index + 18, 3).NumberFormat = "@"
        If Sheet("Fields").Cells(index, 3).Value2 = "Date" Then ws.Cells(index + 18, 3).NumberFormat = "yyyy/mm/dd"
        ws.Cells(index + 18, 3).Interior.Color = RGB(235, 245, 255)
    Next
    AddButton ws, "番号取込", "ApplyNumberPaste", 8, 5
    AddButton ws, "クリップボード取込", "ImportClipboard", 12, 5
    AddButton ws, "貼付解析", "ParsePastedText", 14, 5
    AddButton ws, "選択事案を開く", "OpenRelatedCase", 3, 10
    AddButton ws, "保存", "SaveDetail", 43, 1
    AddButton ws, "保存結果を確認", "ConfirmSaveResult", 43, 4
    AddButton ws, "氏名コピー", "CopyPersonName", 45, 1
    AddButton ws, "住所コピー", "CopyAddress", 45, 4
    AddButton ws, "ふりがなコピー", "CopyFurigana", 47, 1
    AddButton ws, "番号コピー", "CopyNumber", 47, 4
    AddButton ws, "日付カレンダー", "OpenCalendar", 49, 1
    ws.Columns("C").ColumnWidth = 42
    Set ws = Sheet("Calendar")
    ws.Range("B4:H4").Value2 = Array("日", "月", "火", "水", "木", "金", "土")
    AddButton ws, "前月", "PreviousMonth", 12, 2
    AddButton ws, "翌月", "NextMonth", 12, 5
    AddButton ws, "選択日を入力", "PickCalendarDate", 14, 2
    AddButton ws, "日付を空欄にする", "ClearCalendarDate", 14, 5
    Sheet("Recovery").Visible = xlSheetVeryHidden
    Sheet("MoveRecovery").Visible = xlSheetVeryHidden
    For Each sheetName In Array("Settings", "Fields", "検索", "詳細", "Calendar")
        Set ws = Sheet(CStr(sheetName))
        ws.Rows(1).Font.Bold = True
        ws.Rows(1).Interior.Color = RGB(215, 230, 245)
    Next
    isBusy = False
    ApplyUiLayout
    Exit Sub
    Exit Sub
Failed:
    Notify SafeFailure("初期構築", Err.Number) & " 自動削除は行いません。作成途中の別ブックを確認してください。"
End Sub

Private Sub AddUiName(ByVal rangeName As String, ByVal target As Range)
    ThisWorkbook.Names.Add Name:=rangeName, RefersTo:="='" & target.Parent.Name & "'!" & target.Address
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal caption As String, ByVal macroName As String, ByVal rowIndex As Long, ByVal colIndex As Long)
    Dim button As Object
    Set button = ws.Buttons.Add(ws.Cells(rowIndex, colIndex).Left, ws.Cells(rowIndex, colIndex).Top, 180, 25)
    button.Caption = caption
    button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Sub

Public Function UiRange(ByVal rangeName As String) As Range
    Set UiRange = ThisWorkbook.Names(rangeName).RefersToRange
End Function

Public Function IsReservedUiCell(ByVal target As Range) As Boolean
    Dim rangeName As Variant
    Dim area As Range
    For Each rangeName In Array("SearchName", "SearchYear", "SearchField", "SearchValue", "SearchResults", "SearchState", "DetailYear", "DetailState", "NumberPaste", "ImportPaste", "RelatedStart")
        Set area = UiRange(CStr(rangeName))
        If CStr(rangeName) = "RelatedStart" Then
            If HasRelatedViewport() Then
                Set area = RelatedArea()
            Else
                Set area = area.Resize(area.Parent.Rows.Count - area.Row + 1, 7)
            End If
        ElseIf area.CountLarge = 1 Then
            If area.MergeCells Then Set area = area.MergeArea
        End If
        If area.Parent Is target.Parent Then
            If Not Intersect(area, target) Is Nothing Then
                IsReservedUiCell = True
                Exit Function
            End If
        End If
    Next
End Function

Public Sub StartController()
    ExecuteAction "Start"
End Sub

Public Sub FindPeople()
    ExecuteAction "Search"
End Sub

Public Sub PreviousPage()
    ExecuteAction "Previous"
End Sub

Public Sub NextPage()
    ExecuteAction "Next"
End Sub

Public Sub OpenPerson()
    ExecuteAction "Person"
End Sub

Public Sub OpenRelatedCase()
    ExecuteAction "Case"
End Sub

Public Sub NewCase()
    ExecuteAction "New"
End Sub

Public Sub SaveDetail()
    ExecuteAction "Save"
End Sub

Public Sub ConfirmSaveResult()
    ExecuteAction "Confirm"
End Sub

Public Sub ApplyNumberPaste()
    ExecuteAction "NumberPaste"
End Sub

Public Sub ImportClipboard()
    ExecuteAction "ClipboardImport"
End Sub

Public Sub ParsePastedText()
    ExecuteAction "PasteImport"
End Sub

Public Sub CopyPersonName()
    CopyField "PersonName"
End Sub

Public Sub CopyAddress()
    CopyField "Address"
End Sub

Public Sub CopyFurigana()
    CopyField "Furigana"
End Sub

Public Sub CopyNumber()
    Dim number As String
    Dim masked As String
    If Not hasDetail Then Exit Sub
    masked = draftMasked
    If Not isNewMode Then masked = currentCase.values("MaskedNumber")
    If Not RestoreNumber(masked, number) Or Not IsDigits(FieldCell("Branch").Value2, 5) Then
        Notify "番号を復元できないためコピーを中止しました。"
        Exit Sub
    End If
    If Not isNewMode Then
        If CStr(FieldCell("Branch").Value2) <> currentCase.values("Branch") Then
            Notify "枝番が変更されているためコピーを中止しました。"
            Exit Sub
        End If
    End If
    If Not WriteClipboard(number & CStr(FieldCell("Branch").Value2)) Then
        Notify "番号をコピーできません。Officeのクリップボード部品を利用できるか確認してください。番号をセルへ露出する代替処理は行いません。"
    End If
End Sub

Private Sub CopyField(ByVal fieldId As String)
    On Error GoTo Failed
    If Not hasDetail Then Exit Sub
    If Not WriteClipboard(CStr(FieldCell(fieldId).Value2)) Then GoTo Failed
    Exit Sub
Failed:
    Notify "コピーできません。Officeのクリップボード部品を確認してください。"
End Sub

Private Sub ExecuteAction(ByVal actionName As String)
    Dim message As String
    Dim values As Object
    Dim record As CaseRecord
    Dim id As Variant
    Dim text As String
    Dim token As String
    Dim groupKey As String
    Dim number As String
    Dim targetKey As String
    If Not BeginAction() Then Exit Sub
    On Error GoTo Failed
    If actionName = "Start" Then
        If PendingSave() Then
            If LoadConfiguration(message) Then
                If Cases Is Nothing Then Set Cases = NewDictionary()
                Set yearErrors = NewDictionary()
                For Each id In years.Keys
                    yearErrors.Add id, "保存結果確認中のため年度は未読込です。照合後に最新を再読込してください。"
                Next
                message = "保存結果が未確認です。保存結果を確認ボタンで照合してください。入力は保持しています。"
            End If
            GoTo Finished
        End If
        If Not UiCanLeaveDetail() Then GoTo Finished
        If Not LoadConfiguration(message) Then GoTo Finished
        ReloadCases message
        If HasRelatedViewport() Then PutValue UiRange("SearchState"), "読込件数：" & Cases.Count & "件 / 読込失敗年度：" & yearErrors.Count & "。氏名または年度を指定して検索してください。"
        Set currentCase = Nothing
        hasDetail = False
        BuildSearch
        GoTo Finished
    End If
    If fields Is Nothing Or Cases Is Nothing Then
        message = "最初にStartController（最新を再読込）を実行してください。"
        GoTo Finished
    End If
    Select Case actionName
        Case "Search"
            BuildSearch
            If HasRelatedViewport() Then PutValue UiRange("SearchState"), "氏名の部分一致で人物を検索します。年度が空欄なら全年度、関連件数は全年度合計です。"
        Case "Previous"
            If pageIndex > 0 Then pageIndex = pageIndex - 1
            DrawSearchPage
        Case "Next"
            If people Is Nothing Then GoTo Finished
            If (pageIndex + 1) * 12 < people.Count Then pageIndex = pageIndex + 1
            DrawSearchPage
        Case "Person"
            token = SelectedToken(UiRange("SearchResults"))
            If Len(token) = 0 Then GoTo Finished
            If personTokens Is Nothing Then GoTo Finished
            If Not personTokens.Exists(token) Then GoTo Finished
            If Not UiCanLeaveDetail() Then GoTo Finished
            groupKey = personTokens(token)
            ShowRelated groupKey
        Case "Case"
            If caseTokens Is Nothing Then GoTo Finished
            token = SelectedToken(RelatedArea())
            If Len(token) = 0 Then GoTo Finished
            If Not caseTokens.Exists(token) Then GoTo Finished
            If Not UiCanLeaveDetail() Then GoTo Finished
            ShowCase CStr(caseTokens(token))
        Case "New"
            If Not UiCanLeaveDetail() Then GoTo Finished
            Set currentCase = Nothing
            isNewMode = True
            hasDetail = True
            draftMasked = ""
            UiRange("DetailYear").MergeArea.ClearContents
            Set values = NewDictionary()
            For Each id In fields.Keys
                values.Add id, ""
            Next
            DisplayValues values
            ClearRelated
            PutValue UiRange("DetailState"), "新規登録：対象年度を指定し、18桁番号を貼付してください。"
            UiRange("NumberPaste").MergeArea.ClearContents
            UiRange("ImportPaste").MergeArea.ClearContents
            RememberBaseline
            ShowDetailScreen
        Case "NumberPaste"
            If Not isNewMode Or Not hasDetail Then
                message = "番号貼付は新規登録時だけ利用できます。"
                GoTo Finished
            End If
            text = CStr(UiRange("NumberPaste").Value2)
            If Len(text) = 0 Then
                If Not ReadClipboard(text) Then
                    message = "NumberPaste欄へCtrl+Vで18桁文字列を貼り付けてください。"
                    GoTo Finished
                End If
            End If
            If Not IsDigits(text, 18) Then
                message = "番号は半角数字18桁の文字列です。先頭0を保持した状態で貼り付けてください。"
                GoTo Finished
            End If
            draftMasked = MaskNumber(Left$(text, 13))
            PutValue FieldCell("MaskedNumber"), "*************"
            PutValue FieldCell("Branch"), Right$(text, 5)
            UiRange("NumberPaste").MergeArea.ClearContents
            text = ""
            message = "番号を取り込み、貼付欄を消去しました。管理簿は保存していません。"
        Case "ClipboardImport", "PasteImport"
            If Not isNewMode Or Not hasDetail Then
                message = "外部テキスト取込は新規登録時だけ利用できます。"
                GoTo Finished
            End If
            If actionName = "ClipboardImport" Then
                If Not ReadClipboard(text) Then
                    message = "クリップボードを読めません。ImportPaste欄へCtrl+Vして『貼付解析』を実行してください。"
                    GoTo Finished
                End If
            Else
                text = CStr(UiRange("ImportPaste").Value2)
            End If
            Set values = ScreenValues(draftMasked, message)
            If values Is Nothing Then GoTo Finished
            If Not ParseImport(text, values, message) Then GoTo Finished
            draftMasked = values("MaskedNumber")
            DisplayValues values
            UiRange("ImportPaste").MergeArea.ClearContents
            text = ""
        Case "Save"
            If Not hasDetail Then GoTo Finished
            If Not isNewMode Then
                If CStr(UiRange("DetailYear").Value2) <> currentCase.sourceYear Then
                    message = "既存事案の事務年度は変更できません。"
                    GoTo Finished
                End If
                If CStr(FieldCell("MaskedNumber").Value2) <> "*************" Then
                    message = "番号表示欄は13個のアスタリスクに戻してください。"
                    GoTo Finished
                End If
                Set record = currentCase
                draftMasked = currentCase.values("MaskedNumber")
            Else
                Set record = New CaseRecord
                record.sourceYear = CStr(UiRange("DetailYear").Value2)
                If years.Exists(record.sourceYear) Then record.sourceAddress = years(record.sourceYear)(0)
            End If
            If Len(CStr(UiRange("NumberPaste").Value2)) > 0 Or Len(CStr(UiRange("ImportPaste").Value2)) > 0 Then
                message = "未解析の貼付欄があります。番号取込・貼付解析を実行するか、内容を控えて欄を消去してください。"
                GoTo Finished
            End If
            Set values = ScreenValues(draftMasked, message)
            If values Is Nothing Then GoTo Finished
            If Not ValidateValues(values, message) Then GoTo Finished
            If Not RestoreNumber(values("MaskedNumber"), number) Then GoTo Finished
            targetKey = record.sourceYear & "|" & number & values("Branch")
            If MsgBox(record.sourceYear & "年度の選択事案を保存します。氏名・枝番・年度・表示内容を確認してください。", vbOKCancel, "保存前確認") <> vbOK Then GoTo Finished
            If SaveCase(record, values, isNewMode, message) Then
                RefreshOneYear record.sourceYear, message
                If Cases.Exists(targetKey) Then
                    ShowCase targetKey
                Else
                    hasDetail = False
                End If
                BuildSearch
            End If
        Case "Confirm"
            If ReconcileSave(message) Then
                '照合確認後にだけ編集終了を許可。画面の現在入力は保持する。
                message = message & vbCrLf & "現在の入力を控え、最新を再読込して内容を確認してください。"
            End If
    End Select
Finished:
    isBusy = False
    If Len(message) > 0 Then Notify message
    Exit Sub
Failed:
    message = SafeFailure(actionName, Err.Number)
    Resume Finished
End Sub

Public Function UiCanLeaveDetail() As Boolean
    Dim id As Variant
    Dim dirty As Boolean
    If PendingSave() Then
        Notify "保存結果の確認が必要です。照合が終わるまで画面切替と再読込を止めています。"
        Exit Function
    End If
    If hasDetail And Not baseline Is Nothing Then
        If CStr(UiRange("DetailYear").Value2) <> baselineYear Then dirty = True
        For Each id In fields.Keys
            If Not SameValue(FieldCell(CStr(id)).Value2, baseline(id)) Then dirty = True
        Next
        If isNewMode And Len(draftMasked) > 0 Then dirty = True
        If Len(CStr(UiRange("NumberPaste").Value2)) > 0 Or Len(CStr(UiRange("ImportPaste").Value2)) > 0 Then dirty = True
    End If
    If dirty Then
        If MsgBox("未保存の入力があります。入力を破棄して切り替えますか。", vbYesNo + vbExclamation, "未保存の確認") <> vbYes Then Exit Function
    End If
    UiCanLeaveDetail = True
End Function

Private Sub RememberBaseline()
    Dim id As Variant
    Set baseline = NewDictionary()
    baselineYear = CStr(UiRange("DetailYear").Value2)
    For Each id In fields.Keys
        baseline.Add id, FieldCell(CStr(id)).Value2
    Next
End Sub

Private Sub BuildSearch()
    Dim message As String
    Set people = SearchPeople(CStr(UiRange("SearchName").Value2), CStr(UiRange("SearchYear").Value2), _
        CStr(UiRange("SearchField").Value2), CStr(UiRange("SearchValue").Value2), message)
    pageIndex = 0
    DrawSearchPage
    If Len(message) > 0 Then Notify message
End Sub

Private Sub DrawSearchPage()
    Dim index As Long
    Dim rowIndex As Long
    Dim keys As Variant
    Dim item As Variant
    Dim token As String
    Dim target As Range
    Set target = UiRange("SearchResults")
    target.ClearContents
    Set personTokens = NewDictionary()
    If people Is Nothing Then Exit Sub
    keys = people.Keys
    For index = pageIndex * 12 To Application.Min(people.Count - 1, pageIndex * 12 + 11)
        rowIndex = index - pageIndex * 12 + 1
        token = "P" & Format$(index + 1, "000000")
        personTokens.Add token, keys(index)
        item = people(keys(index))
        PutValue target.Cells(rowIndex, 1), token
        PutValue target.Cells(rowIndex, 2), item(0)
        PutValue target.Cells(rowIndex, 3), item(1)
        target.Cells(rowIndex, 4).Value2 = item(2)
        PutValue target.Cells(rowIndex, 5), item(3)
    Next
    If HasRelatedViewport() Then
        PutValue UiRange("SearchSummary"), "検索結果：" & people.Count & "人"
        PutValue UiRange("SearchPageInfo"), "ページ：" & IIf(people.Count = 0, 0, pageIndex + 1) & " / " & (people.Count + 11) \ 12
    Else
        PutValue UiRange("SearchState"), "該当 " & people.Count & "人 / ページ " & pageIndex + 1
    End If
End Sub

Private Function SelectedToken(ByVal area As Range) As String
    If Not ActiveWorkbook Is ThisWorkbook Then Exit Function
    If Not ActiveSheet Is area.Parent Then Exit Function
    If Intersect(ActiveCell, area) Is Nothing Then Exit Function
    SelectedToken = CStr(area.Parent.Cells(ActiveCell.Row, area.Column).Value2)
End Function

Private Sub ClearRelated()
    If HasRelatedViewport() Then
        RelatedArea().ClearContents
    ElseIf relatedCount > 0 Then
        UiRange("RelatedStart").Resize(relatedCount, 7).ClearContents
    End If
    relatedCount = 0
    relatedPageIndex = 0
    Set caseTokens = NewDictionary()
End Sub

Private Sub ShowRelated(ByVal groupKey As String)
    Dim key As Variant
    Dim record As CaseRecord
    Dim recordGroup As String
    Dim token As String
    Dim firstKey As String
    Dim target As Range
    ClearRelated
    For Each key In Cases.Keys
        Set record = Cases(key)
        recordGroup = record.PersonKey
        If Len(recordGroup) = 0 Then recordGroup = "Invalid|" & CStr(key)
        If recordGroup = groupKey Then
            relatedCount = relatedCount + 1
            token = "C" & Format$(relatedCount, "000000")
            caseTokens.Add token, key
            If Not HasRelatedViewport() Then
                Set target = UiRange("RelatedStart").Offset(relatedCount - 1, 0)
                PutValue target.Cells(1, 1), token
                PutValue target.Cells(1, 2), record.sourceYear
                PutValue target.Cells(1, 3), record.values("TaxYear")
                PutValue target.Cells(1, 4), record.values("TaxType")
                PutValue target.Cells(1, 5), record.values("DocumentName")
                PutValue target.Cells(1, 6), record.values("Status")
                PutValue target.Cells(1, 7), record.invalidReason
            End If
            If Len(firstKey) = 0 Then firstKey = CStr(key)
        End If
    Next
    If HasRelatedViewport() Then DrawRelatedPage
    If Len(firstKey) > 0 Then ShowCase firstKey
End Sub

Private Sub ShowCase(ByVal key As String)
    Dim token As Variant
    Dim selected As Range
    Set currentCase = Cases(key)
    hasDetail = True
    isNewMode = False
    draftMasked = currentCase.values("MaskedNumber")
    DisplayValues currentCase.values
    PutValue UiRange("DetailYear"), currentCase.sourceYear
    PutValue UiRange("DetailState"), currentCase.sourceYear & "年度 / 管理簿読込時の行 " & currentCase.sourceRow & " / " & currentCase.invalidReason
    UiRange("NumberPaste").MergeArea.ClearContents
    UiRange("ImportPaste").MergeArea.ClearContents
    If HasRelatedViewport() Then
        DrawRelatedPage
    ElseIf Not caseTokens Is Nothing And relatedCount > 0 Then
        UiRange("RelatedStart").Resize(relatedCount, 1).Interior.ColorIndex = xlColorIndexNone
        For Each token In caseTokens.Keys
            If caseTokens(token) = key Then
                Set selected = UiRange("RelatedStart").Resize(relatedCount, 1).Find(What:=CStr(token), LookIn:=xlValues, LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)
                If Not selected Is Nothing Then selected.Interior.Color = RGB(255, 230, 150)
            End If
        Next
    End If
    RememberBaseline
    ShowDetailScreen
End Sub

Public Sub ApplyUiLayout()
    Dim snapshots As Object
    Dim name As Variant
    Dim savedNames As Object
    Dim rowIndex As Long
    Dim fieldId As String
    Dim ws As Worksheet
    Dim definition As Variant
    Dim originalUpdating As Boolean
    Dim originalEvents As Boolean
    Dim backupPath As String
    Dim fileSystem As Object
    Dim message As String
    If Not BeginAction() Then Exit Sub
    originalUpdating = Application.ScreenUpdating
    originalEvents = Application.EnableEvents
    On Error GoTo Failed
    If ThisWorkbook.ReadOnly Or ThisWorkbook.ProtectStructure Or Len(ThisWorkbook.Path) = 0 Or ThisWorkbook.FileFormat <> 52 Then
        message = "保存可能なCONTROLLER.xlsmで実行してください。ブック構成の保護も確認してください。"
        GoTo Finished
    End If
    If PendingSave() Or Len(CStr(Sheet("MoveRecovery").Range("B1").Value2)) > 0 Then
        message = "保存・移動の照合待ちがあります。照合を終えてから画面改修を実行してください。"
        GoTo Finished
    End If
    Set snapshots = NewDictionary()
    Set savedNames = NewDictionary()
    For Each name In Array("検索", "詳細", "Calendar")
        If Sheet(CStr(name)).ProtectContents Then
            message = CStr(name) & "シートが保護されています。画面を変更せず中止しました。"
            GoTo Finished
        End If
    Next
    Set ws = Sheet("Fields")
    If ws.ProtectContents Then
        message = "Fieldsが保護されています。画面を変更せず中止しました。"
        GoTo Finished
    End If
    For rowIndex = 2 To LastRow(ws)
        fieldId = CStr(ws.Cells(rowIndex, 1).Value2)
        If Len(fieldId) > 0 Then
            If snapshots.Exists(fieldId) Then
                message = "Fieldsの項目IDが重複しています。画面を変更せず中止しました。"
                GoTo Finished
            End If
            Set ws = Sheet("Fields")
            definition = Array(CStr(ws.Cells(rowIndex, 4).Value2), CStr(ws.Cells(rowIndex, 5).Value2))
            If Sheet(definition(0)).Range(definition(1)).HasFormula Then
                message = "入力欄に数式があります。" & fieldId & "。画面を変更せず中止しました。"
                GoTo Finished
            End If
            snapshots.Add fieldId, Sheet(definition(0)).Range(definition(1)).Value2
        End If
    Next
    For Each name In Array("SearchName", "SearchYear", "SearchField", "SearchValue", "DetailYear", "NumberPaste", "ImportPaste")
        savedNames.Add name, UiRange(CStr(name)).Value2
    Next
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    backupPath = ThisWorkbook.Path & "\" & fileSystem.GetBaseName(ThisWorkbook.Name) & "_UiBackup_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsm"
    If fileSystem.FileExists(backupPath) Then
        message = "同名の画面変更前バックアップがあります。少し待って再実行してください。"
        GoTo Finished
    End If
    ThisWorkbook.SaveCopyAs backupPath
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    BuildSearchLayout
    BuildDetailLayout
    BuildImportLayout
    BuildCalendarLayout
    BuildSettingsLayout
    Set ws = Sheet("Fields")
    For rowIndex = 2 To LastRow(ws)
        fieldId = CStr(ws.Cells(rowIndex, 1).Value2)
        If Len(fieldId) > 0 Then
            PutValue Sheet(CStr(ws.Cells(rowIndex, 4).Value2)).Range(CStr(ws.Cells(rowIndex, 5).Value2)), snapshots(fieldId), (ws.Cells(rowIndex, 3).Value2 = "Date")
        End If
    Next
    For Each name In savedNames.Keys
        PutValue UiRange(CStr(name)), savedNames(name)
    Next
    Set fields = Nothing
    Set years = Nothing
    Set Cases = Nothing
    Set currentCase = Nothing
    Set baseline = Nothing
    Set people = Nothing
    Set personTokens = Nothing
    Set caseTokens = Nothing
    relatedCount = 0
    relatedPageIndex = 0
    hasDetail = False
    isNewMode = False
    ThisWorkbook.Names.Add Name:="UiLayoutVersion", RefersTo:="=""2"""
    PutValue UiRange("SearchState"), "画面を更新しました。『最新を再読込』で設定と管理簿を読み込んでください。"
    PutValue UiRange("DetailState"), "入力内容は保持しました。検索から事案を選択して編集を開始してください。"
    ThisWorkbook.Save
    ShowSearchScreen
    message = "画面を更新しました。設定と入力は保持しました。変更前のバックアップ：" & vbCrLf & backupPath & vbCrLf & "『最新を再読込』を実行してください。"
Finished:
    Application.EnableEvents = originalEvents
    Application.ScreenUpdating = originalUpdating
    isBusy = False
    If Len(message) > 0 Then Notify message
    Exit Sub
Failed:
    message = SafeFailure("画面レイアウトの更新", Err.Number) & vbCrLf & "変更前の状態はバックアップを確認してください。" & backupPath
    Resume Finished
End Sub

Private Sub PrepareUiSheet(ByVal ws As Worksheet, ByVal lastUiRow As Long)
    Dim index As Long
    Dim shape As Shape
    Dim widths As Variant
    ws.Range("A1:P" & CStr(lastUiRow)).UnMerge
    ws.Range("A1:P" & CStr(lastUiRow)).Clear
    For index = ws.Shapes.Count To 1 Step -1
        Set shape = ws.Shapes(index)
        If Left$(shape.Name, 4) = "Ctl_" Or shape.Type = msoFormControl Then shape.Delete
    Next
    With ws.Range("A1:P" & CStr(lastUiRow))
        .Font.Name = "Meiryo UI"
        .Font.Size = 10.5
        .Font.Color = RGB(40, 56, 73)
        .Interior.Color = RGB(244, 247, 251)
        .VerticalAlignment = xlCenter
    End With
    ws.Columns("B:P").Hidden = False
    ws.Columns("A").ColumnWidth = 2
    widths = Array(72, 112, 154, 74, 92, 105, 54, 180)
    For index = 0 To UBound(widths)
        SetColumnPoints ws.Columns(index + 2), CDbl(widths(index))
    Next
    ws.Rows("1:" & lastUiRow).RowHeight = 25
    ws.Rows(1).RowHeight = 38
    ws.Rows(4).RowHeight = 8
    ws.Rows(1).Interior.Color = RGB(244, 247, 251)
    ws.Range("B1:I1").Interior.Color = RGB(27, 56, 83)
    ws.DisplayPageBreaks = False
    ws.ScrollArea = ""
End Sub

Private Sub SetColumnPoints(ByVal column As Range, ByVal points As Double)
    Dim attempt As Long
    For attempt = 1 To 3
        column.ColumnWidth = column.ColumnWidth * points / column.Width
    Next
End Sub

Private Sub UiText(ByVal area As Range, ByVal text As String)
    If area.CountLarge > 1 Then area.Merge
    area.Cells(1, 1).Value2 = text
    area.VerticalAlignment = xlCenter
    area.WrapText = True
End Sub

Private Sub UiHeader(ByVal ws As Worksheet, ByVal caption As String)
    Dim header As Shape
    UiText ws.Range("B1:I1"), ""
    Set header = ws.Shapes.AddShape(msoShapeRectangle, ws.Range("B1").Left, ws.Range("B1").Top, ws.Range("B1:I1").Width, 38)
    header.Name = "Ctl_Header"
    header.Line.Visible = msoFalse
    header.Fill.Solid
    header.Fill.ForeColor.RGB = RGB(27, 56, 83)
    With header.TextFrame2
        .TextRange.Text = "CONTROLLER   /   " & caption
        .TextRange.Font.Name = "Meiryo UI"
        .TextRange.Font.Size = 19
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .MarginLeft = 14
        .VerticalAnchor = msoAnchorMiddle
    End With
    header.Placement = xlFreeFloating
End Sub

Private Sub UiSection(ByVal area As Range, ByVal caption As String)
    UiText area, caption
    area.Interior.Color = RGB(218, 231, 242)
    area.Font.Color = RGB(27, 56, 83)
    area.Font.Bold = True
    area.Cells(1, 1).IndentLevel = 1
End Sub

Private Sub UiInput(ByVal area As Range, Optional ByVal isDate As Boolean = False)
    UiText area, ""
    area.Interior.Color = RGB(255, 255, 255)
    area.Font.Color = RGB(35, 51, 68)
    area.Font.Size = 11
    area.Borders.LineStyle = xlContinuous
    area.Borders.Color = RGB(215, 224, 235)
    area.Borders(xlEdgeBottom).Color = RGB(66, 133, 180)
    area.Cells(1, 1).IndentLevel = 1
    area.NumberFormat = "@"
    If isDate Then area.NumberFormat = "yyyy/mm/dd"
End Sub

Private Sub UiShapeButton(ByVal ws As Worksheet, ByVal key As String, ByVal caption As String, ByVal macro As String, _
    ByVal left As Double, ByVal top As Double, ByVal width As Double, Optional ByVal primary As Boolean = False)
    Dim button As Shape
    Dim index As Long
    For index = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(index).Name = "Ctl_" & key Then ws.Shapes(index).Delete
    Next
    Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, width, 24)
    button.Name = "Ctl_" & key
    button.Adjustments(1) = 0.12
    button.Line.ForeColor.RGB = RGB(204, 216, 230)
    button.Fill.ForeColor.RGB = RGB(255, 255, 255)
    If primary Then
        button.Fill.ForeColor.RGB = RGB(34, 111, 166)
        button.Line.ForeColor.RGB = RGB(34, 111, 166)
    End If
    With button.TextFrame2
        .TextRange.Text = caption
        .TextRange.Font.Name = "Meiryo UI"
        .TextRange.Font.Size = 10
        .TextRange.Font.Fill.ForeColor.RGB = RGB(42, 59, 77)
        If primary Then
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.Font.Bold = msoTrue
        End If
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .MarginLeft = 4
        .MarginRight = 4
    End With
    button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macro
    button.Placement = xlFreeFloating
End Sub

Private Sub BuildSearchLayout()
    Dim ws As Worksheet
    Dim rowIndex As Long
    Set ws = Sheet("検索")
    PrepareUiSheet ws, 60
    SetColumnPoints ws.Columns("B"), 65
    SetColumnPoints ws.Columns("C"), 265
    SetColumnPoints ws.Columns("D"), 85
    SetColumnPoints ws.Columns("E"), 130
    SetColumnPoints ws.Columns("F"), 298
    ws.Columns("G:I").Hidden = True
    '検索画面はB:Fの5列全体で同じ幅にする。
    UiText ws.Range("B1:F1"), ""
    UiHeader ws, "人物検索"
    ws.Shapes("Ctl_Header").Width = ws.Range("B1:F1").Width
    ws.Range("B3").Value2 = "氏名"
    ws.Range("D3").Value2 = "事務年度"
    UiInput ws.Range("C3")
    UiInput ws.Range("E3")
    AddUiName "SearchName", ws.Range("C3")
    AddUiName "SearchYear", ws.Range("E3")
    UiShapeButton ws, "Find", "検索", "FindPeople", ws.Range("F3").Left + 8, ws.Range("F3").Top, 72, True
    UiShapeButton ws, "Reload", "最新を再読込", "StartController", ws.Range("F3").Left + 88, ws.Range("F3").Top, 110
    UiShapeButton ws, "New", "新規登録", "NewCase", ws.Range("F3").Left + 206, ws.Range("F3").Top, 82
    ws.Range("B5").Value2 = "追加項目"
    ws.Range("D5").Value2 = "検索値"
    UiInput ws.Range("C5")
    UiInput ws.Range("E5")
    AddUiName "SearchField", ws.Range("C5")
    AddUiName "SearchValue", ws.Range("E5")
    ws.Range("F5").Value2 = "年度が空欄なら全年度。追加項目は省略できます。"
    ws.Range("F5").Font.Size = 9
    UiText ws.Range("B6:F6"), "氏名の一部を入力して検索してください。"
    AddUiName "SearchState", ws.Range("B6")
    ws.Rows(6).RowHeight = 32
    UiText ws.Range("B7:C7"), "検索結果：0人"
    AddUiName "SearchSummary", ws.Range("B7")
    ws.Range("B7").Font.Size = 15
    ws.Range("B7").Font.Bold = True
    ws.Range("B7").Font.Color = RGB(34, 111, 166)
    UiText ws.Range("D7:E7"), "ページ：0 / 0"
    AddUiName "SearchPageInfo", ws.Range("D7")
    UiShapeButton ws, "Previous", "前へ", "PreviousPage", ws.Range("F7").Left + ws.Range("F7").Width - 160, ws.Range("F7").Top, 76
    UiShapeButton ws, "Next", "次へ", "NextPage", ws.Range("F7").Left + ws.Range("F7").Width - 76, ws.Range("F7").Top, 76
    ws.Range("B9:F9").Value2 = Array("選択ID", "氏名", "枝番", "関連事案件数", "確認状態")
    StyleTableHeader ws.Range("B9:F9")
    For rowIndex = 10 To 21
        StyleTableRow ws.Range("B" & rowIndex & ":F" & rowIndex), rowIndex Mod 2 = 0
        ws.Rows(rowIndex).RowHeight = 29
    Next
    AddUiName "SearchResults", ws.Range("B10:F21")
    ws.Range("B23:E23").Merge
    ws.Range("B23").Value2 = "一覧の行を選択し、右のボタンで関連事案を開きます。"
    UiShapeButton ws, "Open", "選択人物を開く", "OpenPerson", ws.Range("F23").Left + ws.Range("F23").Width - 166, ws.Range("F23").Top, 166, True
    UiSection ws.Range("B25:F25"), "業務管理  /  常勤用の管理操作"
    UiShapeButton ws, "Start", "業務開始", "StartBusiness", ws.Range("B26").Left, ws.Range("B26").Top, 110
    UiShapeButton ws, "Finish", "業務終了", "FinishBusiness", ws.Range("B26").Left + 120, ws.Range("B26").Top, 110
    UiShapeButton ws, "Move", "移動結果を確認", "ConfirmMoveResult", ws.Range("B26").Left + 240, ws.Range("B26").Top, 135
    UiShapeButton ws, "Restore", "バックアップ復元", "RestoreBackup", ws.Range("F26").Left + ws.Range("F26").Width - 160, ws.Range("F26").Top, 160
    ws.Shapes("Ctl_Start").Visible = (Setting("ControllerRole") = "Staff")
    ws.Shapes("Ctl_Finish").Visible = (Setting("ControllerRole") = "Staff")
    ws.Shapes("Ctl_Move").Visible = (Setting("ControllerRole") = "Staff")
    ws.Shapes("Ctl_Restore").Visible = (Setting("ControllerRole") = "Staff")
    UiShapeButton ws, "Settings", "設定を開く", "ShowSettingsScreen", ws.Range("B28").Left, ws.Range("B28").Top, 110
    UiShapeButton ws, "Detail", "詳細画面へ", "ShowDetailScreen", ws.Range("F28").Left + ws.Range("F28").Width - 160, ws.Range("F28").Top, 160
End Sub

Private Sub BuildDetailLayout()
    Dim ws As Worksheet
    Dim definitionSheet As Worksheet
    Dim rowIndex As Long
    Dim fieldId As String
    Dim address As String
    Dim extraIndex As Long
    Dim label As String
    Set ws = Sheet("詳細")
    PrepareUiSheet ws, Application.Max(70, LastRow(Sheet("Fields")) + 40)
    UiHeader ws, "事案の詳細"
    UiText ws.Range("B2:I2"), "検索から人物・事案を選択してください。"
    AddUiName "DetailState", ws.Range("B2")
    ws.Rows(2).RowHeight = 28
    UiShapeButton ws, "Back", "検索へ戻る", "ShowSearchScreen", ws.Range("B3").Left, ws.Range("B3").Top, 110
    UiShapeButton ws, "New", "新規登録", "NewCase", ws.Range("B3").Left + 120, ws.Range("B3").Top, 100
    UiShapeButton ws, "Import", "番号・外部入力", "ShowImportScreen", ws.Range("B3").Left + 230, ws.Range("B3").Top, 140
    UiShapeButton ws, "Confirm", "保存結果を確認", "ConfirmSaveResult", ws.Range("I3").Left - 136, ws.Range("I3").Top, 130
    UiShapeButton ws, "Save", "更新・保存", "SaveDetail", ws.Range("I3").Left + 8, ws.Range("I3").Top, 172, True
    ws.Range("B5").Value2 = "事務年度"
    UiInput ws.Range("C5:D5")
    AddUiName "DetailYear", ws.Range("C5")
    UiText ws.Range("F5:I5"), "白い欄を編集します。既存事案の年度・番号・枝番は変更できません。"
    UiText ws.Range("B7:F7"), "同一人物の関連事案"
    AddUiName "RelatedState", ws.Range("B7")
    UiShapeButton ws, "PreviousCase", "前へ", "PreviousRelatedPage", ws.Range("I7").Left + 12, ws.Range("I7").Top, 76
    UiShapeButton ws, "NextCase", "次へ", "NextRelatedPage", ws.Range("I7").Left + 96, ws.Range("I7").Top, 76
    ws.Range("B8:F8").Value2 = Array("選択ID", "事務年度", "文書名", "年分", "税目")
    UiText ws.Range("G8:H8"), "状態"
    ws.Range("I8").Value2 = "確認事項"
    StyleTableHeader ws.Range("B8:I8")
    For rowIndex = 9 To 13
        ws.Range("G" & rowIndex & ":H" & rowIndex).Merge
        StyleTableRow ws.Range("B" & rowIndex & ":I" & rowIndex), rowIndex Mod 2 = 0
        ws.Rows(rowIndex).RowHeight = 26
    Next
    AddUiName "RelatedStart", ws.Range("B9")
    AddUiName "RelatedViewport", ws.Range("B9:I13")
    UiText ws.Range("B14:F14"), "行を選択して『選択事案を開く』。選択中の事案は青く表示します。"
    ws.Range("B14").Font.Size = 9
    UiShapeButton ws, "OpenCase", "選択事案を開く", "OpenRelatedCase", ws.Range("I14").Left + 8, ws.Range("I14").Top, 172, True
    ws.Rows(15).RowHeight = 7
    UiSection ws.Range("B16:I16"), "基本情報"
    UiSection ws.Range("B21:I21"), "事案情報・処理状況"
    UiSection ws.Range("B27:I27"), "メモ"
    Set definitionSheet = Sheet("Fields")
    For rowIndex = 2 To LastRow(definitionSheet)
        fieldId = CStr(definitionSheet.Cells(rowIndex, 1).Value2)
        If Len(fieldId) > 0 Then
            address = DefaultDetailAddress(fieldId)
            If Len(address) = 0 Then
                address = IIf(extraIndex Mod 2 = 0, "C", "G") & CStr(35 + extraIndex \ 2)
                extraIndex = extraIndex + 1
            End If
            PutValue definitionSheet.Cells(rowIndex, 4), "詳細"
            PutValue definitionSheet.Cells(rowIndex, 5), address
            label = CStr(definitionSheet.Cells(rowIndex, 2).Value2)
            If fieldId = "MaskedNumber" Then label = "識別番号"
            ws.Range(address).Offset(0, -1).Value2 = label
            If fieldId = "Memo1" Or fieldId = "Memo2" Or fieldId = "Memo3" Then
                UiInput ws.Range(address).Resize(1, 7)
                ws.Range(address).EntireRow.RowHeight = 31
            Else
                UiInput ws.Range(address).Resize(1, 3), (definitionSheet.Cells(rowIndex, 3).Value2 = "Date")
            End If
            If fieldId = "MaskedNumber" Then ws.Range(address).MergeArea.Interior.Color = RGB(232, 238, 245)
        End If
    Next
    UiShapeButton ws, "CopyName", "氏名コピー", "CopyPersonName", ws.Range("B32").Left, ws.Range("B32").Top, 120
    UiShapeButton ws, "CopyAddress", "住所コピー", "CopyAddress", ws.Range("B32").Left + 130, ws.Range("B32").Top, 120
    UiShapeButton ws, "CopyKana", "フリガナコピー", "CopyFurigana", ws.Range("B32").Left + 260, ws.Range("B32").Top, 130
    UiShapeButton ws, "CopyNumber", "18桁番号コピー", "CopyNumber", ws.Range("F32").Left + 5, ws.Range("F32").Top, 150
    UiShapeButton ws, "Calendar", "日付カレンダー", "OpenCalendar", ws.Range("I32").Left + 8, ws.Range("I32").Top, 172
    If extraIndex > 0 Then UiSection ws.Range("B34:I34"), "追加項目"
End Sub

Private Function DefaultDetailAddress(ByVal fieldId As String) As String
    Select Case fieldId
        Case "PersonName"
            DefaultDetailAddress = "C17"
        Case "Furigana"
            DefaultDetailAddress = "G17"
        Case "Address"
            DefaultDetailAddress = "C18"
        Case "Office"
            DefaultDetailAddress = "G18"
        Case "PostalCode"
            DefaultDetailAddress = "C19"
        Case "PhoneNumber"
            DefaultDetailAddress = "G19"
        Case "MaskedNumber"
            DefaultDetailAddress = "C20"
        Case "Branch"
            DefaultDetailAddress = "G20"
        Case "DocumentName"
            DefaultDetailAddress = "C22"
        Case "TaxYear"
            DefaultDetailAddress = "G22"
        Case "TaxType"
            DefaultDetailAddress = "C23"
        Case "Status"
            DefaultDetailAddress = "G23"
        Case "ReceivedDate"
            DefaultDetailAddress = "C24"
        Case "DraftDate"
            DefaultDetailAddress = "G24"
        Case "ApprovedDate"
            DefaultDetailAddress = "C25"
        Case "IssuedDate"
            DefaultDetailAddress = "G25"
        Case "ApprovalNumber"
            DefaultDetailAddress = "C26"
        Case "InquiryNumber"
            DefaultDetailAddress = "G26"
        Case "Memo1"
            DefaultDetailAddress = "C28"
        Case "Memo2"
            DefaultDetailAddress = "C29"
        Case "Memo3"
            DefaultDetailAddress = "C30"
    End Select
End Function

Private Sub StyleTableHeader(ByVal area As Range)
    area.Interior.Color = RGB(37, 72, 102)
    area.Font.Color = RGB(255, 255, 255)
    area.Font.Bold = True
    area.HorizontalAlignment = xlCenter
    area.WrapText = True
End Sub

Private Sub StyleTableRow(ByVal area As Range, ByVal tinted As Boolean)
    area.Interior.Color = RGB(255, 255, 255)
    If tinted Then area.Interior.Color = RGB(235, 241, 247)
    area.Borders.LineStyle = xlContinuous
    area.Borders.Color = RGB(217, 226, 236)
    area.WrapText = True
    area.NumberFormat = "@"
End Sub

Private Sub BuildImportLayout()
    Dim ws As Worksheet
    Dim candidate As Worksheet
    For Each candidate In ThisWorkbook.Worksheets
        If candidate.Name = "取込" Then Set ws = candidate
    Next
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "取込"
    End If
    PrepareUiSheet ws, 30
    UiHeader ws, "番号・外部テキストの取込"
    UiText ws.Range("B3:I3"), "新規登録で使用します。取込後は詳細画面で内容を確認し、保存してください。"
    UiSection ws.Range("B5:I5"), "1   18桁番号を貼り付ける"
    UiInput ws.Range("C7:F7")
    AddUiName "NumberPaste", ws.Range("C7")
    UiShapeButton ws, "Number", "番号を取り込む", "ApplyNumberPaste", ws.Range("I7").Left + 8, ws.Range("I7").Top, 172, True
    UiText ws.Range("B9:I9"), "半角数字18桁を文字列として貼付。取込成功後、貼付欄は消去します。"
    UiSection ws.Range("B11:I11"), "2   外部システムの文字列を取り込む"
    UiInput ws.Range("C13:I18")
    AddUiName "ImportPaste", ws.Range("C13")
    ws.Range("C13").VerticalAlignment = xlTop
    UiShapeButton ws, "Clipboard", "クリップボードから取込", "ImportClipboard", ws.Range("C20").Left, ws.Range("C20").Top, 200
    UiShapeButton ws, "Parse", "貼付内容を解析", "ParsePastedText", ws.Range("C20").Left + 212, ws.Range("C20").Top, 170, True
    UiText ws.Range("B22:I22"), "取込だけでは正式管理簿を保存しません。見つからなかった項目の入力は保持します。"
    UiShapeButton ws, "Back", "詳細画面で内容を確認", "ShowDetailScreen", ws.Range("I24").Left - 40, ws.Range("I24").Top, 220, True
End Sub

Private Sub BuildCalendarLayout()
    Dim ws As Worksheet
    Set ws = Sheet("Calendar")
    PrepareUiSheet ws, 20
    UiHeader ws, "日付を選択"
    ws.Range("B4:H4").Value2 = Array("日", "月", "火", "水", "木", "金", "土")
    StyleTableHeader ws.Range("B4:H4")
    ws.Range("B5:H10").Interior.Color = RGB(255, 255, 255)
    ws.Range("B5:H10").HorizontalAlignment = xlCenter
    ws.Range("B5:H10").Font.Size = 14
    ws.Range("B5:H10").Borders.Color = RGB(215, 224, 235)
    ws.Rows("5:10").RowHeight = 38
    UiShapeButton ws, "Previous", "前月", "PreviousMonth", ws.Range("B12").Left, ws.Range("B12").Top, 110
    UiShapeButton ws, "Next", "翌月", "NextMonth", ws.Range("B12").Left + 122, ws.Range("B12").Top, 110
    UiShapeButton ws, "Pick", "選択した日付を入力", "PickCalendarDate", ws.Range("F12").Left, ws.Range("F12").Top, 170, True
    UiShapeButton ws, "Clear", "日付を空欄にする", "ClearCalendarDate", ws.Range("F14").Left, ws.Range("F14").Top, 170
    UiShapeButton ws, "Back", "詳細へ戻る", "ShowDetailScreen", ws.Range("B14").Left, ws.Range("B14").Top, 110
End Sub

Private Sub BuildSettingsLayout()
    Dim ws As Worksheet
    Dim sheetName As Variant
    For Each sheetName In Array("Settings", "Fields")
        Set ws = Sheet(CStr(sheetName))
        ws.Cells.Font.Name = "Meiryo UI"
        ws.Cells.Font.Size = 10
        ws.Rows(1).RowHeight = 30
        StyleTableHeader ws.Range("A1:L1")
        ws.Columns("B:P").Hidden = False
    ws.Columns("A").ColumnWidth = 23
        ws.Columns("B").ColumnWidth = 42
        ws.Columns("C").ColumnWidth = 10
        ws.Columns("D").ColumnWidth = 14
        ws.Columns("E").ColumnWidth = 45
        ws.Columns("F:H").ColumnWidth = 15
        ws.Columns("I:J").ColumnWidth = 45
        ws.Columns("K:L").ColumnWidth = 16
        UiShapeButton ws, "Back", "検索画面へ戻る", "ShowSearchScreen", ws.Range("N2").Left, ws.Range("N2").Top, 160
    Next
End Sub

Public Sub ShowSearchScreen()
    ActivateUiSheet "検索", "B3"
End Sub

Public Sub ShowDetailScreen()
    ActivateUiSheet "詳細", "B5"
End Sub

Public Sub ShowImportScreen()
    ActivateUiSheet "取込", "C7"
End Sub

Public Sub ShowSettingsScreen()
    Sheet("Settings").Activate
End Sub

Private Sub ActivateUiSheet(ByVal sheetName As String, ByVal address As String)
    Sheet(sheetName).Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    If sheetName = "詳細" Then
        ActiveWindow.Zoom = Application.Min(100, Application.Max(65, Fix((Application.UsableHeight - 150) / Sheet(sheetName).Range("B1:I32").Height * 100)))
    ElseIf sheetName = "検索" Then
        ActiveWindow.Zoom = Application.Min(100, Application.Max(65, Fix((Application.UsableHeight - 150) / Sheet(sheetName).Range("B1:F28").Height * 100)))
    Else
        ActiveWindow.Zoom = 100
    End If
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Sheet(sheetName).Range(address).Select
End Sub

Private Function HasRelatedViewport() As Boolean
    Dim itemName As Name
    For Each itemName In ThisWorkbook.Names
        If itemName.Name = "RelatedViewport" Then
            HasRelatedViewport = True
            Exit Function
        End If
    Next
End Function

Private Function RelatedArea() As Range
    If HasRelatedViewport() Then
        Set RelatedArea = UiRange("RelatedViewport")
    Else
        Set RelatedArea = UiRange("RelatedStart").Resize(Application.Max(1, relatedCount), 7)
    End If
End Function

Public Sub PreviousRelatedPage()
    If isBusy Then Exit Sub
    If relatedPageIndex > 0 Then relatedPageIndex = relatedPageIndex - 1
    DrawRelatedPage
End Sub

Public Sub NextRelatedPage()
    If isBusy Then Exit Sub
    If (relatedPageIndex + 1) * 5 < relatedCount Then relatedPageIndex = relatedPageIndex + 1
    DrawRelatedPage
End Sub

Private Sub DrawRelatedPage()
    Dim area As Range
    Dim index As Long
    Dim rowIndex As Long
    Dim token As String
    Dim record As CaseRecord
    If Not HasRelatedViewport() Then Exit Sub
    Set area = RelatedArea()
    area.ClearContents
    For rowIndex = 1 To 5
        StyleTableRow area.Rows(rowIndex), rowIndex Mod 2 = 1
    Next
    If Not caseTokens Is Nothing Then
        For index = relatedPageIndex * 5 + 1 To Application.Min(relatedCount, relatedPageIndex * 5 + 5)
            token = "C" & Format$(index, "000000")
            Set record = Cases(caseTokens(token))
            rowIndex = index - relatedPageIndex * 5
            PutValue area.Cells(rowIndex, 1), token
            PutValue area.Cells(rowIndex, 2), record.sourceYear
            PutValue area.Cells(rowIndex, 3), record.values("DocumentName")
            PutValue area.Cells(rowIndex, 4), record.values("TaxYear")
            PutValue area.Cells(rowIndex, 5), record.values("TaxType")
            PutValue area.Cells(rowIndex, 6), record.values("Status")
            PutValue area.Cells(rowIndex, 8), record.invalidReason
            If Not currentCase Is Nothing Then
                If currentCase Is record Then area.Rows(rowIndex).Interior.Color = RGB(211, 231, 247)
            End If
        Next
    End If
    PutValue UiRange("RelatedState"), "同一人物の関連事案：" & relatedCount & "件  /  ページ " & IIf(relatedCount = 0, 0, relatedPageIndex + 1) & " / " & (relatedCount + 4) \ 5
End Sub
