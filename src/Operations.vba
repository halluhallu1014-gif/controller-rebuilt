Option Explicit

Private Function ElapsedSeconds(ByVal started As Double) As Double
    ElapsedSeconds = Timer - started
    If ElapsedSeconds < 0 Then ElapsedSeconds = ElapsedSeconds + 86400#
End Function

Private Function FilesEqual(ByVal first As String, ByVal second As String) As Boolean
    Dim stream As Object
    Dim firstBytes As Variant
    Dim secondBytes As Variant
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.LoadFromFile first
    firstBytes = stream.Read
    stream.Close
    stream.Open
    stream.LoadFromFile second
    secondBytes = stream.Read
    stream.Close
    FilesEqual = BinaryEqual(firstBytes, secondBytes)
End Function

Private Function BinaryEqual(ByVal first As Variant, ByVal second As Variant) As Boolean
    Dim index As Long
    If UBound(first) <> UBound(second) Then Exit Function
    For index = LBound(first) To UBound(first)
        If first(index) <> second(index) Then Exit Function
    Next
    BinaryEqual = True
End Function

Private Function CreateBackup(ByVal book As Workbook, ByVal year As String, ByRef message As String) As String
    Dim folder As String
    Dim expanded As String
    Dim zipFile As String
    Dim destination As String
    Dim fileSystem As Object
    Dim file As Object
    Dim ws As Worksheet
    Dim manifest As String
    Dim index As Long
    Dim restored As Workbook
    Dim state As Object
    Dim separator As String
    On Error GoTo Failed
    If Not IsDigits(year, Len(year)) Or Len(year) = 0 Or Len(year) > 8 Then
        message = "バックアップ対象年度は1～8桁の数字で設定してください。"
        Exit Function
    End If
    If Len(Setting("BackupFolder")) = 0 Then
        message = "BackupFolderが未設定です。バックアップを作成できないため移動は実行しません。"
        Exit Function
    End If
    folder = NewTempFolder()
    expanded = NewTempFolder()
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    For Each ws In book.Worksheets
        index = index + 1
        WriteUtf8 folder & "\sheet" & index & ".csv", SheetCsv(ws)
        manifest = manifest & CsvCell("sheet" & index & ".csv") & vbCrLf
    Next
    WriteUtf8 folder & "\fields.csv", SheetCsv(Sheet("Fields"))
    WriteUtf8 folder & "\manifest.csv", manifest
    WriteUtf8 folder & "\format.txt", "ControllerBackup/3 UTF-8-BOM CRLF CSV typed cells; Date1904=" & CStr(book.Date1904) & "; Year=" & year
    zipFile = NewEvidencePath("Backup.zip")
    CreateZip folder, zipFile
    ExpandZip zipFile, expanded
    For Each file In fileSystem.GetFolder(folder).Files
        If Not fileSystem.FileExists(expanded & "\" & file.Name) Then GoTo Failed
        If Not FilesEqual(file.Path, expanded & "\" & file.Name) Then GoTo Failed
    Next
    Set restored = RestoreFolder(expanded)
    If Not CompareBooks(book, restored) Then GoTo Failed
    restored.Close SaveChanges:=False
    Set restored = Nothing
    separator = "\"
    If IsWebAddress(Setting("BackupFolder")) Then separator = "/"
    destination = Setting("BackupFolder")
    If Right$(destination, 1) = separator Then destination = Left$(destination, Len(destination) - 1)
    destination = destination & separator & "CaseBackup_" & year & "_" & Format$(Date, "yyyymmdd") & "_" & Format$(Now, "hhnnss") & "_" & CStr(CLng(Timer * 10)) & ".zip"
    Set state = FileState(destination, message)
    If state Is Nothing Then Exit Function
    If state("Exists") Then
        message = "同名バックアップが存在します。上書きせず移動を中止しました。"
        Exit Function
    End If
    If Not UploadBackup(zipFile, destination, message) Then Exit Function
    CreateBackup = destination
    '一時ファイルは照合後も残す。必要なら管理担当がTemp内の該当成果を整理する。
    Exit Function
Failed:
    On Error Resume Next
    If Not restored Is Nothing Then restored.Close SaveChanges:=False
    On Error GoTo 0
    message = "バックアップの作成・ZIP展開・復元照合に失敗しました。管理簿は移動しません。WindowsのZIP機能、保存先、セルエラーを確認してください。"
End Function

Private Function UploadBackup(ByVal localFile As String, ByVal destination As String, ByRef message As String) As Boolean
    Dim fileSystem As Object
    Dim stream As Object
    Dim bytes As Variant
    Dim request As Object
    Dim digestValue As String
    Dim folderUrl As String
    Dim relative As String
    Dim fileName As String
    Dim endpoint As String
    On Error GoTo Failed
    If IsWebAddress(destination) Then
        If Len(ApiFile(destination)) = 0 Then
            message = "バックアップ先URLが不正です。同一SharePointサイトの直接フォルダURLを設定してください。"
            Exit Function
        End If
        digestValue = Digest(message)
        If Len(digestValue) = 0 Then Exit Function
        folderUrl = Left$(destination, InStrRev(destination, "/") - 1)
        fileName = Mid$(destination, InStrRev(destination, "/") + 1)
        relative = Mid$(folderUrl, InStr(9, folderUrl, "/"))
        endpoint = Setting("SharePointSite") & "/_api/web/GetFolderByServerRelativeUrl('" & _
            Application.WorksheetFunction.EncodeURL(Replace(relative, "'", "''")) & "')/Files/add(url='" & _
            Application.WorksheetFunction.EncodeURL(Replace(fileName, "'", "''")) & "',overwrite=false)"
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1
        stream.Open
        stream.LoadFromFile localFile
        bytes = stream.Read
        stream.Close
        Set request = HttpRequest("POST", endpoint, bytes, message, digestValue)
        If request Is Nothing Then Exit Function
        If request.status <> 200 And request.status <> 201 Then GoTo Failed
        Set request = HttpRequest("GET", ApiFile(destination) & "/$value", Empty, message)
        If request Is Nothing Then Exit Function
        If request.status <> 200 Then GoTo Failed
        If Not BinaryEqual(bytes, request.responseBody) Then GoTo Failed
    Else
        Set fileSystem = CreateObject("Scripting.FileSystemObject")
        fileSystem.CopyFile localFile, destination, False
        If Not FilesEqual(localFile, destination) Then GoTo Failed
    End If
    UploadBackup = True
    Exit Function
Failed:
    message = "バックアップ保存結果を確認できません。管理簿は移動しません。バックアップ先の同名ファイルを確認してください。"
End Function

Private Function PruneBackups(ByVal year As String, ByRef message As String) As Boolean
    Dim names As New Collection
    Dim dates As Object
    Dim days As Variant
    Dim name As Variant
    Dim folder As String
    Dim prefix As String
    Dim fileSystem As Object
    Dim file As Object
    Dim request As Object
    Dim node As Object
    Dim dateKey As String
    Dim index As Long
    Dim secondIndex As Long
    Dim swap As String
    Dim relative As String
    Dim state As Object
    Dim digestValue As String
    On Error GoTo Failed
    folder = Setting("BackupFolder")
    prefix = "CaseBackup_" & year & "_"
    Set dates = NewDictionary()
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If IsWebAddress(folder) Then
        relative = Mid$(folder, InStr(9, folder, "/"))
        Set request = HttpRequest("GET", Setting("SharePointSite") & "/_api/web/GetFolderByServerRelativeUrl('" & _
            Application.WorksheetFunction.EncodeURL(Replace(relative, "'", "''")) & "')/Files?$select=Name&$top=5000", Empty, message)
        If request Is Nothing Then Exit Function
        If request.status <> 200 Then GoTo Failed
        Set node = request.responseXML.SelectSingleNode("//*[local-name()='link' and @rel='next']")
        If Not node Is Nothing Then GoTo Failed
        For Each node In request.responseXML.SelectNodes("//*[local-name()='Name']")
            names.Add node.Text
        Next
    Else
        For Each file In fileSystem.GetFolder(folder).Files
            names.Add file.Name
        Next
    End If
    For Each name In names
        If ManagedBackupName(CStr(name), prefix) Then
            dateKey = Mid$(CStr(name), Len(prefix) + 1, 8)
            dates(dateKey) = True
        End If
    Next
    If dates.Count <= 5 Then
        PruneBackups = True
        Exit Function
    End If
    days = dates.Keys
    For index = 0 To UBound(days) - 1
        For secondIndex = index + 1 To UBound(days)
            If days(index) < days(secondIndex) Then
                swap = days(index)
                days(index) = days(secondIndex)
                days(secondIndex) = swap
            End If
        Next
    Next
    If IsWebAddress(folder) Then
        digestValue = Digest(message)
        If Len(digestValue) = 0 Then Exit Function
    End If
    For Each name In names
        If ManagedBackupName(CStr(name), prefix) Then
            dateKey = Mid$(CStr(name), Len(prefix) + 1, 8)
            If dateKey < days(4) Then
                If IsWebAddress(folder) Then
                    Set state = FileState(folder & "/" & name, message)
                    If state Is Nothing Then Exit Function
                    Set request = HttpRequest("DELETE", ApiFile(folder & "/" & name), Empty, message, digestValue, state("ETag"))
                    If request Is Nothing Then Exit Function
                    If request.status <> 200 And request.status <> 204 Then GoTo Failed
                Else
                    fileSystem.DeleteFile fileSystem.BuildPath(folder, name), False
                End If
            End If
        End If
    Next
    PruneBackups = True
    Exit Function
Failed:
    message = "新しいバックアップと移動は照合済みですが、古いZIPの整理に失敗しました。残ったZIPを保持してください。直近5作成日を超えた自システム形式だけが整理対象です。"
End Function

Public Sub RestoreBackup()
    Dim dialog As Object
    Dim folder As String
    Dim book As Workbook
    Dim message As String
    Dim events As Boolean
    If Not BeginAction() Then Exit Sub
    events = Application.EnableEvents
    On Error GoTo Failed
    If Setting("ControllerRole") <> "Staff" Then
        message = "バックアップ復元は常勤（Staff）だけが実行できます。"
        GoTo Finished
    End If
    Set dialog = Application.FileDialog(3)
    dialog.Title = "バックアップZIPを別の確認用ブックへ復元"
    dialog.Filters.Clear
    dialog.Filters.Add "CONTROLLERバックアップ", "*.zip"
    dialog.AllowMultiSelect = False
    If dialog.Show <> -1 Then GoTo Finished
    folder = NewTempFolder()
    ExpandZip dialog.SelectedItems(1), folder
    Application.EnableEvents = False
    Set book = RestoreFolder(folder)
    RestoreDefinitionSheet folder & "\fields.csv", book, "_復元項目定義"
    If Left$(ReadUtf8(folder & "\format.txt"), 18) = "CaseLedgerBackup/2" Then
        RestoreDefinitionSheet folder & "\mappings.csv", book, "_復元年度対応"
    End If
    book.Worksheets(1).Activate
    message = "別の確認用ブックへ復元しました。正式管理簿へは上書きしていません。値・型・数式を確認し、管理担当が保存先を決めてください。書式・図形はCSVバックアップの対象外です。"
Finished:
    Application.EnableEvents = events
    isBusy = False
    If Len(message) > 0 Then Notify message
    Exit Sub
Failed:
    message = SafeFailure("バックアップ復元", Err.Number)
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    GoTo Finished
End Sub

Public Function ResolveYearAddress(ByVal year As String, ByRef message As String) As String
    Dim definition As Variant
    Dim first As Object
    Dim second As Object
    definition = years(year)
    If Len(CStr(definition(5))) > 0 Then
        message = year & "年度：" & CStr(definition(5))
        Exit Function
    End If
    If Setting("Transport") = "SharePoint" Then
        If Not IsWebAddress(definition(3)) Or Not IsWebAddress(definition(4)) Then
            message = year & "年度：正式運用の保存先はSharePointの直接URLで設定してください。"
            Exit Function
        End If
    ElseIf Setting("Transport") <> "Local" Then
        message = "Transportは自宅試験のLocalまたは正式運用のSharePointを設定してください。"
        Exit Function
    End If
    If StrComp(definition(3), definition(4), vbTextCompare) = 0 Then
        message = year & "年度：常勤用と非常勤用の保存先が同一です。"
        Exit Function
    End If
    Set first = FileState(CStr(definition(3)), message)
    If first Is Nothing Then Exit Function
    Set second = FileState(CStr(definition(4)), message)
    If second Is Nothing Then Exit Function
    If first("Exists") And second("Exists") Then
        message = year & "年度：複数の候補に管理簿が存在します。正式データを確認してください。推測で保存しません。"
    ElseIf Not first("Exists") And Not second("Exists") Then
        message = year & "年度：両方の保存先に管理簿が存在しません。保存先と移動状態を確認してください。"
    ElseIf first("Exists") Then
        ResolveYearAddress = definition(3)
    Else
        ResolveYearAddress = definition(4)
    End If
End Function

Private Function ApiFile(ByVal address As String) As String
    Dim siteUrl As String
    Dim hostEnd As Long
    Dim relative As String
    siteUrl = Setting("SharePointSite")
    If Right$(siteUrl, 1) = "/" Then siteUrl = Left$(siteUrl, Len(siteUrl) - 1)
    If Not IsWebAddress(siteUrl) Then Exit Function
    If LCase$(Left$(address, Len(siteUrl) + 1)) <> LCase$(siteUrl & "/") Then Exit Function
    If InStr(address, "%") > 0 Or InStr(address, "?") > 0 Or InStr(address, "#") > 0 Then Exit Function
    hostEnd = InStr(9, address, "/")
    If hostEnd = 0 Then Exit Function
    relative = Mid$(address, hostEnd)
    ApiFile = siteUrl & "/_api/web/GetFileByServerRelativeUrl('" & Application.WorksheetFunction.EncodeURL(Replace(relative, "'", "''")) & "')"
End Function

Private Function HttpRequest(ByVal method As String, ByVal endpoint As String, ByVal body As Variant, _
    ByRef message As String, Optional ByVal digest As String = "", Optional ByVal etag As String = "") As Object
    Dim request As Object
    Dim started As Double
    Dim elapsed As Double
    On Error GoTo Failed
    If Len(endpoint) = 0 Then
        message = "URL不正：SettingsのSharePointSiteと復号済み直接URLを確認してください。"
        Exit Function
    End If
    Set request = CreateObject("MSXML2.XMLHTTP.6.0")
    request.Open method, endpoint, True
    request.setRequestHeader "Accept", "application/atom+xml"
    If Len(digest) > 0 Then request.setRequestHeader "X-RequestDigest", digest
    If Len(etag) > 0 Then request.setRequestHeader "IF-MATCH", etag
    If IsEmpty(body) Then
        request.send
    Else
        request.send body
    End If
    started = Timer
    Do While request.readyState <> 4
        DoEvents
        elapsed = Timer - started
        If elapsed < 0 Then elapsed = elapsed + 86400#
        If elapsed > 20 Then
            request.abort
            message = "通信タイムアウト。変更要求の場合は結果不明です。再送せず照合してください。"
            Exit Function
        End If
    Loop
    Set HttpRequest = request
    Exit Function
Failed:
    message = "通信失敗：認証・接続を確認してください。エラー番号 " & Err.Number & "。変更要求の場合は結果を照合してください。"
End Function

Private Function XmlText(ByVal xml As Object, ByVal name As String) As String
    Dim node As Object
    Set node = xml.SelectSingleNode("//*[local-name()='" & name & "']")
    If Not node Is Nothing Then XmlText = node.Text
End Function

Public Function FileState(ByVal address As String, ByRef message As String) As Object
    Dim result As Object
    Dim fileSystem As Object
    Dim file As Object
    Dim request As Object
    Dim status As Long
    On Error GoTo Failed
    Set result = NewDictionary()
    result.Add "Exists", False
    result.Add "Id", ""
    result.Add "Size", ""
    result.Add "ETag", ""
    result.Add "Checkout", ""
    If IsWebAddress(address) Then
        Set request = HttpRequest("GET", ApiFile(address), Empty, message)
        If request Is Nothing Then Exit Function
        status = request.status
        Select Case status
            Case 200
                result("Id") = XmlText(request.responseXML, "UniqueId")
                result("Size") = XmlText(request.responseXML, "Length")
                result("ETag") = XmlText(request.responseXML, "ETag")
                result("Checkout") = XmlText(request.responseXML, "CheckOutType")
                If Len(result("Id")) = 0 Or Len(result("ETag")) = 0 Or Len(result("Size")) = 0 Then
                    message = "認証または応答形式未確認：SharePointのファイル情報を取得できません。"
                    Exit Function
                End If
                result("Exists") = True
            Case 404
                '404でもサインインHTML等は不存在の根拠にしない。
                If Len(XmlText(request.responseXML, "code")) = 0 Then
                    message = "所在未確認：404応答ですがSharePointのエラー情報を確認できません。"
                    Exit Function
                End If
            Case 401
                message = "認証失敗（HTTP 401）：SharePointへサインインし直してください。"
                Exit Function
            Case 403
                message = "アクセス権なし（HTTP 403）：管理担当へ権限を確認してください。"
                Exit Function
            Case 400
                message = "URL不正（HTTP 400）：直接ファイルURLを確認してください。"
                Exit Function
            Case Else
                message = "SharePoint通信失敗：HTTP " & status & "。所在を確定できません。"
                Exit Function
        End Select
    Else
        If InStr(address, ":\") <> 2 And Left$(address, 2) <> "\\" Then
            message = "アドレス不正：絶対パスを設定してください。"
            Exit Function
        End If
        Set fileSystem = CreateObject("Scripting.FileSystemObject")
        'GetFileの例外で不存在と権限エラーを分ける。FileExistsだけで欠落と判定しない。
        On Error Resume Next
        Err.Clear
        Set file = fileSystem.GetFile(address)
        status = Err.Number
        On Error GoTo Failed
        If status = 53 Or status = 76 Then
            result("Exists") = False
        ElseIf status <> 0 Then
            message = "ファイル所在確認失敗：アクセス権またはネットワークを確認してください。エラー番号 " & status
            Exit Function
        Else
            result("Exists") = True
            result("Size") = CStr(file.Size)
            result("ETag") = CStr(CDbl(file.DateLastModified)) & "|" & CStr(file.Size)
            result("Checkout") = "2"
        End If
    End If
    Set FileState = result
    Exit Function
Failed:
    message = "所在情報を取得できません。エラー番号 " & Err.Number & "。認証・権限・通信を確認してください。"
End Function

Private Function Digest(ByRef message As String) As String
    Dim request As Object
    Set request = HttpRequest("POST", Setting("SharePointSite") & "/_api/contextinfo", "", message)
    If request Is Nothing Then Exit Function
    If request.status <> 200 Then
        message = "SharePointの変更用認証を取得できません。HTTP " & request.status
        Exit Function
    End If
    Digest = XmlText(request.responseXML, "FormDigestValue")
    If Len(Digest) = 0 Then message = "SharePointの認証応答が不正です。"
End Function

Public Sub StartBusiness()
    RunBusiness True
End Sub

Public Sub FinishBusiness()
    RunBusiness False
End Sub

Private Sub RunBusiness(ByVal isStart As Boolean)
    Dim warnings As String
    Dim message As String
    Dim year As String
    Dim definition As Variant
    Dim source As String
    Dim destination As String
    Dim state As Object
    Dim afterState As Object
    Dim book As Workbook
    Dim evidence As String
    Dim backup As String
    Dim fileSystem As Object
    Dim moved As Boolean
    Dim ws As Worksheet
    If Not BeginAction() Then Exit Sub
    On Error GoTo Failed
    If Setting("ControllerRole") <> "Staff" Then
        message = "業務開始・終了は常勤（Staff）だけが実行できます。非常勤（General）は管理機能を実行できません。"
        GoTo Finished
    End If
    If PendingSave() Then
        message = "保存結果が未確認です。保存結果の照合を先に行ってください。移動は実行していません。"
        GoTo Finished
    End If
    If ThisWorkbook.ReadOnly Or Len(ThisWorkbook.Path) = 0 Or ThisWorkbook.FileFormat <> 52 Then
        message = "CONTROLLERを保存可能なxlsmとして保存してください。管理簿は移動していません。"
        GoTo Finished
    End If
    If Not UiCanLeaveDetail() Then GoTo Finished
    If Not LoadConfiguration(message) Then GoTo Finished
    If Not ReconcileMove(message) Then GoTo Finished
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    year = LatestEnabledYear(message)
    If Len(year) = 0 Or Len(message) > 0 Then GoTo Finished
    definition = years(year)
    source = ResolveYearAddress(CStr(year), message)
    If Len(message) > 0 Then GoTo Finished
    If isStart Then
        destination = definition(4)
    Else
        destination = definition(3)
    End If
    If StrComp(source, destination, vbTextCompare) <> 0 Then
        Set state = FileState(source, message)
        If state Is Nothing Then GoTo Finished
        If state("Checkout") <> "2" Then
            message = CStr(year) & "年度がチェックアウトされています。移動を中止しました。"
            GoTo Finished
        End If
        'OpenLedgerの書込用はSharePointをチェックアウトするため、移動前は読取専用で検証。
        Set book = OpenLedger(source, True, message)
        If book Is Nothing Then GoTo Finished
        If Not book.Saved Then
            message = CStr(year) & "年度：Excelが開いた時点で未保存状態です。再計算・リンク・管理簿の内容を確認してください。移動は実行しません。"
            GoTo Finished
        End If
        evidence = NewEvidencePath(book.Name)
        book.SaveCopyAs evidence
        If isStart Then
            backup = CreateBackup(book, CStr(year), message)
            If Len(message) > 0 Then GoTo Finished
        End If
        book.Close SaveChanges:=False
        Set book = Nothing
        Set afterState = FileState(source, message)
        If afterState Is Nothing Then GoTo Finished
        If state("ETag") <> afterState("ETag") Then
            message = CStr(year) & "年度：バックアップ中に管理簿が変更されました。移動は実行していません。"
            GoTo Finished
        End If
        Set ws = Sheet("MoveRecovery")
        PutValue ws.Range("B1"), "移動結果確認待ち"
        PutValue ws.Range("B2"), source
        PutValue ws.Range("B3"), destination
        PutValue ws.Range("B4"), state("Id")
        PutValue ws.Range("B5"), state("Size")
        PutValue ws.Range("B6"), state("ETag")
        PutValue ws.Range("B7"), evidence
        ThisWorkbook.Save
        If Not ThisWorkbook.Saved Then
            message = "移動予定を退避できないため、管理簿は移動しません。"
            GoTo Finished
        End If
        moved = MoveFile(source, destination, state, message)
        If Not moved Then GoTo Finished
        If Not ReconcileMove(message) Then GoTo Finished
        If isStart Then
            If Not PruneBackups(CStr(year), message) Then
                warnings = warnings & vbCrLf & CStr(year) & "年度：" & message
                message = ""
            End If
        End If
    End If
    Set Cases = Nothing
    Set currentCase = Nothing
    SetSetting "BusinessState", IIf(isStart, "業務開始", "業務終了") & " " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ThisWorkbook.Save
    message = year & "年度の" & IIf(isStart, "業務開始", "業務終了") & "の移動確認が完了しました。過年度の管理簿は移動していません。StartControllerで最新を再読込してください。" & warnings
Finished:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    isBusy = False
    If Len(message) > 0 Then Notify message
    Exit Sub
Failed:
    message = SafeFailure("最新年度の管理簿移動", Err.Number) & " 移動結果を確認するまで再操作しないでください。"
    Resume Finished
End Sub

Private Function LatestEnabledYear(ByRef message As String) As String
    Dim year As Variant
    Dim latestValue As Double
    If years Is Nothing Then
        message = "年度設定が未読込です。管理簿は移動しません。"
        Exit Function
    End If
    For Each year In years.Keys
        If Len(CStr(year)) = 0 Or Len(CStr(year)) > 8 Then
            LatestEnabledYear = ""
            message = "最新年度を特定できません。Settingsの有効年度は1～8桁の西暦数字で設定してください。管理簿は移動しません。"
            Exit Function
        End If
        If Not IsDigits(CStr(year), Len(CStr(year))) Then
            LatestEnabledYear = ""
            message = "最新年度を特定できません。Settingsの有効年度は西暦数字で設定してください。管理簿は移動しません。"
            Exit Function
        End If
        If CDbl(year) > latestValue Or Len(LatestEnabledYear) = 0 Then
            latestValue = CDbl(year)
            LatestEnabledYear = CStr(year)
        End If
    Next
    If Len(LatestEnabledYear) = 0 Then message = "有効な年度がありません。管理簿は移動しません。"
End Function

Private Function MoveFile(ByVal source As String, ByVal destination As String, ByVal expected As Object, ByRef message As String) As Boolean
    Dim current As Object
    Dim target As Object
    Dim fileSystem As Object
    Dim request As Object
    Dim digestValue As String
    Dim relative As String
    On Error GoTo Failed
    Set current = FileState(source, message)
    If current Is Nothing Then Exit Function
    Set target = FileState(destination, message)
    If target Is Nothing Then Exit Function
    If Not current("Exists") Or target("Exists") Or current("ETag") <> expected("ETag") Then
        message = "移動元・移動先が確認時から変化しています。移動は実行しません。"
        Exit Function
    End If
    If IsWebAddress(source) Then
        If Not IsWebAddress(destination) Or Len(ApiFile(destination)) = 0 Then
            message = "SharePoint内の同一サイトへの移動だけを扱います。移動先設定を確認してください。"
            Exit Function
        End If
        digestValue = Digest(message)
        If Len(digestValue) = 0 Then Exit Function
        relative = Mid$(destination, InStr(9, destination, "/"))
        Set request = HttpRequest("POST", ApiFile(source) & "/moveto(newurl='" & _
            Application.WorksheetFunction.EncodeURL(Replace(relative, "'", "''")) & "',flags=0)", "", message, digestValue, expected("ETag"))
        If request Is Nothing Then Exit Function
        If request.status <> 200 And request.status <> 204 Then
            message = "移動要求の結果未確認：HTTP " & request.status & "。再送せず移動結果を確認してください。"
            Exit Function
        End If
    Else
        If IsWebAddress(destination) Then
            message = "ローカルとSharePointの間では移動しません。"
            Exit Function
        End If
        Set fileSystem = CreateObject("Scripting.FileSystemObject")
        If StrComp(fileSystem.GetDriveName(source), fileSystem.GetDriveName(destination), vbTextCompare) <> 0 Then
            message = "ローカル試験は同一ドライブ内のフォルダ間移動に限定します。"
            Exit Function
        End If
        fileSystem.MoveFile source, destination
    End If
    MoveFile = True
    Exit Function
Failed:
    message = "移動結果未確認です。エラー番号 " & Err.Number & "。再送せず移動結果を確認してください。"
End Function

Public Sub ConfirmMoveResult()
    Dim message As String
    If Not BeginAction() Then Exit Sub
    On Error GoTo Failed
    If Setting("ControllerRole") <> "Staff" Then
        message = "移動結果の管理操作は常勤（Staff）だけが実行できます。"
    ElseIf ReconcileMove(message) Then
        message = "移動待ち記録の照合が完了しました。最新を再読込してください。"
    End If
Finished:
    isBusy = False
    If Len(message) > 0 Then Notify message
    Exit Sub
Failed:
    message = SafeFailure("移動結果確認", Err.Number)
    Resume Finished
End Sub

Private Function ReconcileMove(ByRef message As String) As Boolean
    Dim ws As Worksheet
    Dim first As Object
    Dim second As Object
    Dim book As Workbook
    Dim evidence As Workbook
    Dim address As String
    On Error GoTo Failed
    Set ws = Sheet("MoveRecovery")
    If Len(CStr(ws.Range("B1").Value2)) = 0 Then
        ReconcileMove = True
        Exit Function
    End If
    Set first = FileState(CStr(ws.Range("B2").Value2), message)
    If first Is Nothing Then Exit Function
    Set second = FileState(CStr(ws.Range("B3").Value2), message)
    If second Is Nothing Then Exit Function
    If CBool(first("Exists")) = CBool(second("Exists")) Then GoTo Unconfirmed
    If second("Exists") Then
        If second("Size") <> CStr(ws.Range("B5").Value2) Then GoTo Unconfirmed
        If Len(CStr(ws.Range("B4").Value2)) > 0 Then
            If second("Id") <> CStr(ws.Range("B4").Value2) Then GoTo Unconfirmed
        End If
        address = CStr(ws.Range("B3").Value2)
    Else
        If first("ETag") <> CStr(ws.Range("B6").Value2) Then GoTo Unconfirmed
        address = CStr(ws.Range("B2").Value2)
    End If
    Set book = OpenLedger(address, True, message)
    If book Is Nothing Then Exit Function
    Set evidence = OpenLedger(CStr(ws.Range("B7").Value2), True, message)
    If evidence Is Nothing Then GoTo Unconfirmed
    If Not CompareBooks(book, evidence) Then GoTo Unconfirmed
    book.Close SaveChanges:=False
    Set book = Nothing
    evidence.Close SaveChanges:=False
    Set evidence = Nothing
    ws.Range("B1").ClearContents
    ThisWorkbook.Save
    If Not ThisWorkbook.Saved Then GoTo Unconfirmed
    '照合用の自動削除は行わない。移動証跡は担当者が照合後に整理する。
    ReconcileMove = True
    message = ""
    Exit Function
Unconfirmed:
    message = "移動結果を照合できません。元・先・証跡を管理担当が確認してください。上書き・再送は行いません。"
    GoTo Cleanup
Failed:
    message = SafeFailure("移動結果照合", Err.Number)
Cleanup:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    If Not evidence Is Nothing Then evidence.Close SaveChanges:=False
    PutValue ws.Range("B1"), "移動結果確認待ち"
    On Error GoTo 0
End Function

Private Function NewEvidencePath(ByVal fileName As String) As String
    Dim fileSystem As Object
    Dim folder As String
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    folder = fileSystem.BuildPath(fileSystem.GetSpecialFolder(2), "Controller_" & fileSystem.GetTempName)
    fileSystem.CreateFolder folder
    NewEvidencePath = folder & "\Evidence_" & fileName
End Function

Private Sub SetSetting(ByVal name As String, ByVal value As String)
    Dim rowIndex As Long
    For rowIndex = 2 To LastRow(Sheet("Settings"))
        If CStr(Sheet("Settings").Cells(rowIndex, 1).Value2) = name Then
            PutValue Sheet("Settings").Cells(rowIndex, 2), value
            Exit Sub
        End If
    Next
End Sub


Private Function EncodeValue(ByVal value As Variant) As String
    If IsError(value) Then
        Err.Raise vbObjectError + 120, "Backup", "Excelエラー値はバックアップできません。元データを確認してください。"
    End If
    Select Case VarType(value)
        Case vbEmpty
        EncodeValue = "E:"
        Case vbString
        EncodeValue = "S:" & CStr(value)
        Case vbBoolean
        EncodeValue = "B:" & IIf(value, "1", "0")
        Case Else
            If Not (IsNumeric(value)) Then
                Err.Raise vbObjectError + 120, "Backup", "未対応のデータ型です。"
            End If
            EncodeValue = "N:" & Trim$(Str$(CDbl(value)))
    End Select
End Function

Private Function DecodeValue(ByVal value As String) As Variant
    Select Case Left$(value, 2)
        Case "E:"
        DecodeValue = Empty
        Case "S:"
        DecodeValue = Mid$(value, 3)
        Case "B:"
        DecodeValue = (Mid$(value, 3) = "1")
        Case "N:"
        DecodeValue = Val(Mid$(value, 3))
        Case Else
            Err.Raise vbObjectError + 120, "Backup", "バックアップの型情報が不正です。"
    End Select
End Function

Private Function CsvCell(ByVal value As String) As String
    CsvCell = """" & Replace(value, """", """""") & """"
End Function

Private Sub WriteUtf8(ByVal address As String, ByVal value As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText value
    stream.SaveToFile address, 2
    stream.Close
End Sub

Private Function ReadUtf8(ByVal address As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile address
    ReadUtf8 = stream.ReadText
    stream.Close
End Function

Private Function ParseCsv(ByVal text As String) As Collection
    Dim rows As New Collection
    Dim row As Collection
    Dim rowIndex As Long
    Dim ch As String
    Dim value As String
    Dim quoted As Boolean
    Set row = New Collection
    rowIndex = 1
    Do While rowIndex <= Len(text)
        ch = Mid$(text, rowIndex, 1)
        If ch = """" Then
            If quoted And Mid$(text, rowIndex + 1, 1) = """" Then
                value = value & """"
                rowIndex = rowIndex + 1
            Else
                quoted = Not quoted
            End If
        ElseIf ch = "," And Not quoted Then
            row.Add value
            value = ""
        ElseIf ch = vbCr And Not quoted Then
            If Not (Mid$(text, rowIndex + 1, 1) = vbLf) Then
                Err.Raise vbObjectError + 120, "Backup", "CSVの改行が不正です。"
            End If
            row.Add value
            rows.Add row
            Set row = New Collection
            value = ""
            rowIndex = rowIndex + 1
        Else
            value = value & ch
        End If
        rowIndex = rowIndex + 1
    Loop
    If quoted Then
        Err.Raise vbObjectError + 120, "Backup", "CSVの引用符が閉じていません。"
    End If
    If Len(value) > 0 Or row.Count > 0 Then
        row.Add value
        rows.Add row
    End If
    Set ParseCsv = rows
End Function

Private Function SheetCsv(ByVal ws As Worksheet) As String
    Dim data As Variant
    Dim formulas As Variant
    Dim lines() As String
    Dim cells() As String
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim rowCount As Long
    Dim columnCount As Long
    Dim value As String
    rowCount = LastRow(ws)
    columnCount = LastColumn(ws)
    '最低2x2として取得し、単一セルでも必ず2次元配列にする。
    data = ws.Cells(1, 1).Resize(Application.Max(2, rowCount), Application.Max(2, columnCount)).Value2
    formulas = ws.Cells(1, 1).Resize(Application.Max(2, rowCount), Application.Max(2, columnCount)).Formula
    ReDim lines(0 To rowCount)
    lines(0) = CsvCell(ws.Name) & "," & CsvCell(CStr(rowCount)) & "," & CsvCell(CStr(columnCount)) & "," & CsvCell(CStr(ws.Visible))
    ReDim cells(0 To columnCount - 1)
    For rowIndex = 1 To rowCount
        For colIndex = 1 To columnCount
            value = EncodeValue(data(rowIndex, colIndex))
            If VarType(formulas(rowIndex, colIndex)) = vbString Then
                If Left$(CStr(formulas(rowIndex, colIndex)), 1) = "=" And ws.Cells(rowIndex, colIndex).HasFormula Then value = "F:" & formulas(rowIndex, colIndex)
            End If
            cells(colIndex - 1) = CsvCell(value)
        Next
        lines(rowIndex) = Join(cells, ",")
    Next
    SheetCsv = Join(lines, vbCrLf) & vbCrLf
End Function

Private Function NewTempFolder() As String
    Dim fileSystem As Object
    Dim folder As String
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    folder = fileSystem.BuildPath(fileSystem.GetSpecialFolder(2), fileSystem.GetTempName)
    fileSystem.CreateFolder folder
    NewTempFolder = folder
End Function

Private Sub CreateZip(ByVal sourceFolder As String, ByVal zipFile As String)
    Dim shell As Object
    Dim folder As Object
    Dim destination As Object
    Dim stream As Object
    Dim bytes(0 To 21) As Byte
    Dim started As Double
    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If fileSystem.FileExists(zipFile) Then
        Err.Raise vbObjectError + 120, "Backup", "ZIPの上書きは行いません。"
    End If
    'Windows標準の圧縮フォルダー機能。Windows API・外部実行ファイルは使わない。
    bytes(0) = 80
    bytes(1) = 75
    bytes(2) = 5
    bytes(3) = 6
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write bytes
    stream.SaveToFile zipFile, 1
    stream.Close
    Set shell = CreateObject("Shell.Application")
    Set folder = shell.Namespace(CVar(sourceFolder))
    Set destination = shell.Namespace(CVar(zipFile))
    If Not (Not folder Is Nothing And Not destination Is Nothing) Then
        Err.Raise vbObjectError + 120, "Backup", "WindowsのZIP機能を利用できません。バックアップ未作成のため移動しません。"
    End If
    destination.CopyHere folder.Items, 1556
    started = Timer
    Do
        DoEvents
        If destination.Items.Count = folder.Items.Count Then Exit Do
        If Not (ElapsedSeconds(started) < 30) Then
            Err.Raise vbObjectError + 120, "Backup", "ZIP作成が完了しませんでした。管理簿は移動しません。"
        End If
    Loop
End Sub

Private Sub ExpandZip(ByVal zipFile As String, ByVal folderPath As String)
    Dim shell As Object
    Dim source As Object
    Dim target As Object
    Dim started As Double
    Set shell = CreateObject("Shell.Application")
    Set source = shell.Namespace(CVar(zipFile))
    Set target = shell.Namespace(CVar(folderPath))
    If Not (Not source Is Nothing And Not target Is Nothing) Then
        Err.Raise vbObjectError + 120, "Backup", "ZIPの展開機能を利用できません。"
    End If
    target.CopyHere source.Items, 1556
    started = Timer
    Do
        DoEvents
        If target.Items.Count = source.Items.Count Then Exit Do
        If Not (ElapsedSeconds(started) < 30) Then
            Err.Raise vbObjectError + 120, "Backup", "ZIPの展開が完了しませんでした。"
        End If
    Loop
End Sub

Private Function RestoreFolder(ByVal folder As String) As Workbook
    Dim book As Workbook
    Dim ws As Worksheet
    Dim manifest As Collection
    Dim entry As Variant
    Dim rows As Collection
    Dim header As Collection
    Dim line As Collection
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim count As Long
    Dim data() As Variant
    Dim value As String
    Dim formulaCells As New Collection
    Dim cell As Variant
    Dim description As String
    Dim errorNumber As Long
    On Error GoTo Failed
    If Left$(ReadUtf8(folder & "\format.txt"), 18) <> "ControllerBackup/3" And Left$(ReadUtf8(folder & "\format.txt"), 18) <> "CaseLedgerBackup/2" Then
        Err.Raise vbObjectError + 120, "Backup", "復元形式が違います。"
    End If
    Set manifest = ParseCsv(ReadUtf8(folder & "\manifest.csv"))
    Set book = Application.Workbooks.Add(xlWBATWorksheet)
    book.Date1904 = (InStr(ReadUtf8(folder & "\format.txt"), "Date1904=True") > 0)
    For Each entry In manifest
        If Not (Left$(CStr(entry(1)), 5) = "sheet" And InStr(CStr(entry(1)), "\") = 0 And InStr(CStr(entry(1)), "/") = 0) Then
            Err.Raise vbObjectError + 120, "Backup", "復元ファイル名が不正です。"
        End If
        Set rows = ParseCsv(ReadUtf8(folder & "\" & entry(1)))
        Set header = rows(1)
        count = count + 1
        If count = 1 Then Set ws = book.Worksheets(1) Else Set ws = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count))
        ws.Name = CStr(header(1))
        ReDim data(1 To CLng(header(2)), 1 To CLng(header(3)))
        Set formulaCells = New Collection
        If Not (rows.Count = UBound(data, 1) + 1) Then
            Err.Raise vbObjectError + 120, "Backup", "復元行数が一致しません。"
        End If
        For rowIndex = 1 To UBound(data, 1)
            Set line = rows(rowIndex + 1)
            If Not (line.Count = UBound(data, 2)) Then
                Err.Raise vbObjectError + 120, "Backup", "復元列数が一致しません。"
            End If
            For colIndex = 1 To UBound(data, 2)
                value = CStr(line(colIndex))
                If Left$(value, 2) = "F:" Then
                    formulaCells.Add Array(rowIndex, colIndex, Mid$(value, 3))
                Else
                    data(rowIndex, colIndex) = DecodeValue(value)
                End If
            Next
        Next
        '先に全体を文字形式にし、数式として解釈される文字列の実行を防ぐ。
        ws.Cells(1, 1).Resize(UBound(data, 1), UBound(data, 2)).NumberFormat = "@"
        ws.Cells(1, 1).Resize(UBound(data, 1), UBound(data, 2)).Value2 = data
        For Each cell In formulaCells
            ws.Cells(cell(0), cell(1)).NumberFormat = "General"
            ws.Cells(cell(0), cell(1)).Formula = cell(2)
        Next
    Next
    count = 0
    For Each entry In manifest
        count = count + 1
        Set rows = ParseCsv(ReadUtf8(folder & "\" & entry(1)))
        Set header = rows(1)
        book.Worksheets(count).Visible = CLng(header(4))
    Next
    Set RestoreFolder = book
    Exit Function
Failed:

    errorNumber = Err.Number
    description = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise errorNumber, "RestoreFolder", description
End Function

Private Function CompareBooks(ByVal first As Workbook, ByVal second As Workbook) As Boolean
    Dim rowIndex As Long
    If first.Date1904 <> second.Date1904 Or first.Worksheets.Count <> second.Worksheets.Count Then Exit Function
    For rowIndex = 1 To first.Worksheets.Count
        If SheetCsv(first.Worksheets(rowIndex)) <> SheetCsv(second.Worksheets(rowIndex)) Then Exit Function
    Next
    CompareBooks = True
End Function

Private Function ManagedBackupName(ByVal name As String, ByVal prefix As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^" & prefix & "[0-9]{8}_[0-9]{6}_[0-9]+\.zip$"
    ManagedBackupName = regex.Test(name)
End Function

Private Sub RestoreDefinitionSheet(ByVal address As String, ByVal book As Workbook, ByVal name As String)
    Dim rows As Collection
    Dim header As Collection
    Dim line As Collection
    Dim data() As Variant
    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim colIndex As Long
    Set rows = ParseCsv(ReadUtf8(address))
    Set header = rows(1)
    ReDim data(1 To CLng(header(2)), 1 To CLng(header(3)))
    For rowIndex = 1 To UBound(data, 1)
        Set line = rows(rowIndex + 1)
        For colIndex = 1 To UBound(data, 2)
            data(rowIndex, colIndex) = DecodeValue(CStr(line(colIndex)))
        Next
    Next
    Set ws = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count))
    ws.Name = name
    ws.Cells(1, 1).Resize(UBound(data, 1), UBound(data, 2)).NumberFormat = "@"
    ws.Cells(1, 1).Resize(UBound(data, 1), UBound(data, 2)).Value2 = data
    ws.Visible = xlSheetVeryHidden
End Sub
