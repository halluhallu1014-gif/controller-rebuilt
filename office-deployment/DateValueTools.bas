Attribute VB_Name = "DateValueTools"
Option Explicit

' General date parser: Date, Excel serial, yyyyMMdd, yyyy/m/d, yyyy-m-d, yyyy年m月d日.
' Serial values use the specified Excel workbook date system.
' A date field is day-granular; time fractions are discarded.
Public Function TryParseDateValue(ByVal sourceValue As Variant, ByRef parsedDate As Date, _
                                  Optional ByVal uses1904 As Boolean = False) As Boolean
    Dim textValue As String
    Dim parts As Variant
    Dim numericValue As Double
    Dim serialDay As Long

    On Error GoTo InvalidValue
    If IsObject(sourceValue) Or IsArray(sourceValue) Then Exit Function
    If IsError(sourceValue) Or IsNull(sourceValue) Or IsEmpty(sourceValue) Then Exit Function
    If VarType(sourceValue) = vbBoolean Then Exit Function
    If VarType(sourceValue) = vbDate Then
        parsedDate = DateSerial(Year(sourceValue), Month(sourceValue), Day(sourceValue))
        TryParseDateValue = True
        Exit Function
    End If

    textValue = Trim$(CStr(sourceValue))
    If Len(textValue) = 0 Then Exit Function

    ' Eight digits always denote a calendar date, never an Excel serial.
    If Len(textValue) = 8 And IsDigitText(textValue) Then
        TryParseDateValue = TryBuildCalendarDate(Left$(textValue, 4), _
                            Mid$(textValue, 5, 2), Right$(textValue, 2), parsedDate)
        Exit Function
    End If

    ' Explicit year/month/day only; do not guess ambiguous locale-dependent dates.
    textValue = Replace(textValue, "年", "/")
    textValue = Replace(textValue, "月", "/")
    If Right$(textValue, 1) = "日" Then textValue = Left$(textValue, Len(textValue) - 1)
    textValue = Replace(textValue, "-", "/")
    If InStr(textValue, "/") > 0 Then
        parts = Split(textValue, "/")
        If UBound(parts) <> 2 Then Exit Function
        TryParseDateValue = TryBuildCalendarDate(CStr(parts(0)), CStr(parts(1)), _
                                                 CStr(parts(2)), parsedDate)
        Exit Function
    End If

    If Not IsNumeric(sourceValue) Then Exit Function
    numericValue = CDbl(sourceValue)
    If uses1904 Then
        If numericValue < 0 Or numericValue >= 2957004# Then Exit Function
        serialDay = CLng(Fix(numericValue))
        parsedDate = DateAdd("d", serialDay, DateSerial(1904, 1, 1))
    Else
        If numericValue < 1 Or numericValue >= 2958466# Then Exit Function
        serialDay = CLng(Fix(numericValue))
        ' Excel's fictional 1900-02-29 cannot be represented as a real date.
        If serialDay = 60 Then Exit Function
        If serialDay < 60 Then serialDay = serialDay + 1
        parsedDate = CDate(serialDay)
    End If
    TryParseDateValue = True
    Exit Function
InvalidValue:
    TryParseDateValue = False
End Function

Private Function IsDigitText(ByVal textValue As String) As Boolean
    If Len(textValue) = 0 Then Exit Function
    IsDigitText = Not (textValue Like "*[!0-9]*")
End Function

Private Function TryBuildCalendarDate(ByVal yearText As String, ByVal monthText As String, _
                                      ByVal dayText As String, ByRef parsedDate As Date) As Boolean
    Dim yearNumber As Long
    Dim monthNumber As Long
    Dim dayNumber As Long
    Dim candidate As Date
    If Len(yearText) <> 4 Or Len(monthText) > 2 Or Len(dayText) > 2 Then Exit Function
    If Not IsDigitText(yearText) Or Not IsDigitText(monthText) Or Not IsDigitText(dayText) Then Exit Function
    yearNumber = CLng(yearText)
    monthNumber = CLng(monthText)
    dayNumber = CLng(dayText)
    If yearNumber < 100 Or monthNumber < 1 Or monthNumber > 12 Or dayNumber < 1 Or dayNumber > 31 Then Exit Function
    candidate = DateSerial(yearNumber, monthNumber, dayNumber)
    ' DateSerial rolls invalid dates forward; reject them rather than changing the date.
    If Year(candidate) <> yearNumber Or Month(candidate) <> monthNumber Or Day(candidate) <> dayNumber Then Exit Function
    parsedDate = candidate
    TryBuildCalendarDate = True
End Function

Public Function IsBlankDateValue(ByVal sourceValue As Variant) As Boolean
    If IsError(sourceValue) Or IsObject(sourceValue) Or IsArray(sourceValue) Then Exit Function
    If IsNull(sourceValue) Or IsEmpty(sourceValue) Then
        IsBlankDateValue = True
    ElseIf VarType(sourceValue) = vbString Then
        IsBlankDateValue = (Len(Trim$(CStr(sourceValue))) = 0)
    End If
End Function

' Convert a real calendar date to a serial in the destination workbook.
Public Function ExcelDateSerial(ByVal parsedDate As Date, Optional ByVal uses1904 As Boolean = False) As Double
    Dim dayValue As Date
    dayValue = DateSerial(Year(parsedDate), Month(parsedDate), Day(parsedDate))
    If uses1904 Then
        If dayValue < DateSerial(1904, 1, 1) Then Err.Raise vbObjectError + 951, "ExcelDateSerial", "日付が1904日付システムの範囲外です。"
        ExcelDateSerial = CDbl(dayValue) - CDbl(DateSerial(1904, 1, 1))
    Else
        If dayValue < DateSerial(1900, 1, 1) Then Err.Raise vbObjectError + 952, "ExcelDateSerial", "日付がExcelの日付範囲外です。"
        ExcelDateSerial = CDbl(dayValue)
        If dayValue < DateSerial(1900, 3, 1) Then ExcelDateSerial = ExcelDateSerial - 1
    End If
End Function

' Valid dates compare by calendar day. Invalid existing values compare as raw text.
Public Function AreDateValuesEqual(ByVal firstValue As Variant, ByVal secondValue As Variant, _
                                   Optional ByVal firstUses1904 As Boolean = False, _
                                   Optional ByVal secondUses1904 As Boolean = False) As Boolean
    Dim firstDate As Date
    Dim secondDate As Date
    Dim firstValid As Boolean
    Dim secondValid As Boolean

    If IsBlankDateValue(firstValue) Or IsBlankDateValue(secondValue) Then
        AreDateValuesEqual = IsBlankDateValue(firstValue) And IsBlankDateValue(secondValue)
        Exit Function
    End If
    If IsError(firstValue) Or IsError(secondValue) Then
        If IsError(firstValue) And IsError(secondValue) Then AreDateValuesEqual = (CStr(firstValue) = CStr(secondValue))
        Exit Function
    End If
    firstValid = TryParseDateValue(firstValue, firstDate, firstUses1904)
    secondValid = TryParseDateValue(secondValue, secondDate, secondUses1904)
    If firstValid And secondValid Then
        AreDateValuesEqual = (firstDate = secondDate)
    ElseIf Not firstValid And Not secondValid Then
        AreDateValuesEqual = (CStr(firstValue) = CStr(secondValue))
    End If
End Function

' Use actual Excel dates in cells; keep unrecognized source values visible and intact.
Public Sub DisplayDateValue(ByVal targetCell As Range, ByVal sourceValue As Variant, _
                            Optional ByVal sourceUses1904 As Boolean = False)
    Dim parsedDate As Date
    Dim targetUses1904 As Boolean
    targetUses1904 = targetCell.Worksheet.Parent.Date1904
    targetCell.NumberFormat = "yyyy/mm/dd"
    If IsBlankDateValue(sourceValue) Then
        targetCell.MergeArea.ClearContents
    ElseIf TryParseDateValue(sourceValue, parsedDate, sourceUses1904) Then
        If (targetUses1904 And parsedDate < DateSerial(1904, 1, 1)) Or parsedDate < DateSerial(1900, 1, 1) Then
            targetCell.NumberFormat = "@"
            targetCell.Value2 = Format$(parsedDate, "yyyy/mm/dd")
        Else
            targetCell.Value2 = ExcelDateSerial(parsedDate, targetUses1904)
        End If
    Else
        targetCell.NumberFormat = "@"
        If IsError(sourceValue) Then
            targetCell.Value2 = sourceValue
        Else
            targetCell.Value2 = CStr(sourceValue)
        End If
    End If
End Sub

' Pure-function test: no workbook is opened, edited or saved.
Public Sub TestDateValueTools()
    Dim expectedDate As Date
    Dim actualDate As Date
    Dim serialValue As Double
    Dim samples As Variant
    Dim sample As Variant
    Dim invalidSamples As Variant

    On Error GoTo Failed
    expectedDate = DateSerial(2026, 11, 24)
    serialValue = ExcelDateSerial(expectedDate)
    samples = Array("20261124", 20261124, "2026/11/24", "2026-11-24", "2026年11月24日", _
                    expectedDate, serialValue, CStr(serialValue), serialValue + 0.75)
    For Each sample In samples
        If Not TryParseDateValue(sample, actualDate) Then Err.Raise vbObjectError + 953, , "変換失敗: " & CStr(sample)
        If actualDate <> expectedDate Then Err.Raise vbObjectError + 954, , "日付不一致: " & CStr(sample)
        If Not AreDateValuesEqual(sample, serialValue) Then Err.Raise vbObjectError + 955, , "比較失敗: " & CStr(sample)
    Next sample

    invalidSamples = Array(Empty, Null, "", " ", "未処理", "20260230", 20261301, "2026/2/29", _
                           0, 60, 2958466#, True, CVErr(xlErrNA))
    For Each sample In invalidSamples
        If TryParseDateValue(sample, actualDate) Then Err.Raise vbObjectError + 956, , "不正値を日付として受理しました。"
    Next sample
    If Not TryParseDateValue("20240229", actualDate) Then Err.Raise vbObjectError + 957, , "うるう日変換失敗"
    If actualDate <> DateSerial(2024, 2, 29) Then Err.Raise vbObjectError + 957, , "うるう日不一致"
    If Not AreDateValuesEqual(Empty, "") Then Err.Raise vbObjectError + 958, , "空欄比較失敗"
    If Not AreDateValuesEqual("未処理", "未処理") Then Err.Raise vbObjectError + 958, , "不正値保持失敗"
    If Not AreDateValuesEqual(CVErr(xlErrNA), CVErr(xlErrNA)) Then Err.Raise vbObjectError + 958, , "エラー値比較失敗"
    If AreDateValuesEqual("20261124", "20261125") Then Err.Raise vbObjectError + 958, , "異なる日付の比較失敗"
    If Not AreDateValuesEqual(serialValue, ExcelDateSerial(expectedDate, True), False, True) Then Err.Raise vbObjectError + 959, , "1904日付比較失敗"
    If Not TryParseDateValue(1, actualDate) Then Err.Raise vbObjectError + 959, , "初期シリアル変換失敗"
    If actualDate <> DateSerial(1900, 1, 1) Then Err.Raise vbObjectError + 959, , "初期シリアル不一致"
    Debug.Print "DATE_TOOLS=PASS"
    Exit Sub
Failed:
    Debug.Print "DATE_TOOLS=ERROR : " & Err.Description
    MsgBox Err.Description, vbExclamation, "日付単体テスト"
End Sub
