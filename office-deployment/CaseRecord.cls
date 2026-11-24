VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CaseRecord"
Option Explicit

Private fieldValues As Object
Private originalValues As Object
Private uses1904Dates As Boolean

' Date system of the workbook supplying this record's raw Value2 data.
Public Property Get SourceUses1904() As Boolean
    SourceUses1904 = uses1904Dates
End Property

Public Property Let SourceUses1904(ByVal newValue As Boolean)
    uses1904Dates = newValue
End Property

Private Sub Class_Initialize()
    Set fieldValues = CreateObject("Scripting.Dictionary")
    Set originalValues = CreateObject("Scripting.Dictionary")
End Sub

' 2次元配列の指定行を1事案として取り込む。
' fieldNames はデータの列順と一致する1次元配列。
Public Sub LoadData(ByRef data As Variant, ByVal rowIndex As Long, ByRef fieldNames As Variant)
    Dim columnIndex As Long
    Dim fieldIndex As Long
    Dim fieldName As String

    fieldValues.RemoveAll
    originalValues.RemoveAll

    For columnIndex = LBound(data, 2) To UBound(data, 2)
        fieldIndex = LBound(fieldNames) + columnIndex - LBound(data, 2)
        fieldName = CStr(fieldNames(fieldIndex))
        ' 同じ見出しが複数存在する場合は先頭の列を採用する。
        If Len(fieldName) > 0 Then
            If Not fieldValues.Exists(fieldName) Then
                fieldValues.Add fieldName, data(rowIndex, columnIndex)
            End If
        End If
    Next columnIndex
End Sub

' 指定した項目の現在値。
Public Property Get Value(ByVal fieldName As String) As Variant
    Value = fieldValues(fieldName)
End Property

' 指定した項目を変更し、初回変更時の値を保存する。
Public Property Let Value(ByVal fieldName As String, ByVal newValue As Variant)
    If Not fieldValues.Exists(fieldName) Then
        Err.Raise vbObjectError + 510, "CaseRecord", "項目がありません: " & fieldName
    End If
    If Not originalValues.Exists(fieldName) Then
        originalValues.Add fieldName, fieldValues(fieldName)
    End If
    fieldValues(fieldName) = newValue
End Property

' 読込時と現在値が異なる項目だけを返す。
Public Function GetChanges() As Object
    Dim changes As Object
    Dim fieldName As Variant

    Set changes = CreateObject("Scripting.Dictionary")
    For Each fieldName In originalValues.Keys
        If Not SameValue(originalValues(fieldName), fieldValues(fieldName)) Then
            changes.Add CStr(fieldName), fieldValues(fieldName)
        End If
    Next fieldName
    Set GetChanges = changes
End Function

' 指定した変更項目の読込時の値を返す。
Public Function OriginalValue(ByVal fieldName As String) As Variant
    OriginalValue = originalValues(fieldName)
End Function

' 比較は保存前の競合確認にも使用する。
Public Function SameValue(ByVal leftValue As Variant, ByVal rightValue As Variant) As Boolean
    If IsError(leftValue) Or IsError(rightValue) Then
        SameValue = False
    ElseIf IsNull(leftValue) Or IsNull(rightValue) Then
        SameValue = IsNull(leftValue) And IsNull(rightValue)
    ElseIf IsEmpty(leftValue) Or IsEmpty(rightValue) Then
        SameValue = (Len(CStr(leftValue)) = 0 And Len(CStr(rightValue)) = 0)
    Else
        SameValue = (CStr(leftValue) = CStr(rightValue))
    End If
End Function
