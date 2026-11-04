Option Explicit

Public sourceYear As String
Public sourceAddress As String
Public sourceSheet As String
Public sourceRow As Long
Public originalRow As Variant
Public sourceColumns As Object
Public values As Object
Public invalidReason As String

Private Sub Class_Initialize()
    Set values = CreateObject("Scripting.Dictionary")
End Sub

Public Property Get PersonKey() As String
    Dim number As String
    If Not RestoreNumber(values("MaskedNumber"), number) Then Exit Property
    If Not IsDigits(values("Branch"), 5) Then Exit Property
    PersonKey = number & values("Branch")
End Property

Public Property Get Key() As String
    If Len(PersonKey) = 0 Then
        Key = sourceYear & "|InvalidRow|" & CStr(sourceRow)
    Else
        Key = sourceYear & "|" & PersonKey
    End If
End Property
