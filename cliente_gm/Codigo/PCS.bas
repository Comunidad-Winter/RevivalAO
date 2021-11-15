Attribute VB_Name = "PCS"
Public hds(1 To 1000) As String

Function Hayhd(serial As String) As Boolean
Dim i As Integer
For i = 1 To 1000
If (hds(i) = serial) Then
Hayhd = True
Exit Function
End If
Next i
Hayhd = False
End Function


