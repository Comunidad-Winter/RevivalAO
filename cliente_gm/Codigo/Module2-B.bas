Attribute VB_Name = "Module2"
Option Explicit
Declare Function KILL_PROC_BY_NAME Lib "killproc" (ByVal FileName As String) As Long

Private Function ConvToHex(X As Integer) As String
    If X > 9 Then
        ConvToHex = Chr(X + 55)
    Else
        ConvToHex = CStr(X)
    End If
End Function

Private Function ConvToInt(X As String) As Integer
    
    Dim X1 As String
    Dim X2 As String
    Dim Temp As Integer
    
    X1 = mid(X, 1, 1)
    X2 = mid(X, 2, 1)
    
    If IsNumeric(X1) Then
        Temp = 16 * Int(X1)
    Else
        Temp = (Asc(X1) - 55) * 16
    End If
    
    If IsNumeric(X2) Then
        Temp = Temp + Int(X2)
    Else
        Temp = Temp + (Asc(X2) - 55)
    End If
    
    ' retorno
    ConvToInt = Temp
    
End Function

