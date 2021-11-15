Attribute VB_Name = "MD5"
Private Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal r As String)
Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal t As Long, ByVal r As String)



Public Function MD5File(f As String) As String
' compute MD5 digest on o given file, returning the result
    Dim r As String * 32
    r = Space(32)
    MDFile f, r
    MD5File = r
End Function


Function ENCRYPT(ByVal STRG As String) As String
If val(STRG) <> 5 Then
    For asd = 1 To Len(STRG)
        suma = suma + Asc(mid$(STRG, asd, 1))
    Next
    For asd = 1 To Asc(mid$(STRG, 1, 1))
        If ENCRYPT = "" Then
            ENCRYPT = MD5String(CStr(suma * 0.512))
        Else
            ENCRYPT = MD5String(ENCRYPT)
        End If
    Next


End If
End Function







