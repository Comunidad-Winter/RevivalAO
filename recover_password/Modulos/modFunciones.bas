Attribute VB_Name = "modFunciones"
Sub Main()
PuedeRecuperar = True
'Creamos una instancia de la clase Class1
Set clsCache = New clsCache
urlUsers = "http://www.revivalao.com.ar/recu/users.txt"
urlMails = "http://www.revivalao.com.ar/recu//mails.txt"

 'path donde guardar usuarios y mails
  pathUsers = App.Path & "\Recursos\users.txt"
  pathMails = App.Path & "\Recursos\mails.txt"
' path donde estan los charfiles
  pathChar = App.Path & "\Charfile\"
  'user y pass gmail
  GmailUser = "saturos@revivalao.com.ar"
  GmailPass = "Gym123"
  ' url borrar users y mails
  urlBorrar = "http://www.revivalao.com.ar/recu/borrar.php?index=borrar"
  Form1.Show
End Sub

Public Sub Log(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\Logs\logs.txt" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub
Public Sub Cargar_Codeseg()
On Error GoTo error
Dim linea As String
   
   Open pathCodeseg For Input As #1

   While Not EOF(1)
     'Lee la linea del archivo
      Line Input #1, linea
      'La carga en el textbox
     Codeseg = Split(linea, ",")
   Wend
   'Cierra el archivo abierto
   Close
error:
   If Err.Number = 53 Then
   Call Log("No hay personajes que recuperar")
   ' PuedeRecuperar = False
   End If
End Sub

Public Sub Cargar_Users()
On Error GoTo error
Dim linea As String
   
   Open pathUsers For Input As #1

   While Not EOF(1)
     'Lee la linea del archivo
      Line Input #1, linea
      'La carga en el textbox
     Users = Split(linea, ",")
   Wend
   'Cierra el archivo abierto
   Close
error:
   If Err.Number = 53 Then
   Call Log("No hay personajes que recuperar")
   PuedeRecuperar = False
   End If
End Sub
Public Sub Cargar_Mails()
On Error GoTo error
Dim linea As String
   
   Open pathMails For Input As #1

   While Not EOF(1)
     'Lee la linea del archivo
      Line Input #1, linea
      'La carga en el textbox
     Mails = Split(linea, ",")
   Wend
   'Cierra el archivo abierto
   Close
error:
   If Err.Number = 53 Then
   Call Log("No hay personajes que recuperar")
    'PuedeRecuperar = False
   End If
End Sub
Public Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

RandomNumber = Int(Rnd * (UpperBound - LowerBound + 1)) + LowerBound

End Function
Public Sub Borrar_Archivos()
On Error Resume Next
Kill pathUsers
Kill pathMails
'Kill pathCodeseg
Form1.Inet1.URL = urlBorrar
Form1.Inet1.OpenURL
End Sub
Function Existe(sArchivo As String) As Boolean
    Existe = Len(Dir$(sArchivo))
End Function
Sub WriteVar(File As String, Main As String, Var As String, Value As Variant)
writeprivateprofilestring Main, Var, Value, File
End Sub
Function GetVar(File As String, Main As String, Var As String) As String
Dim L As Integer
Dim Char As String
Dim sSpaces As String
Dim szReturn As String
szReturn = ""
sSpaces = Space(5000)
getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File
GetVar = RTrim(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Public Function MD5String(p As String) As String
' compute MD5 digest on a given string, returning the result
    Dim r As String * 32, t As Long
    r = Space(32)
    t = Len(p)
    MDStringFix p, t, r
    MD5String = r
End Function

Public Function MD5File(f As String) As String
' compute MD5 digest on o given file, returning the result
    Dim r As String * 32
    r = Space(32)
    MDFile f, r
    MD5File = r
End Function
