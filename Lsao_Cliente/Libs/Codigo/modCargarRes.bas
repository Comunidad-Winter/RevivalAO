Attribute VB_Name = "modCargarRes"
Option Explicit
Public Sub cargarImagenRes(ByVal frm As Object, ByVal numero As Integer)
On Error GoTo err
Dim sData As String
sData = StrConv(LoadResData(numero, "CUSTOM"), vbUnicode)
Open App.Path & "\..\Recursos\interfaz.jpg" For Binary As #1
Put #1, , sData
Close
frm.Picture = LoadPicture(App.Path & "\..\Recursos\interfaz.jpg")
Kill App.Path & "\..\Recursos\interfaz.jpg"
err:

End Sub

