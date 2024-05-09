Attribute VB_Name = "Mod_ErrorLOG"



Option Explicit

Public Sub LogError(Desc As String)
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\errores.log" For Append As #nfile
Print #nfile, Desc
Close #nfile
End Sub

Public Sub LogCustom(Desc As String)
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\custom.log" For Append As #nfile
Print #nfile, Now & " " & Desc
Close #nfile
End Sub


