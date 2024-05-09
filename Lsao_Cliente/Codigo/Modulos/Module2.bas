Attribute VB_Name = "VERSERIALHD"
Sub disco()
Dim fso As New Scripting.FileSystemObject
Dim dr As Scripting.Drive
Set dr = fso.GetDrive("c:")
hd = dr.SerialNumber
End Sub
