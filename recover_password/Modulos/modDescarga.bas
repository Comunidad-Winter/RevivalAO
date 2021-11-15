Attribute VB_Name = "modDescarga"
Public Sub Descargar_Users()
 On Error Resume Next
  Dim Obj As clsDownload
  Set Obj = New clsDownload
  Dim bRet As Boolean
  
 
  Form1.Refresh
  
     Screen.MousePointer = vbHourglass
       bRet = Obj.Get_File(urlUsers, pathUsers)
        If bRet = False Then MsgBox "Error downloading!"
          Screen.MousePointer = vbDefault
     Set Obj = Nothing
End Sub
Public Sub Descargar_Mails()
 On Error Resume Next
  Dim Obj As clsDownload
  Set Obj = New clsDownload
  Dim bRet As Boolean
  
 
  Form1.Refresh
  
     Screen.MousePointer = vbHourglass
       bRet = Obj.Get_File(urlMails, pathMails)
        If bRet = False Then MsgBox "Error downloading!"
          Screen.MousePointer = vbDefault
     Set Obj = Nothing
End Sub

