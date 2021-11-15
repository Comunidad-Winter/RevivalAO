VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ------------------------------------------------------------
'  Copyright ©2001 Mike G --> IvbNET.COM
'  All Rights Reserved, http://www.ivbnet.com
'  EMAIL : webmaster@ivbnet.com
' ------------------------------------------------------------
'  You are free to use this code within your own applications,
'  but you are forbidden from selling or distributing this
'  source code without prior written consent.
' ------------------------------------------------------------


Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
        "URLDownloadToFileA" (ByVal pCaller As Long, _
        ByVal szURL As String, _
        ByVal szFileName As String, _
        ByVal dwReserved As Long, _
        ByVal lpfnCB As Long) As Long

Private Declare Function InternetOpen Lib "wininet" Alias _
        "InternetOpenA" (ByVal sAgent As String, _
        ByVal lAccessType As Long, _
        ByVal sProxyName As String, _
        ByVal sProxyBypass As String, _
        ByVal lFlags As Long) As Long


Private Declare Function InternetCloseHandle Lib "wininet" _
        (ByVal hInet As Long) As Integer


Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000
Const INTERNET_OPEN_TYPE_DIRECT = 1
Const INTERNET_OPEN_TYPE_PROXY = 3
Const INTERNET_FLAG_RELOAD = &H80000000



Public Function Get_File(sURLFileName As String, _
                         sSaveFileName As String) As Boolean
    Dim lRet As Long
    On Error GoTo err_Fix
  
    lRet = InternetOpen("", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    lRet = URLDownloadToFile(0, sURLFileName, sSaveFileName, 0, 0)
    Get_File = True
    Exit Function
err_Fix:
    Debug.Print Err.LastDllError, lRet
    Err.Clear
    Get_File = False
End Function

