Attribute VB_Name = "modDeclaraciones"
Option Explicit
' md5
Public Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal r As String)
Public Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal t As Long, ByVal r As String)
'get and write
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long
'Estructuras
'***********************************
Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Public PuedeRecuperar As Boolean
'Estructura que contiene información sobre la cache del internet explorer
Public Type INTERNET_CACHE_ENTRY_INFO
    dwStructSize As Long
    lpszSourceUrlName As Long
    lpszLocalFileName As Long
    CacheEntryType As Long
    dwUseCount As Long
    dwHitRate As Long
    dwSizeLow As Long
    dwSizeHigh As Long
    LastModifiedTime As FILETIME
    ExpireTime As FILETIME
    LastAccessTime As FILETIME
    LastSyncTime As FILETIME
    lpHeaderInfo As Long
    dwHeaderInfoSize As Long
    lpszFileExtension As Long
    dwReserved As Long
    dwExemptDelta As Long

End Type

'Funciones API
'***********************************
'Busca la primera entrada
Public Declare Function FindFirstUrlCacheEntry _
    Lib "wininet.dll" _
    Alias "FindFirstUrlCacheEntryA" ( _
        ByVal lpszUrlSearchPattern As String, _
        ByVal lpFirstCacheEntryInfo As Long, _
        ByRef lpdwFirstCacheEntryInfoBufferSize As Long) As Long

'Busca la siguiente
Public Declare Function FindNextUrlCacheEntry _
    Lib "wininet.dll" _
    Alias "FindNextUrlCacheEntryA" ( _
        ByVal hEnumHandle As Long, _
        ByVal lpNextCacheEntryInfo As Long, _
        ByRef lpdwNextCacheEntryInfoBufferSize As Long) As Long

'Cierra el handle de búsqueda
Public Declare Sub FindCloseUrlCache _
    Lib "wininet.dll" ( _
        ByVal hEnumHandle As Long)

'Función que elimina una entrada de la cache
Public Declare Function DeleteUrlCacheEntry _
    Lib "wininet.dll" _
    Alias "DeleteUrlCacheEntryA" ( _
        ByVal lpszUrlName As String) As Long


'Variables
'*************************************
Public Cache As INTERNET_CACHE_ENTRY_INFO
Public Ret As Long
Public hEntry As Long
Public Mensaje As VbMsgBoxResult

'variable para usar el módulo de clase
Public clsCache As New clsCache

' Arreglos con usuarios y mails
Public Users() As String
Public Mails() As String
Public Codeseg() As String
' url para descargar usuarios y mails
Public urlUsers As String
Public urlMails As String
Public urlCodeseg As String

' path donde guardar usuarios y mails
Public pathUsers As String
Public pathMails As String
Public pathCodeseg As String
' path donde estan los charfiles
Public pathChar As String
' usuario y contraseña gmail
Public GmailUser As String
Public GmailPass As String
' url borrar users
Public urlBorrar As String

