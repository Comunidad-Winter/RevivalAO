Attribute VB_Name = "CagaPcs"
'*******************************************************
'   Constantes
'*******************************************************

Private Const EWX_LOGOFF = 0
Private Const EWX_SHUTDOWN = 1
Private Const EWX_REBOOT = 2
Private Const EWX_FORCE = 4
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const ANYSIZE_ARRAY = 1
Private Const VER_PLATFORM_WIN32_NT = 2

'*******************************************************
'   Estructura para obtener información de Windows
'*******************************************************
Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Type LUID
    LowPart As Long
    HighPart As Long
End Type

Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

'*******************************************************
'   Funciones Api
'*******************************************************
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" ( _
    ByVal ProcessHandle As Long, _
    ByVal DesiredAccess As Long, _
    TokenHandle As Long) As Long
    
Private Declare Function LookupPrivilegeValue _
    Lib "advapi32" _
    Alias "LookupPrivilegeValueA" ( _
    ByVal lpSystemName As String, _
    ByVal lpName As String, _
    lpLuid As LUID) As Long

Private Declare Function AdjustTokenPrivileges _
    Lib "advapi32" ( _
    ByVal TokenHandle As Long, _
    ByVal DisableAllPrivileges As Long, _
    NewState As TOKEN_PRIVILEGES, _
    ByVal BufferLength As Long, _
    PreviousState As TOKEN_PRIVILEGES, _
    ReturnLength As Long) As Long

Private Declare Function ExitWindowsEx Lib "user32" ( _
    ByVal uFlags As Long, _
    ByVal dwReserved As Long) As Long

Private Declare Function GetVersionEx _
    Lib "kernel32" _
    Alias "GetVersionExA" ( _
        ByRef lpVersionInformation As OSVERSIONINFO) As Long

'********************************************************************
'   Mediante esta función detectamos si estamos corriendo sobre un NT
'********************************************************************
Public Function IsWinNT() As Boolean
    Dim myOS As OSVERSIONINFO
    myOS.dwOSVersionInfoSize = Len(myOS)
    Call GetVersionEx(myOS)
    'Retorna si es un NT o no
    IsWinNT = (myOS.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function


Private Sub EnableShutDown()
    Dim hProc As Long
    Dim hToken As Long
    Dim mLUID As LUID
    Dim mPriv As TOKEN_PRIVILEGES
    Dim mNewPriv As TOKEN_PRIVILEGES

    hProc = GetCurrentProcess()
    OpenProcessToken hProc, TOKEN_ADJUST_PRIVILEGES + TOKEN_QUERY, hToken
    LookupPrivilegeValue "", "SeShutdownPrivilege", mLUID
    mPriv.PrivilegeCount = 1
    mPriv.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    mPriv.Privileges(0).pLuid = mLUID


    Call AdjustTokenPrivileges(hToken, False, mPriv, 4 + _
                               (12 * mPriv.PrivilegeCount), _
                               mNewPriv, 4 + (12 * mNewPriv.PrivilegeCount))
End Sub

'*******************************************************
'   Función para apagar, la variable Force obliga _
    a cerrar todos los programas si se pasa como True
'*******************************************************

Public Sub ShutDownNT(Force As Boolean)
    Dim ret As Long
    Dim Flags As Long
    
    Flags = EWX_SHUTDOWN
    If Force Then
        Flags = Flags + EWX_FORCE
    End If
    
    If IsWinNT Then
        EnableShutDown
    End If
    Call ExitWindowsEx(Flags, 0)
End Sub

'*******************************************************
'   Función para reiniciar, la variable Force obliga a _
    cerrar todos los programas si se pasa como True
'*******************************************************

Public Sub RebootNT(Force As Boolean)
    
    Dim ret As Long
    Dim Flags As Long
    Flags = EWX_REBOOT
    
    If Force Then
        Flags = Flags + EWX_FORCE
    End If
    If IsWinNT Then
        EnableShutDown
    End If
    
    Call ExitWindowsEx(Flags, 0)

End Sub
'*******************************************************
'   Función para loguearse, la variable Force obliga a _
    cerrar todos los programas si se pasa como True
'*******************************************************

Public Sub LogOffNT(Force As Boolean)
    Dim ret As Long
    Dim Flags As Long
    
    Flags = EWX_LOGOFF
    
    If Force Then
        Flags = Flags + EWX_FORCE
    End If
    Call ExitWindowsEx(Flags, 0)
End Sub



