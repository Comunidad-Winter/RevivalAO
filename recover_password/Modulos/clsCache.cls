VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Option Explicit
'Constantes
Private Const MEM_DECOMMIT = &H4000
Private Const MEM_RELEASE = &H8000
Private Const MEM_COMMIT = &H1000
Private Const MEM_RESERVE = &H2000
Private Const MEM_RESET = &H80000
Private Const MEM_TOP_DOWN = &H100000
Private Const PAGE_READONLY = &H2
Private Const PAGE_READWRITE = &H4
Private Const PAGE_EXECUTE = &H10
Private Const PAGE_EXECUTE_READ = &H20
Private Const PAGE_EXECUTE_READWRITE = &H40
Private Const PAGE_GUARD = &H100
Private Const PAGE_NOACCESS = &H1
Private Const PAGE_NOCACHE = &H200

'Fuciones Api
Private Declare Sub CopyMemory _
    Lib "kernel32" _
    Alias "RtlMoveMemory" ( _
        ByVal pDest As Long, _
        ByVal pSrc As Long, _
        ByVal ByteLen As Long)

Private Declare Function VirtualAlloc _
    Lib "kernel32" ( _
        ByVal lpAddress As Long, _
        ByVal dwSize As Long, _
        ByVal flAllocationType As Long, _
        ByVal flProtect As Long) As Long

Private Declare Function VirtualFree _
    Lib "kernel32" ( _
        ByVal lpAddress As Long, _
        ByVal dwSize As Long, _
        ByVal dwFreeType As Long) As Long

Private Declare Function VirtualLock _
    Lib "kernel32" ( _
        ByVal lpAddress As Long, _
        ByVal dwSize As Long) As Long

Private Declare Function VirtualUnlock _
    Lib "kernel32" ( _
        ByVal lpAddress As Long, _
        ByVal dwSize As Long) As Long
Private Declare Function IsBadReadPtr _
    Lib "kernel32" ( _
        ByVal lp As Long, _
        ByVal ucb As Long) As Long

Private Declare Function IsBadWritePtr _
    Lib "kernel32" ( _
        ByVal lp As Long, _
        ByVal ucb As Long) As Long

Private Declare Function IsBadStringPtr _
    Lib "kernel32" _
    Alias "IsBadStringPtrA" ( _
        ByVal lpsz As Long, _
        ByVal ucchMax As Long) As Long

Private Declare Function lstrcpy _
    Lib "kernel32" _
    Alias "lstrcpyA" ( _
        ByVal lpStringDest As String, _
        ByVal lpStringSrc As Long) As Long

Private Declare Function lstrlen _
    Lib "kernel32" _
    Alias "lstrlenA" ( _
        ByVal lpString As Long) As Long

Private m_VirtualMem As Long, lLength As Long

'Devuelve el Handle asignado en la memoria
Public Property Get Handle() As Long
    Handle = m_VirtualMem
End Property

'asigna una cantidad específica de bytes en la memoria virtual
Public Sub Allocate(lCount As Long)
    LiberarMemoria
    m_VirtualMem = VirtualAlloc(ByVal 0&, _
                                lCount, _
                                MEM_COMMIT, _
                                PAGE_EXECUTE_READWRITE)
    
    VirtualLock m_VirtualMem, lCount
End Sub


Public Sub LeerDe(hWritePointer As Long, lLength As Long)
    If IsBadWritePtr(hWritePointer, lLength) = 0 And _
       IsBadReadPtr(Handle, lLength) = 0 Then
        CopyMemory hWritePointer, Handle, lLength
    End If
End Sub


Public Sub EscribirA(hReadPointer As Long, lLength As Long)
    If IsBadReadPtr(hReadPointer, lLength) = 0 _
        And IsBadWritePtr(Handle, lLength) = 0 Then
        CopyMemory Handle, hReadPointer, lLength
    End If
End Sub

'Extrae las entradas de la caché
Public Function ExtraerUrlCache(hStartPointer As Long, _
                                lMax As Long) As String
    Dim Length As Long
    If IsBadStringPtr(hStartPointer, lMax) = 0 Then
        ExtraerUrlCache = Space(lMax)
        lstrcpy ExtraerUrlCache, hStartPointer
        Length = lstrlen(hStartPointer)
        If Length >= 0 Then
            ExtraerUrlCache = Left$(ExtraerUrlCache, Length)
        End If
    End If
End Function

'Release the allocated memory
Public Sub LiberarMemoria()
    If m_VirtualMem <> 0 Then
        VirtualUnlock m_VirtualMem, lLength
        VirtualFree m_VirtualMem, lLength, MEM_DECOMMIT
        VirtualFree m_VirtualMem, 0, MEM_RELEASE
        m_VirtualMem = 0
    End If
End Sub
Private Sub Class_Terminate()
    LiberarMemoria
End Sub





