VERSION 5.00
Begin VB.Form frmMapa 
   BackColor       =   &H0000C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   4710
   ClientLeft      =   1080
   ClientTop       =   2250
   ClientWidth     =   5025
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&
Private Const LWA_COLORKEY = &H1&
Private Const WS_CHILD = &H40000000
Private Const GWL_HWNDPARENT = (-8)
Private Const GW_OWNER = 4

Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_SHOWWINDOW = &H40

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal CRef As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Dim PicboxOffset As POINTAPI, WindowPos As POINTAPI, resizer As Boolean
Private Sub LoadBackPicture()
    Dim winpos As POINTAPI
    Me.ScaleMode = vbPixels
    Me.BorderStyle = 0
    Me.ScaleMode = vbPixels
    Me.AutoRedraw = True
  '  Set Me.Picture = LoadPicture(Path)
    Me.AutoRedraw = False
    ScaleMode = vbPixels
    
    ClientToScreen Me.hwnd, WindowPos
    PicboxOffset.X = Me.Left
    PicboxOffset.Y = Me.Top
    SetParent Me.hwnd, 0
    SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED Or WS_CHILD
    SetLayeredWindowAttributes Me.hwnd, 0, 64, LWA_ALPHA
    SetLayeredWindowAttributes Me.hwnd, RGB(0, 0, 0), 0, LWA_COLORKEY
    Me.Refresh
End Sub
Private Sub Form_Load()
Me.Picture = LoadResPicture(101, vbResBitmap)
    Call LoadBackPicture
End Sub


