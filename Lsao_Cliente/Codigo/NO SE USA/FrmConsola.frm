VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmConsola 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Consola de Eventos"
   ClientHeight    =   1245
   ClientLeft      =   8760
   ClientTop       =   300
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleMode       =   0  'User
   ScaleWidth      =   95.735
   Begin RichTextLib.RichTextBox Consola 
      CausesValidation=   0   'False
      Height          =   1500
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2646
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"FrmConsola.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmConsola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Constantes para pasarle a la función Api SetWindowPos
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2 '

' Función Api SetWindowPos
Private Declare Function SetWindowPos _
    Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal X As Long, ByVal Y As Long, _
        ByVal cX As Long, _
        ByVal cY As Long, _
        ByVal wFlags As Long) As Long

'En el primer parámetro se le pasa el Hwnd de la ventana
'El segundo es la constante que permite hacer el OnTop
'Los parámetros que están en 0 son las coordenadas, o sea la _
 pocición, obviamente opcionales
'El último parámetro es para que al establecer el OnTop la ventana _
no se mueva de lugar y no se redimensione

Private Sub Form_Resize()
Consola.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub



Private Sub Form_Load()
'Valores máximos y mínimos para el ScrollBar

'Me.Top = 0
'Me.Left = 0
    ' Le establecemos un valor por defecto _
    a la barra apenas carga el form

   Call Aplicar_Transparencia(Me.hWnd, CByte(200))
  SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
                            SWP_NOMOVE Or SWP_NOSIZE
End Sub

