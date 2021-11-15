VERSION 5.00
Begin VB.Form frmSoporte 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   360
      ScrollBars      =   1  'Horizontal
      TabIndex        =   1
      Text            =   "Escribe aqui tu consulta."
      Top             =   5400
      Width           =   6855
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   4800
      MouseIcon       =   "frmSoporte.frx":0000
      MousePointer    =   99  'Custom
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   600
      MouseIcon       =   "frmSoporte.frx":0CCA
      MousePointer    =   99  'Custom
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Staff"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   0
      Top             =   5040
      Width           =   2895
   End
End
Attribute VB_Name = "frmSoporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call cargarImagenRes(Me, 132)
    Me.Top = 0
    Me.Left = 0
End Sub

Private Sub Image1_Click()
If Text1.Text <> "" Then
SendData "/SOPR " & Text1.Text
Unload frmSoporte
End If
End Sub

Private Sub Image2_Click()
Unload frmSoporte
End Sub
