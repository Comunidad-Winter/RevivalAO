VERSION 5.00
Begin VB.Form frmGuildFoundation 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Creación de un Clan"
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4605
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   277
      Left            =   550
      TabIndex        =   1
      Top             =   3330
      Width           =   3495
   End
   Begin VB.TextBox txtClanName 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   285
      Left            =   540
      TabIndex        =   0
      Top             =   2370
      Width           =   3510
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   2760
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   360
      Top             =   3840
      Width           =   1455
   End
End
Attribute VB_Name = "frmGuildFoundation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub

Private Sub Form_Load()
'Valores máximos y mínimos para el ScrollBar
 ' Me.Picture = LoadPicture(App.Path & _
  '  "\Graficos\FundClan.jpg")
    Call cargarImagenRes(frmGuildFoundation, 119)
If Len(txtClanName.Text) <= 30 Then
    If Not AsciiValidos(txtClanName) Then
        MsgBox "Nombre invalido."
        Exit Sub
    End If
Else
        MsgBox "Nombre demasiado extenso."
        Exit Sub
End If



End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image2_Click()
ClanName = txtClanName
Site = Text2
Unload Me
frmGuildDetails.Show , Me
End Sub
