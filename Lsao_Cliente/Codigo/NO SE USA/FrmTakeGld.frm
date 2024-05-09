VERSION 5.00
Begin VB.Form FrmTransferir 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1875
   ClientLeft      =   3300
   ClientTop       =   4410
   ClientWidth     =   3480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H0000C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   285
      Left            =   1400
      TabIndex        =   1
      Top             =   840
      Width           =   1890
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   285
      Left            =   1400
      TabIndex        =   0
      Text            =   "0"
      Top             =   540
      Width           =   1890
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   1800
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   120
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "FrmTransferir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub


Private Sub Form_Load()
'Me.Picture = LoadPicture(App.Path & _
'    "\Graficos\mandaoro.jpg")
Call cargarImagenRes(Me, 135)
End Sub

Private Sub Image1_Click()
Dim a As String
a = Encriptar(Text2.Text & " " & Text1.Text)
Call SendData("/ALPETE " & a)
Unload Me
End Sub

Private Sub Image2_Click()
Unload Me
End Sub
