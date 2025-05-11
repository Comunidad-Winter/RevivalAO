VERSION 5.00
Begin VB.Form frmGuildSol 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Ingreso"
   ClientHeight    =   4740
   ClientLeft      =   -60
   ClientTop       =   -165
   ClientWidth     =   5295
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   1280
      Left            =   360
      MaxLength       =   400
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2340
      Width           =   4575
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   720
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3000
      Top             =   3840
      Width           =   1695
   End
End
Attribute VB_Name = "frmGuildSol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim CName As String



Public Sub RecieveSolicitud(ByVal GuildName As String)

CName = GuildName

End Sub

Private Sub Form_Load()

Call cargarImagenRes(frmGuildSol, 122)
' frmGuildSol.Picture = LoadPicture(App.Path & _
 '   "\Graficos\Soli.jpg")
End Sub

Private Sub Image1_Click()
Dim f$

f$ = "SOLICITUD" & CName
f$ = f$ & "," & Replace(Replace(Text1.Text, ",", ";"), vbCrLf, "º")

Call SendData(f$)

Unload Me
End Sub

Private Sub Image2_Click()
Unload Me
End Sub
