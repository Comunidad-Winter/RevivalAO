VERSION 5.00
Begin VB.Form frmGuildLeader 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Administración del Clan"
   ClientHeight    =   6540
   ClientLeft      =   -60
   ClientTop       =   -165
   ClientWidth     =   6270
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
   ScaleHeight     =   6540
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox solicitudes 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000005&
      Height          =   810
      ItemData        =   "frmGuildLeader.frx":0000
      Left            =   330
      List            =   "frmGuildLeader.frx":0002
      TabIndex        =   5
      Top             =   4620
      Width           =   2655
   End
   Begin VB.TextBox txtguildnews 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000004&
      Height          =   735
      Left            =   330
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2940
      Width           =   5535
   End
   Begin VB.ListBox members 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000004&
      Height          =   1395
      ItemData        =   "frmGuildLeader.frx":0004
      Left            =   3210
      List            =   "frmGuildLeader.frx":0006
      TabIndex        =   3
      Top             =   660
      Width           =   2655
   End
   Begin VB.ListBox guildslist 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000004&
      Height          =   1395
      ItemData        =   "frmGuildLeader.frx":0008
      Left            =   330
      List            =   "frmGuildLeader.frx":000A
      TabIndex        =   2
      Top             =   660
      Width           =   2655
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Propuestas de alianzas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      MouseIcon       =   "frmGuildLeader.frx":000C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7740
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Propuestas de paz"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      MouseIcon       =   "frmGuildLeader.frx":015E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7230
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image Image8 
      Height          =   495
      Left            =   3360
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Image Image7 
      Height          =   375
      Left            =   3360
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   3360
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   3360
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Label Miembros 
      BackColor       =   &H00404040&
      Caption         =   "El clan cuenta con x miembros"
      ForeColor       =   &H8000000E&
      Height          =   247
      Left            =   330
      TabIndex        =   6
      Top             =   6000
      Width           =   2650
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   480
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   360
      Top             =   3720
      Width           =   5415
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   3360
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   480
      Top             =   2160
      Width           =   2415
   End
End
Attribute VB_Name = "frmGuildLeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub cmdElecciones_Click()
 
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()



End Sub

Private Sub Command3_Click()



End Sub

Private Sub Command4_Click()


End Sub

Private Sub Command5_Click()


End Sub

Private Sub Command6_Click()

End Sub

Private Sub Command7_Click()
Call SendData("ENVPROPP")
End Sub
Private Sub Command9_Click()
Call SendData("ENVALPRO")
End Sub


Private Sub Command8_Click()

End Sub


Public Sub ParseLeaderInfo(ByVal data As String)

If Me.Visible Then Exit Sub

Dim r%, T%

r% = Val(ReadField(1, data, Asc("¬")))

For T% = 1 To r%
    guildslist.AddItem ReadField(1 + T%, data, Asc("¬"))
Next T%

r% = Val(ReadField(T% + 1, data, Asc("¬")))
Miembros.Caption = "El clan cuenta con " & r% & " miembros."

Dim k%

For k% = 1 To r%
    members.AddItem ReadField(T% + 1 + k%, data, Asc("¬"))
Next k%

txtguildnews = Replace(ReadField(T% + k% + 1, data, Asc("¬")), "º", vbCrLf)

T% = T% + k% + 2

r% = Val(ReadField(T%, data, Asc("¬")))

For k% = 1 To r%
    solicitudes.AddItem ReadField(T% + k%, data, Asc("¬"))
Next k%

Me.Show , frmMain

End Sub


Private Sub Form_Deactivate()
'Me.SetFocus
End Sub

Private Sub Form_Load()
'
 'frmGuildLeader.Picture = LoadPicture(App.Path & _
  '  "\Graficos\ClanLider.jpg")
  
Call cargarImagenRes(frmGuildLeader, 120)
End Sub

Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Image1_Click()
frmGuildBrief.EsLeader = True
Call SendData("CLANDETAILS" & guildslist.List(guildslist.ListIndex))

'Unload Me

End Sub

Private Sub Image2_Click()
frmCharInfo.frmmiembros = True
Call SendData("1HRINFO<" & members.List(members.ListIndex))

'Unload Me
End Sub

Private Sub Image3_Click()
Dim k$

k$ = Replace(txtguildnews, vbCrLf, "º")

Call SendData("ACTGNEWS" & k$)
End Sub

Private Sub Image4_Click()

frmCharInfo.frmsolicitudes = True
Call SendData("1HRINFO<" & solicitudes.List(solicitudes.ListIndex))

'Unload Me

End Sub

Private Sub Image5_Click()
   Call SendData("ABREELEC")
    Unload Me
End Sub

Private Sub Image6_Click()
Call frmGuildDetails.Show(vbModal, frmGuildLeader)

'Unload Me

End Sub

Private Sub Image7_Click()
Call frmGuildURL.Show(vbModeless, frmGuildLeader)
'Unload Me
End Sub

Private Sub Image8_Click()
Unload Me
frmMain.SetFocus
End Sub
