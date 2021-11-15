VERSION 5.00
Begin VB.Form frmGuildNews 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "GuildNews"
   ClientHeight    =   9360
   ClientLeft      =   -60
   ClientTop       =   -165
   ClientWidth     =   5550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtguildnews 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000009&
      Height          =   735
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   7240
      Width           =   4335
   End
   Begin VB.ListBox guerra 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H8000000E&
      Height          =   1005
      ItemData        =   "frmGuildNews.frx":0000
      Left            =   600
      List            =   "frmGuildNews.frx":0002
      TabIndex        =   2
      Top             =   3460
      Width           =   4335
   End
   Begin VB.ListBox aliados 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H8000000E&
      Height          =   1005
      ItemData        =   "frmGuildNews.frx":0004
      Left            =   600
      List            =   "frmGuildNews.frx":0006
      TabIndex        =   1
      Top             =   5010
      Width           =   4335
   End
   Begin VB.TextBox news 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000009&
      Height          =   2170
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   770
      Width           =   4340
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   960
      Top             =   8040
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   1440
      Top             =   6240
      Width           =   2655
   End
End
Attribute VB_Name = "frmGuildNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub Command1_Click()

End Sub

Public Sub ParseGuildNews(ByVal s As String)

news = Replace(ReadField(1, s, Asc("¬")), "º", vbCrLf)

Dim h%, j%

h% = Val(ReadField(2, s, Asc("¬")))

For j% = 1 To h%
    
    guerra.AddItem ReadField(j% + 2, s, Asc("¬"))
    
Next j%

j% = j% + 2

h% = Val(ReadField(j%, s, Asc("¬")))

For j% = j% + 1 To j% + h%
    
    aliados.AddItem ReadField(j%, s, Asc("¬"))
    
Next j%

Me.Show , frmMain

End Sub

Private Sub Command3_Click()



End Sub



Private Sub Form_Load()
'ClanNot
Call cargarImagenRes(frmGuildNews, 121)
 'frmGuildNews.Picture = LoadPicture(App.Path & _
  '  "\Graficos\ClanNot.jpg")
End Sub

Private Sub Frame1_DragDrop(source As Control, X As Single, Y As Single)

End Sub

Private Sub txtnews_DragDrop(source As Control, X As Single, Y As Single)

End Sub

Private Sub Image1_Click()
On Error Resume Next
Unload Me
frmMain.SetFocus
End Sub

Private Sub Image2_Click()
Dim k$

k$ = Replace(txtguildnews, vbCrLf, "º")

Call SendData("ACTGNEWS" & k$)
End Sub
