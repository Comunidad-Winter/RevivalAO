VERSION 5.00
Begin VB.Form frmGuildAdm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Lista de Clanes Registrados"
   ClientHeight    =   3015
   ClientLeft      =   -60
   ClientTop       =   -165
   ClientWidth     =   5400
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox GuildsList 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   1980
      ItemData        =   "frmGuildAdm.frx":0000
      Left            =   1830
      List            =   "frmGuildAdm.frx":0002
      TabIndex        =   0
      Top             =   630
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   240
      Top             =   600
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   240
      Top             =   2160
      Width           =   1335
   End
End
Attribute VB_Name = "frmGuildAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub Command1_Click()



End Sub


Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub

Public Sub ParseGuildList(ByVal Rdata As String)

Dim j As Integer, k As Integer
For j = 0 To guildslist.ListCount - 1
    Me.guildslist.RemoveItem 0
Next j
k = CInt(ReadField(1, Rdata, 44))

For j = 1 To k
    guildslist.AddItem ReadField(1 + j, Rdata, 44)
Next j

Me.Show vbModal, frmMain

End Sub



Private Sub Form_Load()
Call cargarImagenRes(frmGuildAdm, 116)
'GuildAdm
  'frmGuildAdm.Picture = LoadPicture(App.Path & _
   ' "\Graficos\GuildAdm.jpg")
End Sub

Private Sub Frame1_DragDrop(source As Control, X As Single, Y As Single)

End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image2_Click()
'If GuildsList.ListIndex = 0 Then Exit Sub
Call SendData("CLANDETAILS" & guildslist.List(guildslist.ListIndex))
End Sub
