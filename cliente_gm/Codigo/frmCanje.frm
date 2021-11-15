VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCanje 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   480
      ScaleHeight     =   510
      ScaleWidth      =   540
      TabIndex        =   1
      Top             =   720
      Width           =   570
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3910
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6906
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Golpe/Def Min"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Golpe/Def Max"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   2575
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "grhindex"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "numero"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3360
      TabIndex        =   2
      Top             =   5500
      Width           =   90
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   1560
      Top             =   720
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3840
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmCanje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call cargarImagenRes(frmCanje, 105)
'Me.Picture = LoadPicture(App.Path & "\Graficos\Canje.jpg")
End Sub

Private Sub Image1_Click()
Unload Me
End Sub
Private Sub Image2_Click()
Call SendData("INTE" & ListView1.SelectedItem.index)
End Sub
Private Sub ListView1_Click()
Call DibujaGrh(ListView1.ListItems.Item(ListView1.SelectedItem.index).SubItems(4))
Picture1.Refresh
End Sub
Public Sub DibujaGrh(Grh As Integer)
Dim SR As RECT, dr As RECT

SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.Bottom = 32

dr.Left = 0
dr.Top = 0
dr.Right = 32
dr.Bottom = 32

Call DrawGrhtoHdc(Picture1.hwnd, Picture1.hdc, Grh, SR, dr)

End Sub
