VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmRecuPass 
   BorderStyle     =   0  'None
   ClientHeight    =   4440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8040
   ControlBox      =   0   'False
   Icon            =   "frmRecuPass.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmRecuPass.frx":0CCA
   MousePointer    =   99  'Custom
   ScaleHeight     =   4440
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser asdf 
      Height          =   1935
      Left            =   1200
      TabIndex        =   0
      Top             =   1680
      Width           =   5655
      ExtentX         =   9975
      ExtentY         =   3413
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2880
      MouseIcon       =   "frmRecuPass.frx":1994
      MousePointer    =   99  'Custom
      Top             =   3840
      Width           =   2175
   End
End
Attribute VB_Name = "frmRecuPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Call cargarImagenRes(Me, 130)
asdf.Navigate ("http://www.revivalao.com.ar/recu/recuperar.php")
End Sub

Private Sub Image1_Click()
Unload frmRecuPass
End Sub
