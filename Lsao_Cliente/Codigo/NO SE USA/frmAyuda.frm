VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmAyuda 
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   345
   ClientTop       =   315
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmAyuda.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2775
      Left            =   2880
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   6735
      ExtentX         =   11880
      ExtentY         =   4895
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
      Location        =   ""
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   7800
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Atrás"
      Height          =   375
      Left            =   9480
      TabIndex        =   1
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   10560
      TabIndex        =   0
      Top             =   7440
      Width           =   975
   End
   Begin VB.Image Image7 
      Height          =   615
      Left            =   360
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   10560
      Top             =   8280
      Width           =   975
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   9240
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   7440
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   5880
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   3480
      Top             =   8280
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   1920
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Image eventos3 
      Height          =   360
      Left            =   11640
      Picture         =   "frmAyuda.frx":62544
      Top             =   3840
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image eventos2 
      Height          =   360
      Left            =   11640
      Picture         =   "frmAyuda.frx":181D86
      Top             =   3240
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image eventos 
      Height          =   480
      Left            =   11640
      Picture         =   "frmAyuda.frx":2A15C8
      Top             =   2640
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image comandos2 
      Height          =   360
      Left            =   11640
      Picture         =   "frmAyuda.frx":3C0E0A
      Top             =   2160
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image comandos 
      Height          =   480
      Left            =   11640
      Picture         =   "frmAyuda.frx":4E064C
      Top             =   1560
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image controles 
      Height          =   360
      Left            =   11760
      Picture         =   "frmAyuda.frx":5FFE8E
      Top             =   1080
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image entrena2 
      Height          =   360
      Left            =   11640
      Picture         =   "frmAyuda.frx":71F6D0
      Top             =   600
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image entrena 
      Height          =   360
      Left            =   11640
      Picture         =   "frmAyuda.frx":83EF12
      Top             =   120
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image balance 
      Height          =   7680
      Left            =   248
      Picture         =   "frmAyuda.frx":95E754
      Top             =   200
      Visible         =   0   'False
      Width           =   11490
   End
End
Attribute VB_Name = "frmAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If frmAyuda.entrena.Visible = True Then
frmAyuda.entrena.Visible = False
frmAyuda.entrena2.Visible = True
End If
If frmAyuda.comandos.Visible = True Then
frmAyuda.comandos.Visible = False
frmAyuda.comandos2.Visible = True
End If
If frmAyuda.eventos.Visible = True Then
frmAyuda.eventos.Visible = False
frmAyuda.eventos2.Visible = True
Exit Sub
End If
If frmAyuda.eventos2.Visible = True Then
frmAyuda.eventos2.Visible = False
frmAyuda.eventos3.Visible = True
End If
End Sub

Private Sub Command2_Click()
If frmAyuda.entrena2.Visible = True Then
frmAyuda.entrena.Visible = True
frmAyuda.entrena2.Visible = False
End If
If frmAyuda.comandos2.Visible = True Then
frmAyuda.comandos2.Visible = False
frmAyuda.comandos.Visible = True
End If
If frmAyuda.eventos2.Visible = True Then
frmAyuda.eventos2.Visible = False
frmAyuda.eventos.Visible = True
Exit Sub
End If
If frmAyuda.eventos3.Visible = True Then
frmAyuda.eventos3.Visible = False
frmAyuda.eventos2.Visible = True
End If
End Sub

Private Sub Form_Load()
Me.Left = 0
   Me.Top = 0
frmAyuda.controles.Width = 11490
frmAyuda.controles.Top = 200
frmAyuda.controles.Height = 7680
frmAyuda.controles.Left = 248

frmAyuda.entrena.Width = 11490
frmAyuda.entrena.Top = 200
frmAyuda.entrena.Height = 7680
frmAyuda.entrena.Left = 248

frmAyuda.entrena2.Width = 11490
frmAyuda.entrena2.Top = 200
frmAyuda.entrena2.Height = 7680
frmAyuda.entrena2.Left = 248

frmAyuda.comandos.Width = 11490
frmAyuda.comandos.Top = 200
frmAyuda.comandos.Height = 7680
frmAyuda.comandos.Left = 248

frmAyuda.comandos2.Width = 11490
frmAyuda.comandos2.Top = 200
frmAyuda.comandos2.Height = 7680
frmAyuda.comandos2.Left = 248

frmAyuda.eventos.Width = 11490
frmAyuda.eventos.Top = 200
frmAyuda.eventos.Height = 7680
frmAyuda.eventos.Left = 248

frmAyuda.eventos2.Width = 11490
frmAyuda.eventos2.Top = 200
frmAyuda.eventos2.Height = 7680
frmAyuda.eventos2.Left = 248

frmAyuda.eventos3.Width = 11490
frmAyuda.eventos3.Top = 200
frmAyuda.eventos3.Height = 7680
frmAyuda.eventos3.Left = 248

frmAyuda.balance.Width = 11490
frmAyuda.balance.Top = 200
frmAyuda.balance.Height = 7680
frmAyuda.balance.Left = 248
  HScroll1.max = 255
    HScroll1.min = 50

    ' Le establecemos un valor por defecto _
    a la barra apenas carga el form

    HScroll1.value = 210


End Sub
Private Sub HScroll1_Change()

    'Llamamos a la función pasándole el handle del form _
    y el valor de la transparencia, que es el de la barra

    Call Aplicar_Transparencia(Me.hWnd, CByte(HScroll1.value))

End Sub

Private Sub Image1_Click()
frmAyuda.controles.Visible = False
frmAyuda.entrena.Visible = False
frmAyuda.entrena2.Visible = False
frmAyuda.comandos.Visible = False
frmAyuda.comandos2.Visible = False
frmAyuda.eventos.Visible = False
frmAyuda.eventos2.Visible = False
frmAyuda.eventos3.Visible = False
frmAyuda.balance.Visible = True
frmAyuda.WebBrowser1.Visible = False
End Sub

Private Sub Image2_Click()
frmAyuda.controles.Visible = False
frmAyuda.entrena.Visible = True
frmAyuda.entrena2.Visible = False
frmAyuda.comandos.Visible = False
frmAyuda.comandos2.Visible = False
frmAyuda.eventos.Visible = False
frmAyuda.eventos2.Visible = False
frmAyuda.eventos3.Visible = False
frmAyuda.balance.Visible = False
frmAyuda.WebBrowser1.Visible = False
End Sub

Private Sub Image3_Click()
frmAyuda.controles.Visible = True
frmAyuda.entrena.Visible = False
frmAyuda.entrena2.Visible = False
frmAyuda.comandos.Visible = False
frmAyuda.comandos2.Visible = False
frmAyuda.eventos.Visible = False
frmAyuda.eventos2.Visible = False
frmAyuda.eventos3.Visible = False
frmAyuda.balance.Visible = False
frmAyuda.WebBrowser1.Visible = False
End Sub

Private Sub Image4_Click()
frmAyuda.controles.Visible = False
frmAyuda.entrena.Visible = False
frmAyuda.entrena2.Visible = False
frmAyuda.comandos.Visible = True
frmAyuda.comandos2.Visible = False
frmAyuda.eventos.Visible = False
frmAyuda.eventos2.Visible = False
frmAyuda.eventos3.Visible = False
frmAyuda.balance.Visible = False
frmAyuda.WebBrowser1.Visible = False
End Sub

Private Sub Image5_Click()
frmAyuda.controles.Visible = False
frmAyuda.entrena.Visible = False
frmAyuda.entrena2.Visible = False
frmAyuda.comandos.Visible = False
frmAyuda.comandos2.Visible = False
frmAyuda.eventos.Visible = True
frmAyuda.eventos2.Visible = False
frmAyuda.eventos3.Visible = False
frmAyuda.balance.Visible = False
frmAyuda.WebBrowser1.Visible = False
End Sub

Private Sub Image6_Click()
Unload Me
End Sub

Private Sub Image7_Click()
frmAyuda.controles.Visible = False
frmAyuda.entrena.Visible = False
frmAyuda.entrena2.Visible = False
frmAyuda.comandos.Visible = False
frmAyuda.comandos2.Visible = False
frmAyuda.eventos.Visible = False
frmAyuda.eventos2.Visible = False
frmAyuda.eventos3.Visible = False
frmAyuda.balance.Visible = False
frmAyuda.WebBrowser1.Visible = True
frmAyuda.WebBrowser1.Navigate ("http://www.symxsoft.net/revival/asdf.php")
End Sub
