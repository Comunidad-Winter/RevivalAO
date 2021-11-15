VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   662
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox MP3Files 
      Height          =   480
      Left            =   180
      Pattern         =   "*.mp3"
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox LOGO 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   0
      Picture         =   "frmCargando.frx":0000
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   663
      TabIndex        =   0
      Top             =   0
      Width           =   9945
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   6840
         TabIndex        =   11
         Top             =   6720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Salir"
         Height          =   495
         Left            =   4560
         TabIndex        =   8
         Top             =   6360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtTo 
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   5880
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.TextBox txtFrom 
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Text            =   "http://www.symxsoft.net/zeusao/hds.txt"
         Top             =   5400
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.CommandButton cmdDownload 
         Caption         =   "&Descargar"
         Height          =   495
         Left            =   3240
         TabIndex        =   5
         Top             =   6360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   960
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   5880
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin RichTextLib.RichTextBox Status 
         Height          =   2775
         Left            =   2160
         TabIndex        =   2
         Top             =   2520
         Visible         =   0   'False
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   4895
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"frmCargando.frx":54D1B
      End
      Begin VB.Label Label1 
         Caption         =   "A:"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   10
         Top             =   5880
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   9
         Top             =   5400
         Visible         =   0   'False
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private Sub Form_Load()



' LOGO.Picture = LoadPicture(App.Path & "\Graficos\Cargando.jpg")
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

