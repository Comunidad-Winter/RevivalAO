VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurador RevivalAo"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3330
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones Generales"
      Height          =   1575
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3255
      Begin VB.CheckBox Check3 
         Caption         =   "Musica"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Sonidos"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Usar Memoria de Video"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "FPS"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Tag             =   "i e"
      Top             =   1560
      Width           =   3255
      Begin VB.OptionButton Option1 
         Caption         =   "72"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   555
      End
      Begin VB.OptionButton Option1 
         Caption         =   "36"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "18"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type tSetupMods
    bDinamic    As Boolean
    byMemory    As Byte
    bUseVideo   As Boolean
    bNoMusic    As Boolean
    bNoSound    As Boolean
    bFPS        As Byte
End Type

Private ClientSetup As tSetupMods

Private IsLoaded As Boolean

Private Sub Image5_Click()
End Sub

Private Sub Check1_Click()
    ClientSetup.bUseVideo = Check1.value = 1
    Call SaveClientSetup
End Sub

Private Sub Check2_Click()
    ClientSetup.bNoSound = Check2.value = 0
    Call SaveClientSetup
End Sub

Private Sub Check3_Click()
    ClientSetup.bNoMusic = Check3.value = 0
    Call SaveClientSetup
End Sub

Private Sub Command1_Click()
    Call SaveClientSetup
    Unload Me
End Sub

Private Sub Form_Load()
    Call LoadClientSetup
End Sub

Private Sub Option1_Click(Index As Integer)
    ClientSetup.bFPS = Index
    Call SaveClientSetup
End Sub

Private Sub LoadClientSetup()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'
'**************************************************************
    Dim ConfigPath As String
    ConfigPath = App.Path & "\Recursos\GameConfig.revival"
    ClientSetup.bDinamic = Val(GetVar(ConfigPath, "CONFIG", "bDinamic")) = 1
    ClientSetup.bFPS = Val(GetVar(ConfigPath, "CONFIG", "bFPS"))
    ClientSetup.bNoMusic = Val(GetVar(ConfigPath, "CONFIG", "bNoMusic")) = 1
    ClientSetup.bNoSound = Val(GetVar(ConfigPath, "CONFIG", "bNoSound")) = 1
    ClientSetup.bUseVideo = Val(GetVar(ConfigPath, "CONFIG", "bUseVideo")) = 1
    ClientSetup.byMemory = Val(GetVar(ConfigPath, "CONFIG", "byMemory"))

    Check1.value = IIf(ClientSetup.bUseVideo, vbChecked, vbUnchecked)
    Check2.value = IIf(ClientSetup.bNoMusic, vbUnchecked, vbChecked)
    Check3.value = IIf(ClientSetup.bNoSound, vbUnchecked, vbChecked)
    Option1(ClientSetup.bFPS).value = True
    
    IsLoaded = True
End Sub
Public Sub SaveClientSetup()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'
'**************************************************************
    If Not IsLoaded Then Exit Sub
    Dim ConfigPath As String
    ConfigPath = App.Path & "\Recursos\GameConfig.revival"
    With ClientSetup
        Call WriteVar(ConfigPath, "CONFIG", "bDinamic", IIf(.bDinamic, 1, 0))
        Call WriteVar(ConfigPath, "CONFIG", "bFPS", .bFPS)
        Call WriteVar(ConfigPath, "CONFIG", "bNoMusic", IIf(.bNoMusic, 1, 0))
        Call WriteVar(ConfigPath, "CONFIG", "bNoSound", IIf(.bNoSound, 1, 0))
        Call WriteVar(ConfigPath, "CONFIG", "bUseVideo", IIf(.bUseVideo, 1, 0))
        Call WriteVar(ConfigPath, "CONFIG", "byMemory", .byMemory)
    End With
End Sub
