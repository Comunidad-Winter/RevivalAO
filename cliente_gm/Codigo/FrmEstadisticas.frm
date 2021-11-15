VERSION 5.00
Begin VB.Form frmEstadisticas 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Estadisticas"
   ClientHeight    =   7050
   ClientLeft      =   -60
   ClientTop       =   -165
   ClientWidth     =   7305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmEstadisticas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2640
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   2280
      TabIndex        =   38
      Top             =   6000
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   1080
      TabIndex        =   37
      Top             =   5880
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   2040
      TabIndex        =   36
      Top             =   5640
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   2520
      TabIndex        =   33
      Top             =   4920
      Width           =   675
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   5040
      TabIndex        =   32
      Top             =   5880
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   4920
      TabIndex        =   31
      Top             =   5640
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   6000
      TabIndex        =   30
      Top             =   5400
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   5640
      TabIndex        =   29
      Top             =   5160
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   4800
      TabIndex        =   28
      Top             =   4920
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   4680
      TabIndex        =   27
      Top             =   4680
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   4920
      TabIndex        =   26
      Top             =   4440
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   4680
      TabIndex        =   25
      Top             =   4200
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   4440
      TabIndex        =   24
      Top             =   3960
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   5640
      TabIndex        =   23
      Top             =   6600
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   1320
      TabIndex        =   22
      Top             =   3960
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   1320
      TabIndex        =   21
      Top             =   3720
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   1320
      TabIndex        =   20
      Top             =   3480
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   1440
      TabIndex        =   19
      Top             =   3240
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1440
      TabIndex        =   18
      Top             =   3000
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1440
      TabIndex        =   17
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   5880
      TabIndex        =   16
      Top             =   3720
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   4800
      TabIndex        =   15
      Top             =   3480
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   5040
      TabIndex        =   14
      Top             =   3300
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   5160
      TabIndex        =   13
      Top             =   3090
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   4800
      TabIndex        =   12
      Top             =   2860
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   4800
      TabIndex        =   11
      Top             =   2620
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   4560
      TabIndex        =   10
      Top             =   2390
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   5760
      TabIndex        =   9
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   5880
      TabIndex        =   8
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   4440
      TabIndex        =   7
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   4440
      TabIndex        =   6
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   4440
      TabIndex        =   5
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   1980
      TabIndex        =   4
      Top             =   2040
      Width           =   390
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   1980
      TabIndex        =   3
      Top             =   1800
      Width           =   390
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   1980
      TabIndex        =   2
      Top             =   1480
      Width           =   390
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1980
      TabIndex        =   1
      Top             =   1215
      Width           =   390
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1980
      TabIndex        =   0
      Top             =   990
      Width           =   390
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   2640
      TabIndex        =   34
      Top             =   5160
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   2280
      TabIndex        =   35
      Top             =   5400
      Width           =   2475
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub Command1_Click()

End Sub

Public Sub Iniciar_Labels()
'Iniciamos los labels con los valores de los atributos y los skills
Dim I As Integer
For I = 1 To NUMATRIBUTOS
    Atri(I).Caption = UserAtributos(I)
Next
For I = 1 To NUMSKILLS
    Skills(I).Caption = UserSkills(I)
Next


Label4(1).Caption = UserReputacion.AsesinoRep
Label4(2).Caption = UserReputacion.BandidoRep
Label4(3).Caption = UserReputacion.BurguesRep
Label4(4).Caption = UserReputacion.LadronesRep
Label4(5).Caption = UserReputacion.NobleRep
Label4(6).Caption = UserReputacion.PlebeRep

If UserReputacion.Promedio < 0 Then
    Label4(7).ForeColor = vbRed
    Label4(7).Caption = "Status: CRIMINAL"
Else
    Label4(7).ForeColor = vbBlue
    Label4(7).Caption = "Status: Ciudadano"
End If

With UserEstadisticas
    Label6(0).Caption = .CriminalesMatados
    Label6(1).Caption = .CiudadanosMatados
    Label6(2).Caption = .UsuariosMatados
    Label6(3).Caption = .NpcsMatados
    Label6(4).Caption = .Clase
    Label6(5).Caption = .PenaCarcel
End With

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub
Private Sub Form_Load()
Call cargarImagenRes(frmEstadisticas, 115)
'Valores máximos y mínimos para el ScrollBar
 '  frmEstadisticas.Picture = LoadPicture(App.Path & _
  '  "\Graficos\Estadisticas.jpg")
   
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

