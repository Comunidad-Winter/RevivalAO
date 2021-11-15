VERSION 5.00
Begin VB.Form frmEligeAlineacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5250
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5790
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   5790
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblSalir 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Top             =   4860
      Width           =   915
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmEligeAlineacion.frx":0000
      ForeColor       =   &H00000000&
      Height          =   645
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   4140
      Width           =   5505
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmEligeAlineacion.frx":00D5
      ForeColor       =   &H00000000&
      Height          =   645
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   5505
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00400040&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmEligeAlineacion.frx":01B1
      ForeColor       =   &H00000000&
      Height          =   645
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   2295
      Width           =   5505
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmEligeAlineacion.frx":025D
      ForeColor       =   &H00000000&
      Height          =   645
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1350
      Width           =   5505
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmEligeAlineacion.frx":0326
      ForeColor       =   &H00000000&
      Height          =   825
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   270
      Width           =   5505
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineación del mal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   3915
      Width           =   1680
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineación criminal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   1680
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineación neutral"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1635
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineación legal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1125
      Width           =   1455
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineación Real"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   45
      Width           =   1455
   End
End
Attribute VB_Name = "frmEligeAlineacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'odio programar sin tiempo (c) el oso

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte

    For i = 0 To 4
        lblDescripcion(i).BorderStyle = 0
        lblDescripcion(i).BackStyle = 0
    Next i
    
End Sub


Private Sub lblDescripcion_Click(index As Integer)
Dim s As String
    
    Select Case index
        Case 0
            s = "armada"
        Case 1
            s = "legal"
        Case 2
            s = "neutro"
        Case 3
            s = "criminal"
        Case 4
            s = "mal"
    End Select
    
    s = "/fundarclan " & s
    Call SendData(s)
    Unload Me
End Sub

Private Sub lblDescripcion_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblDescripcion(index).BorderStyle = 1
    lblDescripcion(index).BackStyle = 1
    Select Case index
        Case 0
            lblDescripcion(index).BackColor = &H400000
        Case 1
            lblDescripcion(index).BackColor = &H800000
        Case 2
            lblDescripcion(index).BackColor = 4194368
        Case 3
            lblDescripcion(index).BackColor = &H80&
        Case 4
            lblDescripcion(index).BackColor = &H40&
    End Select
End Sub




Private Sub lblSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
'Valores máximos y mínimos para el ScrollBar
   
End Sub
