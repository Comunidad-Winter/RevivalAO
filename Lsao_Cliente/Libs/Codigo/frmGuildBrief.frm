VERSION 5.00
Begin VB.Form frmGuildBrief 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Detalles del Clan"
   ClientHeight    =   7830
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7605
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
   ScaleHeight     =   7830
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Desc 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   6000
      Width           =   6855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ofrecer Paz"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      MouseIcon       =   "frmGuildBrief.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton aliado 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ofrecer Alianza"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      MouseIcon       =   "frmGuildBrief.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton Guerra 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Declarar Guerra"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      MouseIcon       =   "frmGuildBrief.frx":02A4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   4800
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   960
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label Codex 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   21
      Top             =   3600
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   20
      Top             =   3840
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   19
      Top             =   4080
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   18
      Top             =   4320
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   17
      Top             =   4560
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   16
      Top             =   4800
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   15
      Top             =   5040
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   14
      Top             =   5280
      Width           =   6735
   End
   Begin VB.Label nombre 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1200
      TabIndex        =   13
      Top             =   480
      Width           =   615
   End
   Begin VB.Label fundador 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Fundador:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1200
      TabIndex        =   12
      Top             =   720
      Width           =   750
   End
   Begin VB.Label creacion 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de creacion:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1920
      TabIndex        =   11
      Top             =   960
      Width           =   1365
   End
   Begin VB.Label lider 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Lider:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   10
      Top             =   1200
      Width           =   405
   End
   Begin VB.Label web 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Web site:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   720
      TabIndex        =   9
      Top             =   1440
      Width           =   690
   End
   Begin VB.Label Miembros 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Miembros:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1320
      TabIndex        =   8
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label eleccion 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Elecciones:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3480
      TabIndex        =   7
      Top             =   1920
      Width           =   795
   End
   Begin VB.Label lblAlineacion 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Alineacion:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1320
      TabIndex        =   6
      Top             =   2160
      Width           =   780
   End
   Begin VB.Label Enemigos 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Clanes Enemigos:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Top             =   2400
      Width           =   1260
   End
   Begin VB.Label Aliados 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Clanes Aliados:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1560
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label antifaccion 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Puntos Antifaccion:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10680
      TabIndex        =   3
      Top             =   9480
      Width           =   1395
   End
End
Attribute VB_Name = "frmGuildBrief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public EsLeader As Boolean


Public Sub ParseGuildInfo(ByVal Buffer As String)

'[MaTeO 1]
Dim IsMyClan As Boolean
If UserCharIndex > 0 Then
    Dim MyClan As String
    
    MyClan = Right$(charlist(UserCharIndex).Nombre, Len(charlist(UserCharIndex).Nombre) - InStr(1, charlist(UserCharIndex).Nombre, "<"))
    MyClan = Left$(MyClan, Len(MyClan) - 1)
    
    If UCase$(MyClan) = ReadField(1, Buffer, Asc("¬")) Then
        IsMyClan = True
    End If
End If

If Not EsLeader Or (EsLeader And IsMyClan) Then
'[/MaTeO 1]
    Guerra.Visible = False
    aliado.Visible = False
    Command3.Visible = False
Else
    Guerra.Visible = True
    aliado.Visible = True
    Command3.Visible = True
End If

Nombre.Caption = ReadField(1, Buffer, Asc("¬"))
fundador.Caption = ReadField(2, Buffer, Asc("¬"))
creacion.Caption = ReadField(3, Buffer, Asc("¬"))
lider.Caption = ReadField(4, Buffer, Asc("¬"))
web.Caption = ReadField(5, Buffer, Asc("¬"))
Miembros.Caption = ReadField(6, Buffer, Asc("¬"))
eleccion.Caption = ReadField(7, Buffer, Asc("¬"))
'Oro.Caption = "Oro:" & ReadField(8, Buffer, Asc("¬"))
lblAlineacion.Caption = ReadField(8, Buffer, Asc("¬"))
Enemigos.Caption = ReadField(9, Buffer, Asc("¬"))
Aliados.Caption = ReadField(10, Buffer, Asc("¬"))
antifaccion.Caption = ReadField(11, Buffer, Asc("¬"))

Dim T As Long

For T = 1 To 8
    Codex(T - 1).Caption = ReadField(11 + T, Buffer, Asc("¬"))
Next T

Dim des As String

des = ReadField(20, Buffer, Asc("¬"))
desc.Text = Replace(des, "º", vbCrLf)

Me.Show vbModal, frmMain

End Sub

Private Sub aliado_Click()
frmCommet.Nombre = Right(Nombre.Caption, Len(Nombre.Caption) - 7)
frmCommet.T = ALIANZA
frmCommet.Caption = "Ingrese propuesta de alianza"
Call frmCommet.Show(vbModal, frmGuildBrief)

'Call SendData("OFRECALI" & Right(Nombre, Len(Nombre) - 7))
'Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()


End Sub

Private Sub Command3_Click()
frmCommet.Nombre = Right(Nombre.Caption, Len(Nombre.Caption) - 7)
frmCommet.T = PAZ
frmCommet.Caption = "Ingrese propuesta de paz"
Call frmCommet.Show(vbModal, frmGuildBrief)
'Unload Me
End Sub



Private Sub Form_Load()

Call cargarImagenRes(frmGuildBrief, 117)
'Valores máximos y mínimos para el ScrollBar
'  frmGuildBrief.Picture = LoadPicture(App.Path & _
 '   "\Graficos\InfoClan.jpg")
End Sub

Private Sub Guerra_Click()
Call SendData("DECGUERR" & Right(Nombre.Caption, Len(Nombre.Caption) - 7))
Unload Me
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image2_Click()
Call frmGuildSol.RecieveSolicitud(Nombre)
Call frmGuildSol.Show(vbModal, frmGuildBrief)
'Unload Me

End Sub
