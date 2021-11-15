VERSION 5.00
Begin VB.Form frmCharInfo 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Información del personaje"
   ClientHeight    =   6705
   ClientLeft      =   -60
   ClientTop       =   -165
   ClientWidth     =   6705
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MouseIcon       =   "frmCharInfo.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6000
      Width           =   1485
   End
   Begin VB.CommandButton echar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Echar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      MouseIcon       =   "frmCharInfo.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6000
      Width           =   1125
   End
   Begin VB.TextBox txtMiembro 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000004&
      Height          =   1110
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   4600
      Width           =   5790
   End
   Begin VB.TextBox txtPeticiones 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000004&
      Height          =   1110
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   3280
      Width           =   5790
   End
   Begin VB.CommandButton desc 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Peticion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      MouseIcon       =   "frmCharInfo.frx":02A4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   1125
   End
   Begin VB.CommandButton Aceptar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4980
      MouseIcon       =   "frmCharInfo.frx":03F6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   1440
   End
   Begin VB.CommandButton Rechazar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rechazar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      MouseIcon       =   "frmCharInfo.frx":0548
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   1035
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   4800
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Nombre 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1320
      TabIndex        =   15
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Nivel 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1320
      TabIndex        =   14
      Top             =   1920
      Width           =   405
   End
   Begin VB.Label Clase 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Clase:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1320
      TabIndex        =   13
      Top             =   1440
      Width           =   450
   End
   Begin VB.Label Raza 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Raza:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1320
      TabIndex        =   12
      Top             =   1200
      Width           =   420
   End
   Begin VB.Label Genero 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Genero:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1320
      TabIndex        =   11
      Top             =   1680
      Width           =   585
   End
   Begin VB.Label Oro 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1320
      TabIndex        =   10
      Top             =   2160
      Width           =   330
   End
   Begin VB.Label Banco 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Banco:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1320
      TabIndex        =   9
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label status 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   9360
      TabIndex        =   8
      Top             =   7200
      Width           =   525
   End
   Begin VB.Label guildactual 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Clan Actual:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3840
      TabIndex        =   7
      Top             =   1200
      Width           =   870
   End
   Begin VB.Label ejercito 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Faccion:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4200
      TabIndex        =   6
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label Ciudadanos 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ciuda"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5520
      TabIndex        =   5
      Top             =   1680
      Width           =   405
   End
   Begin VB.Label criminales 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Crimin"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5400
      TabIndex        =   4
      Top             =   1920
      Width           =   435
   End
   Begin VB.Label reputacion 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4560
      TabIndex        =   3
      Top             =   2160
      Width           =   870
   End
End
Attribute VB_Name = "frmCharInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public frmmiembros As Boolean
Public frmsolicitudes As Boolean
Private Sub Aceptar_Click()
frmmiembros = False
frmsolicitudes = False
Call SendData("ACEPTARI" & Nombre)
Unload frmGuildLeader
Call SendData("GLINFO")
Unload Me
End Sub
Private Sub Command1_Click()
Call SendData("ECHARCLA" & Right(Nombre, Len(Nombre) - 7))
frmmiembros = False
frmsolicitudes = False
Unload frmGuildLeader
Call SendData("GLINFO")
Unload Me
End Sub
Public Sub parseCharInfo(ByVal Rdata As String)

If frmmiembros Then
    Rechazar.Visible = False
    Aceptar.Visible = False
    echar.Visible = True
    desc.Visible = False
Else
    Rechazar.Visible = True
    Aceptar.Visible = True
    echar.Visible = False
    desc.Visible = True
End If

'    tstr = Personaje & "¬"1
'    tstr = tstr & GetVar(UserFile, "INIT", "Raza") & "¬"2
'    tstr = tstr & GetVar(UserFile, "INIT", "Clase") & "¬"3
'    tstr = tstr & GetVar(UserFile, "INIT", "Genero") & "¬"4
'    tstr = tstr & GetVar(UserFile, "STATS", "ELV") & "¬"5
'    tstr = tstr & GetVar(UserFile, "STATS", "GLD") & "¬"6
'    tstr = tstr & GetVar(UserFile, "STATS", "Banco") & "¬"7
'    tstr = tstr & GetVar(UserFile, "REP", "Promedio") & "¬"8


Nombre.Caption = ReadField(1, Rdata, Asc("¬"))
Raza.Caption = ReadField(2, Rdata, Asc("¬"))
Clase.Caption = ReadField(3, Rdata, Asc("¬"))
Genero.Caption = ReadField(4, Rdata, Asc("¬"))
Nivel.Caption = ReadField(5, Rdata, Asc("¬"))
Oro.Caption = ReadField(6, Rdata, Asc("¬"))
Banco.Caption = ReadField(7, Rdata, Asc("¬"))
Me.reputacion.Caption = ReadField(8, Rdata, Asc("¬"))


'    Peticiones = GetVar(UserFile, "GUILDS", "Pedidos")9
'    tstr = tstr & IIf(Len(Peticiones > 400), ".." & Right$(Peticiones, 400), Peticiones) & "¬"
    
'    Miembro = GetVar(UserFile, "GUILDS", "Miembro")10
'    tstr = tstr & IIf(Len(Miembro) > 400, ".." & Right$(Miembro, 400), Miembro) & "¬"

Me.txtPeticiones.Text = ReadField(9, Rdata, Asc("¬"))
Me.txtMiembro.Text = ReadField(10, Rdata, Asc("¬"))


'GuildActual = val(GetVar(UserFile, "GUILD", "GuildIndex"))11
Me.guildactual.Caption = "Clan: " & ReadField(11, Rdata, Asc("¬"))


'    tstr = tstr & GetVar(UserFile, "FACCIONES", "EjercitoReal") & "¬"12
'    tstr = tstr & GetVar(UserFile, "FACCIONES", "EjercitoCaos") & "¬"13
'    tstr = tstr & GetVar(UserFile, "FACCIONES", "CiudMatados") & "¬"14
'    tstr = tstr & GetVar(UserFile, "FACCIONES", "CrimMatados") & "¬"15

Me.ejercito.Caption = IIf(Val(ReadField(12, Rdata, Asc("¬"))) <> 0, "Armada Real", IIf(Val(ReadField(13, Rdata, Asc("¬"))) <> 0, "Legión Oscura", "-"))

Ciudadanos.Caption = ReadField(14, Rdata, Asc("¬"))
criminales.Caption = ReadField(15, Rdata, Asc("¬"))


status.Caption = IIf(Val(ReadField(8, Rdata, Asc("¬"))) > 0, " (Ciudadano)", " (Criminal)")
status.ForeColor = IIf(Val(ReadField(8, Rdata, Asc("¬"))) > 0, vbBlue, vbRed)
Me.Show vbModeless, frmMain


End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub desc_Click()
Call SendData("ENVCOMEN" & Nombre)
End Sub

Private Sub echar_Click()
Call SendData("ECHARCLA" & Nombre)
frmmiembros = False
frmsolicitudes = False
Unload frmGuildLeader
Call SendData("GLINFO")
Unload Me
End Sub

Private Sub Form_Load()
Call cargarImagenRes(frmCharInfo, 107)
'frmCharInfo.Picture = LoadPicture(App.Path & _
 '   "\Graficos\CharInfo.jpg")
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Rechazar_Click()
Load frmCommet
frmCommet.T = RECHAZOPJ
frmCommet.Nombre = Nombre
frmCommet.Caption = "Ingrese motivo para rechazo"
frmCommet.Show vbModeless, frmCharInfo

End Sub

