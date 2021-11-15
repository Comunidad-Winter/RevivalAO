VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox PasswordTXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   13
      Top             =   6720
      Width           =   3615
   End
   Begin SHDocVwCtl.WebBrowser Noticias 
      CausesValidation=   0   'False
      Height          =   3135
      Left            =   360
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1200
      Width           =   6435
      ExtentX         =   11351
      ExtentY         =   5530
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
   Begin VB.TextBox NombreTXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   480
      TabIndex        =   11
      Top             =   5760
      Width           =   3615
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   11400
      Top             =   2880
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Slot 6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   10
      Top             =   6540
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Slot 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4320
      MaskColor       =   &H0000FFFF&
      TabIndex        =   9
      Top             =   5040
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Slot 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   8
      Top             =   5340
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Slot 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   7
      Top             =   5640
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Slot 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   6
      Top             =   5940
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Slot 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   5
      Top             =   6240
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ListBox lst_servers 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   450
      ItemData        =   "frmConnect.frx":000C
      Left            =   9480
      List            =   "frmConnect.frx":0013
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   10680
      TabIndex        =   0
      Text            =   "7666"
      Top             =   2760
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   9720
      TabIndex        =   1
      Text            =   "localhost"
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   600
      MouseIcon       =   "frmConnect.frx":0024
      MousePointer    =   99  'Custom
      Top             =   8280
      Width           =   3255
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   1320
      Top             =   2520
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   4440
      Top             =   3120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Image Image7 
      Height          =   495
      Left            =   5400
      Top             =   2160
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5520
      TabIndex        =   4
      Top             =   7080
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   4680
      TabIndex        =   3
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   600
      MouseIcon       =   "frmConnect.frx":0CEE
      MousePointer    =   99  'Custom
      Top             =   7200
      Width           =   3285
   End
   Begin VB.Image imgServEspana 
      Height          =   315
      Left            =   3360
      MousePointer    =   99  'Custom
      Top             =   360
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgServArgentina 
      Height          =   435
      Left            =   4080
      MousePointer    =   99  'Custom
      Top             =   240
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgGetPass 
      Height          =   495
      Left            =   4800
      MousePointer    =   99  'Custom
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   585
      Index           =   0
      Left            =   600
      MouseIcon       =   "frmConnect.frx":19B8
      MousePointer    =   99  'Custom
      Top             =   7680
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   1
      Left            =   3360
      MousePointer    =   99  'Custom
      Top             =   240
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image1 
      Height          =   210
      Index           =   2
      Left            =   10320
      MousePointer    =   99  'Custom
      Top             =   1920
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image FONDO 
      Height          =   9000
      Left            =   0
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public Sub CargarLst()

Dim i As Integer

lst_servers.Clear

If ServersRecibidos Then
    Call WriteVar(App.Path & "\init\sinfo.dat", "INIT", "Cant", UBound(ServersLst))
    For i = 1 To UBound(ServersLst)
        Call WriteVar(App.Path & "\init\sinfo.dat", "S" & i, "Desc", ServersLst(i).desc)
        Call WriteVar(App.Path & "\init\sinfo.dat", "S" & i, "IP", ServersLst(i).ip)
        Call WriteVar(App.Path & "\init\sinfo.dat", "S" & i, "PJ", Str(ServersLst(i).Puerto))
        Call WriteVar(App.Path & "\init\sinfo.dat", "S" & i, "P2", Str(ServersLst(i).PassRecPort))
        lst_servers.AddItem ServersLst(i).ip & ":" & ServersLst(i).Puerto & " - Desc:" & ServersLst(i).desc
    Next i
End If

End Sub

Private Sub Command1_Click()
CurServer = 0
IPdelServidor = IPTxt
PuertoDelServidor = PortTxt
End Sub


Private Sub Command2_Click()

frmMain.Inet1.URL = "http://ao.alkon.com.ar/admin/iplist2.txt"
RawServersList = frmMain.Inet1.OpenURL


If RawServersList = "" Then
    ServersRecibidos = False
    Call MsgBox("No se pudo cargar la lista de servidores")
    ReDim ServersLst(1)
    Exit Sub
Else
    ServersRecibidos = True
End If

Call InitServersList(RawServersList)
Call CargarLst

End Sub






Private Sub Form_Activate()
'On Error Resume Next
NombreTXT.SetFocus
If ServersRecibidos Then
    If CurServer <> 0 Then
        IPTxt = ServersLst(1).ip
        PortTxt = ServersLst(1).Puerto
    Else
        IPTxt = IPdelServidor
        PortTxt = PuertoDelServidor
    End If
    
    Call CargarLst
Else
    lst_servers.Clear
End If

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
        frmCargando.Show
        frmCargando.Refresh
        AddtoRichTextBox frmCargando.Status, "Cerrando LSAO.", 0, 0, 0, 1, 0, 1
        
        Call SaveGameini
        frmConnect.MousePointer = 1
        frmMain.MousePointer = 1
        prgRun = False
        
        AddtoRichTextBox frmCargando.Status, "Liberando recursos..."
        Call WriteVar(App.Path & "\init\sinfo.dat", "s10", "PJ", " 0")
        frmCargando.Refresh
        LiberarObjetosDX
        AddtoRichTextBox frmCargando.Status, "Hecho", 0, 0, 0, 1, 0, 1
        AddtoRichTextBox frmCargando.Status, "¡¡Gracias por jugar LSAO!!", 0, 0, 0, 1, 0, 1
        frmCargando.Refresh
        Call UnloadAllForms
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'Make Server IP and Port box visible
If KeyCode = vbKeyI And Shift = vbCtrlMask Then
    
    'Port
    PortTxt.Visible = False
    'Label4.Visible = True
    
    'Server IP
    PortTxt.Text = "9879"
    IPTxt.Text = "190.210.25.107"
    IPTxt.Visible = False
    'Label5.Visible = True
    
    KeyCode = 0
    Exit Sub
End If

End Sub

Private Sub Form_Load()
    '[CODE 002]:MatuX
    EngineRun = False
    '[END]
   
 Dim j
 For Each j In Image1()
    j.Tag = "0"
 Next
 PortTxt.Text = Config_Inicio.Puerto
 
 FONDO.Picture = LoadPicture(App.Path & "\Graficos\Conectar.jpg")
 Timer1.Enabled = False
 '[CODE]:MatuX
 '
 '  El código para mostrar la versión se genera acá para
 ' evitar que por X razones luego desaparezca, como suele
 ' pasar a veces :)
    version.Caption = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
 '[END]'
Noticias.Navigate ("http://www.symxsoft.net/noticia.html")
frmMain.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmBancoObj.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmBorrar.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmCambiaMotd.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmCantidad.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmCaptions.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmCargando.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmCarp.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmCharInfo.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmComerciar.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmComerciarUsu.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmCommet.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmConnect.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
FrmConsolaTorneo.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmEligeAlineacion.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmEntrenador.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmEstadisticas.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmForo.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmGuildAdm.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmGuildBrief.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmGuildDetails.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmGuildFoundation.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmGuildLeader.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmGuildNews.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmGuildSol.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmGuildURL.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmHerrero.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmKeypad.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmMapa.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmMensaje.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmMSG.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmOldPersonaje.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmOpciones.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmPanelGm.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmPasswdSinPadrinos.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmPeaceProp.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
FrmProcesos.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmProsesos.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmRecuperar.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmSkills3.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmSpawnList.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmtip.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
FrmTransferir.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmUserRequest.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmSoporte.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmSoporteGm.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmSoporteResp.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmRank.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")
frmContra.MouseIcon = LoadPicture(App.Path & "\Graficos\diablo.ico")

frmMain.MousePointer = vbCustom
frmRank.MousePointer = vbCustom
frmContra.MousePointer = vbCustom
frmBancoObj.MousePointer = vbCustom
frmSoporte.MousePointer = vbCustom
frmSoporteGm.MousePointer = vbCustom
frmSoporteResp.MousePointer = vbCustom
frmBorrar.MousePointer = vbCustom
frmCambiaMotd.MousePointer = vbCustom
frmCantidad.MousePointer = vbCustom
frmCaptions.MousePointer = vbCustom
frmCargando.MousePointer = vbCustom
frmCarp.MousePointer = vbCustom
frmCharInfo.MousePointer = vbCustom
frmComerciar.MousePointer = vbCustom
frmComerciarUsu.MousePointer = vbCustom
frmCommet.MousePointer = vbCustom
frmConnect.MousePointer = vbCustom
FrmConsolaTorneo.MousePointer = vbCustom
frmEligeAlineacion.MousePointer = vbCustom
frmEntrenador.MousePointer = vbCustom
frmEstadisticas.MousePointer = vbCustom
frmForo.MousePointer = vbCustom
frmGuildAdm.MousePointer = vbCustom
frmGuildBrief.MousePointer = vbCustom
frmGuildDetails.MousePointer = vbCustom
frmGuildFoundation.MousePointer = vbCustom
frmGuildLeader.MousePointer = vbCustom
frmGuildNews.MousePointer = vbCustom
frmGuildSol.MousePointer = vbCustom
frmGuildURL.MousePointer = vbCustom
frmHerrero.MousePointer = vbCustom
frmKeypad.MousePointer = vbCustom
frmMapa.MousePointer = vbCustom
frmMensaje.MousePointer = vbCustom
frmMSG.MousePointer = vbCustom
frmOldPersonaje.MousePointer = vbCustom
frmOpciones.MousePointer = vbCustom
frmPanelGm.MousePointer = vbCustom
frmPasswdSinPadrinos.MousePointer = vbCustom
frmPeaceProp.MousePointer = vbCustom
FrmProcesos.MousePointer = vbCustom
frmProsesos.MousePointer = vbCustom
frmRecuperar.MousePointer = vbCustom
frmSkills3.MousePointer = vbCustom
frmSpawnList.MousePointer = vbCustom
frmtip.MousePointer = vbCustom
FrmTransferir.MousePointer = vbCustom
frmUserRequest.MousePointer = vbCustom


End Sub



Private Sub Image1_Click(index As Integer)

CurServer = 0
'IPdelServidor = IPTxt
PuertoDelServidor = PortTxt
IPdelServidor = "190.210.25.107" '

Call Audio.PlayWave(SND_CLICK)

Select Case index
    Case 0
        
        If Musica Then
            Call Audio.PlayMIDI("7.mid")
        End If
        
        
        
        'frmCrearPersonaje.Show vbModal
        EstadoLogin = Dados
#If UsarWrench = 1 Then
        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
        End If
        frmMain.Socket1.HostName = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect
#Else
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
        End If
        frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If
        Me.MousePointer = 11

        
    Case 1
    
        frmOldPersonaje.Show vbModal
        
    Case 2
        On Error GoTo errH
        Call Shell(App.Path & "\RECUPERAR.EXE", vbNormalFocus)

End Select
Exit Sub

errH:
    Call MsgBox("No se encuentra el programa recuperar.exe", vbCritical, "LSAO")
End Sub

Private Sub Image2_Click()
Call Audio.PlayWave(SND_CLICK)

#If UsarWrench = 1 Then
            If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
    #Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
    #End If
            If frmConnect.MousePointer = 11 Then
                Exit Sub
            End If
           
           
            UserName = NombreTXT.Text

        Dim aux As String
        aux = PasswordTXT.Text
#If SeguridadAlkon Then
        UserPassword = MD5.GetMD5String(aux)
        Call MD5.MD5Reset
#Else
        UserPassword = aux
#End If
            If CheckUserData(False) = True Then
                'SendNewChar = False
                EstadoLogin = Normal
                Me.MousePointer = 11
    #If UsarWrench = 1 Then
                frmMain.Socket1.HostName = CurServerIp
                frmMain.Socket1.RemotePort = CurServerPort
                frmMain.Socket1.Connect
    #Else
                If frmMain.Winsock1.State <> sckClosed Then _
                    frmMain.Winsock1.Close
                frmMain.Winsock1.Connect CurServerIp, CurServerPort
    #End If
            End If
End Sub

Private Sub Image3_Click()
frmRecuPass.Show
End Sub

Private Sub Image5_Click()
If Option1(0).value = True Then
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "NICK1", NombreTXT.Text)
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "PASS1", PasswordTXT.Text)
ElseIf Option1(1).value = True Then
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "NICK2", NombreTXT.Text)
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "PASS2", PasswordTXT.Text)
ElseIf Option1(2).value = True Then
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "NICK3", NombreTXT.Text)
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "PASS3", PasswordTXT.Text)
ElseIf Option1(3).value = True Then
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "NICK4", NombreTXT.Text)
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "PASS4", PasswordTXT.Text)
ElseIf Option1(4).value = True Then
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "NICK5", NombreTXT.Text)
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "PASS5", PasswordTXT.Text)
ElseIf Option1(5).value = True Then
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "NICK6", "Slot Vacio")
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "PASS6", " ")
End If
End Sub
'[/Standelf]

Private Sub Image6_Click()
'[Standelf]
'Sistema de Memorias por Slots de Usuarios
If Option1(0).value = True Then
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "NICK1", "Slot Vacio")
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "PASS1", " ")
ElseIf Option1(1).value = True Then
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "NICK2", "Slot Vacio")
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "PASS2", " ")
ElseIf Option1(2).value = True Then
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "NICK3", "Slot Vacio")
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "PASS3", " ")
ElseIf Option1(3).value = True Then
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "NICK4", "Slot Vacio")
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "PASS4", " ")
ElseIf Option1(4).value = True Then
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "NICK5", "Slot Vacio")
Call WriteVar(IniPath & "LSAO.ini", "SLOTS", "PASS5", " ")
ElseIf Option1(5).value = True Then
Call WriteVar(IniPath & "Geo-AO.ini", "SLOTS", "NICK6", NombreTXT.Text)
Call WriteVar(IniPath & "Geo-AO.ini", "SLOTS", "PASS6", PasswordTXT.Text)
End If

'[/Standelf]
End Sub

Private Sub Image7_Click()
If Option1(0).value = True Then
NombreTXT.Text = GetVar(IniPath & "LSAO.ini", "SLOTS", "NICK1")
PasswordTXT.Text = GetVar(IniPath & "LSAO.ini", "SLOTS", "PASS1")
ElseIf Option1(1).value = True Then
NombreTXT.Text = GetVar(IniPath & "LSAO.ini", "SLOTS", "NICK2")
PasswordTXT.Text = GetVar(IniPath & "LSAO.ini", "SLOTS", "PASS2")
ElseIf Option1(2).value = True Then
NombreTXT.Text = GetVar(IniPath & "LSAO.ini", "SLOTS", "NICK3")
PasswordTXT.Text = GetVar(IniPath & "LSAO.ini", "SLOTS", "PASS3")
ElseIf Option1(3).value = True Then
NombreTXT.Text = GetVar(IniPath & "LSAO.ini", "SLOTS", "NICK4")
PasswordTXT.Text = GetVar(IniPath & "LSAO.ini", "SLOTS", "PASS4")
ElseIf Option1(4).value = True Then
NombreTXT.Text = GetVar(IniPath & "LSAO.ini", "SLOTS", "NICK5")
PasswordTXT.Text = GetVar(IniPath & "LSAO.ini", "SLOTS", "PASS5")
ElseIf Option1(5).value = True Then
NombreTXT.Text = GetVar(IniPath & "LSAO.ini", "SLOTS", "NICK6")
PasswordTXT.Text = GetVar(IniPath & "LSAO.ini", "SLOTS", "PASS6")
End If
End Sub

Private Sub imgGetPass_Click()
On Error GoTo errH

    Call Audio.PlayWave(SND_CLICK)
    Call Shell(App.Path & "\RECUPERAR.EXE", vbNormalFocus)
    'Call frmRecuperar.Show(vbModal, frmConnect)
    Exit Sub
errH:
    Call MsgBox("No se encuentra el programa recuperar.exe", vbCritical, "LSAO")
End Sub

Private Sub imgServArgentina_Click()
    Call Audio.PlayWave(SND_CLICK)
    IPTxt.Text = IPdelServidor
    PortTxt.Text = PuertoDelServidor
End Sub

Private Sub imgServEspana_Click()
    Call Audio.PlayWave(SND_CLICK)
    IPTxt.Text = "62.42.193.233"
    PortTxt.Text = "9879"

End Sub


Private Sub lst_servers_Click()
If ServersRecibidos Then
    CurServer = lst_servers.listIndex + 1
    IPTxt = ServersLst(CurServer).ip
    PortTxt = ServersLst(CurServer).Puerto
End If

End Sub

Private Sub PasswordTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Image2_Click
    End If
End Sub

Private Sub Timer1_Timer()
frmConnect.Label1.Caption = " "
frmConnect.Timer1.Enabled = False
frmConnect.MousePointer = 1
End Sub

