VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Argentum Online"
   ClientHeight    =   3720
   ClientLeft      =   1950
   ClientTop       =   1815
   ClientWidth     =   5190
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3720
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Timer bandas 
      Interval        =   60000
      Left            =   960
      Top             =   1800
   End
   Begin VB.Timer deat 
      Interval        =   60000
      Left            =   360
      Top             =   1800
   End
   Begin VB.Timer Mascotas 
      Interval        =   60000
      Left            =   3960
      Top             =   1800
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   3000
      Top             =   120
   End
   Begin VB.Frame Frame2 
      Caption         =   "Logs Usuarios"
      Height          =   1575
      Left            =   0
      TabIndex        =   11
      Top             =   2160
      Width           =   5175
      Begin RichTextLib.RichTextBox RichTextBox2 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         TextRTF         =   $"frmMain.frx":1042
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   0
         ReadOnly        =   -1  'True
         TextRTF         =   $"frmMain.frx":10BB
      End
      Begin VB.Label Label3 
         Caption         =   "Log Clan"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Log Normal"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   60000
      Left            =   4560
      Top             =   0
   End
   Begin VB.Timer torneos 
      Interval        =   60000
      Left            =   3960
      Top             =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Poner rey del castillo"
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   1800
      Width           =   2535
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3600
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer tmrResp 
      Interval        =   60000
      Left            =   480
      Top             =   120
   End
   Begin VB.CheckBox SUPERLOG 
      Caption         =   "log"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton CMDDUMP 
      Caption         =   "dump"
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.Timer FX 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1440
      Top             =   540
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   945
      Top             =   540
   End
   Begin VB.Timer CmdExec 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   60
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1440
      Top             =   60
   End
   Begin VB.Timer tPiqueteC 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   480
      Top             =   540
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   480
      Top             =   1020
   End
   Begin VB.Timer npcataca 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   960
      Top             =   1020
   End
   Begin VB.Timer KillLog 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1920
      Top             =   60
   End
   Begin VB.Timer TIMER_AI 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1935
      Top             =   540
   End
   Begin VB.Frame Frame1 
      Caption         =   "BroadCast"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4935
      Begin VB.Timer RespawnNPC 
         Interval        =   60000
         Left            =   1320
         Top             =   240
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Broadcast consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Broadcast clientes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox BroadMsg 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   4320
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Escuch 
      Caption         =   "Label2"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label CantUsuarios 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de usuarios:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1725
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   45
   End
   Begin VB.Menu mnuControles 
      Caption         =   "Argentum"
      Begin VB.Menu mnuServidor 
         Caption         =   "Configuracion"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "Systray Servidor"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar Servidor"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public ESCUCHADAS As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
   
Const NIM_ADD = 0
Const NIM_MODIFY = 1
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, ID As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uID = ID
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function

Sub CheckIdleUser()
Dim iUserIndex As Integer

For iUserIndex = 1 To MaxUsers
   
   'Conexion activa? y es un usuario loggeado?
   If UserList(iUserIndex).ConnID <> -1 And UserList(iUserIndex).flags.UserLogged Then
        'Actualiza el contador de inactividad
        UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
        If UserList(iUserIndex).Counters.IdleCount >= IdleLimit Then
            Call SendData(SendTarget.toindex, iUserIndex, 0, "Has sido desconectado por permanecer mas de 30 minutos inactivo.")
            'mato los comercios seguros
            If UserList(iUserIndex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(iUserIndex).ComUsu.DestUsu).flags.UserLogged Then
                    If UserList(UserList(iUserIndex).ComUsu.DestUsu).ComUsu.DestUsu = iUserIndex Then
                        Call SendData(SendTarget.toindex, UserList(iUserIndex).ComUsu.DestUsu, 0, "||Comercio cancelado por el otro usuario" & FONTTYPE_TALK)
                        Call FinComerciarUsu(UserList(iUserIndex).ComUsu.DestUsu)
                    End If
                End If
                Call FinComerciarUsu(iUserIndex)
            End If
            Call Cerrar_Usuario(iUserIndex)
        End If
  End If
  
Next iUserIndex

End Sub



Private Sub Auditoria_Timer()
On Error GoTo errhand

Call PasarSegundo 'sistema de desconexion de 10 segs

Call ActualizaEstadisticasWeb
Call ActualizaStatsES



Exit Sub

errhand:
Call LogError("Error en Timer Auditoria. Err: " & Err.Description & " - " & Err.Number)
End Sub

Private Sub AutoSave_Timer()

On Error GoTo errhandler
'fired every minute
Static Minutos As Long
Static MinutosLatsClean As Long
Static MinsSocketReset As Long
Static MinsPjesSave As Long
Static MinutosNumUsersCheck As Long

Dim i As Integer
Dim num As Long

MinsRunning = MinsRunning + 1

If MinsRunning = 60 Then
    Horas = Horas + 1
    If Horas = 24 Then
        Call SaveDayStats
        DayStats.MaxUsuarios = 0
        DayStats.Segundos = 0
        DayStats.Promedio = 0
        
        Horas = 0
        
    End If
    MinsRunning = 0
End If

    
Minutos = Minutos + 1

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Call ModAreas.AreasOptimizacion
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

#If UsarQueSocket = 1 Then
' ok la cosa es asi, este cacho de codigo es para
' evitar los problemas de socket. a menos que estes
' seguro de lo que estas haciendo, te recomiendo
' que lo dejes tal cual está.
' alejo.
MinsSocketReset = MinsSocketReset + 1
' cada 1 minutos hacer el checkeo
If MinsSocketReset >= 5 Then
    MinsSocketReset = 0
    For i = 1 To MaxUsers
        If UserList(i).ConnID <> -1 And Not UserList(i).flags.UserLogged Then
            If UserList(i).Counters.IdleCount > ((IntervaloCerrarConexion * 2) / 3) Then
                Call CloseSocket(i)
            End If
        End If
    Next i
    'Call ReloadSokcet
    
    Call LogCriticEvent("NumUsers: " & NumUsers & " WSAPISock2Usr: " & WSAPISock2Usr.Count)
End If
#End If

MinutosNumUsersCheck = MinutosNumUsersCheck + 1

If MinutosNumUsersCheck >= 2 Then
    MinutosNumUsersCheck = 0
    num = 0
    For i = 1 To MaxUsers
        If UserList(i).ConnID <> -1 And UserList(i).flags.UserLogged Then
            num = num + 1
        End If
    Next i
    If num <> NumUsers Then
        NumUsers = num
        'Call SendData(SendTarget.ToAdmins, 0, 0, "Servidor> Error en NumUsers. Contactar a algun Programador." & FONTTYPE_SERVER)
        Call LogCriticEvent("Num <> NumUsers")
    End If
End If

If Minutos = MinutosWs - 1 Then
    Call SendData(SendTarget.toall, 0, 0, "||Worldsave y Limpeza en 1 minuto ..." & FONTTYPE_VENENO)
End If

If Minutos >= MinutosWs Then
    Call DoBackUp
    Call aClon.VaciarColeccion
    Call mdParty.ActualizaExperiencias
    Call GuardarUsuarios
    Call SendData(SendTarget.toall, 0, 0, "||RevivalAo> WORLDSAVE TERMINADO CORRECTAMENTE. YA PUEDES SEGUIR JUGANDO!." & FONTTYPE_SERVER)
    Minutos = 0
End If

If MinutosLatsClean >= 15 Then
        MinutosLatsClean = 0
        Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
        Call LimpiarMundo
Else
        MinutosLatsClean = MinutosLatsClean + 1
End If

Call PurgarPenas
Call CheckIdleUser

'<<<<<-------- Log the number of users online ------>>>
Dim n As Integer
n = FreeFile()
Open App.Path & "\logs\numusers.log" For Output Shared As n
Print #n, NumUsers
Close #n
'<<<<<-------- Log the number of users online ------>>>

Exit Sub
errhandler:
    Call LogError("Error en TimerAutoSave " & Err.Number & ": " & Err.Description)

End Sub






Private Sub bandas_Timer()
On Error GoTo errordm:
bandasqls = bandasqls + 1
Select Case bandasqls
Case 38
Call SendData(SendTarget.toall, 0, 0, "||Guerra> En 10 minutos se realizará una Guerra RevivalAo." & FONTTYPE_GUILD)
Case 43
Call SendData(SendTarget.toall, 0, 0, "||Guerra> En 5 minutos se realizará una Guerra RevivalAo." & FONTTYPE_GUILD)
Case 47
Call SendData(SendTarget.toall, 0, 0, "||Guerra> En 1 minuto se realizará un Guerra RevivalAo." & FONTTYPE_GUILD)
Case 48
Call Ban_Comienza(32)
Case 49
If Banac = True Then
If CantidadGuerra < 6 Then
Call Banauto_Cancela
bandasqls = 1
Else
If Banesp = True Then
Call Banauto_Empieza
bandasqls = 1
Else
bandasqls = 1
End If
End If
End If
End Select
errordm:
End Sub

Private Sub CMDDUMP_Click()
On Error Resume Next

Dim i As Integer
For i = 1 To MaxUsers
    Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).name & " UserLogged: " & UserList(i).flags.UserLogged)
Next i

Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

End Sub

Private Sub CmdExec_Timer()
Dim i As Integer
Static n As Long

On Error Resume Next ':(((

n = n + 1

For i = 1 To MaxUsers
    If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
        If Not UserList(i).CommandsBuffer.IsEmpty Then
            Call HandleData(i, UserList(i).CommandsBuffer.Pop)
        End If
        If n >= 10 Then
            If UserList(i).ColaSalida.Count > 0 Then ' And UserList(i).SockPuedoEnviar Then
    #If UsarQueSocket = 1 Then
                Call IntentarEnviarDatosEncolados(i)
    '#ElseIf UsarQueSocket = 0 Then
    '            Call WrchIntentarEnviarDatosEncolados(i)
    '#ElseIf UsarQueSocket = 2 Then
    '            Call ServIntentarEnviarDatosEncolados(i)
    #ElseIf UsarQueSocket = 3 Then
        'NADA, el control deberia ocuparse de esto!!!
        'si la cola se llena, dispara un on close
    #End If
            End If
        End If
    End If
Next i

If n >= 10 Then
    n = 0
End If

Exit Sub
hayerror:

End Sub

Private Sub Command1_Click()
Call SendData(SendTarget.toall, 0, 0, "!!" & BroadMsg.Text & ENDC)
End Sub

Public Sub InitMain(ByVal f As Byte)

If f = 1 Then
    Call mnuSystray_Click
Else
    frmMain.Show
End If

End Sub

Private Sub Command2_Click()
Call SendData(SendTarget.toall, 0, 0, "||RevivalAo> " & BroadMsg.Text & FONTTYPE_SERVER)
End Sub

Private Sub Command3_Click()
Dim Npc1 As Integer
Dim Npc1Pos As WorldPos
Npc1 = 906
Npc1Pos.Map = 75
Npc1Pos.x = 48
Npc1Pos.y = 56
Dim Npc2 As Integer
Dim Npc2Pos As WorldPos
Npc2 = 910
Npc2Pos.Map = 77
Npc2Pos.x = 48
Npc2Pos.y = 56
Dim Npc3 As Integer
Dim Npc3Pos As WorldPos
Npc3 = 616
Npc3Pos.Map = 106
Npc3Pos.x = 48
Npc3Pos.y = 56
Dim Npc4 As Integer
Dim Npc4Pos As WorldPos
Npc4 = 617
Npc4Pos.Map = 107
Npc4Pos.x = 48
Npc4Pos.y = 56


        Call SpawnNpc(val(Npc1), Npc1Pos, True, False)
        Call SpawnNpc(val(Npc2), Npc2Pos, True, False)
        Call SpawnNpc(val(Npc3), Npc3Pos, True, False)
        Call SpawnNpc(val(Npc4), Npc4Pos, True, False)
      
End Sub

Private Sub deat_Timer()
On Error GoTo errordm:
tukiql = tukiql + 1
Select Case tukiql
Case 53
Call SendData(SendTarget.toall, 0, 0, "||DeathMatch> En 10 minutos se realizará un deathmatch automatico." & FONTTYPE_GUILD)
Case 58
Call SendData(SendTarget.toall, 0, 0, "||DeathMatch> En 5 minutos se realizará un deathmatch automatico." & FONTTYPE_GUILD)
Case 62
Call SendData(SendTarget.toall, 0, 0, "||DeathMatch> En 1 minutos se realizará un deathmatch automatico." & FONTTYPE_GUILD)
Case 63
Call death_comienza(RandomNumber(8, 16))
Case 65
If deathesp = True Then
Call Deathauto_Cancela
tukiql = 2
Else
tukiql = 2
End If
End Select
errordm:
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
   
   If Not Visible Then
        Select Case x \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
   
End Sub

Private Sub QuitarIconoSystray()
On Error Resume Next

'Borramos el icono del systray
Dim i As Integer
Dim nid As NOTIFYICONDATA

nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

i = Shell_NotifyIconA(NIM_DELETE, nid)
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Call QuitarIconoSystray

#If UsarQueSocket = 1 Then
Call LimpiaWsApi(frmMain.hWnd)
#ElseIf UsarQueSocket = 0 Then
Socket1.Cleanup
#ElseIf UsarQueSocket = 2 Then
Serv.Detener
#End If

Call DescargaNpcsDat


Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
Next

'Log
Dim n As Integer
n = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #n
Print #n, Date & " " & Time & " server cerrado."
Close #n

End

Set SonidosMapas = Nothing

End Sub

Private Sub FX_Timer()
On Error GoTo hayerror

Call SonidosMapas.ReproducirSonidosDeMapas

Exit Sub
hayerror:

End Sub

Private Sub GameTimer_Timer()
Dim iUserIndex As Integer
Dim bEnviarStats As Boolean
Dim bEnviarAyS As Boolean
Dim iNpcIndex As Integer

Static lTirarBasura As Long
Static lPermiteAtacar As Long
Static lPermiteCast As Long
Static lPermiteTrabajar As Long

'[Alejo]
If lPermiteAtacar < IntervaloUserPuedeAtacar Then
    lPermiteAtacar = lPermiteAtacar + 1
End If

If lPermiteCast < IntervaloUserPuedeCastear Then
    lPermiteCast = lPermiteCast + 1
End If

If lPermiteTrabajar < IntervaloUserPuedeTrabajar Then
     lPermiteTrabajar = lPermiteTrabajar + 1
End If
'[/Alejo]

On Error GoTo hayerror

 '<<<<<< Procesa eventos de los usuarios >>>>>>
 For iUserIndex = 1 To MaxUsers
   'Conexion activa?
   If UserList(iUserIndex).ConnID <> -1 Then
      '¿User valido?

      If UserList(iUserIndex).ConnIDValida And UserList(iUserIndex).flags.UserLogged Then
         
         '[Alejo-18-5]
         bEnviarStats = False
         bEnviarAyS = False
         
         UserList(iUserIndex).NumeroPaquetesPorMiliSec = 0

         
         Call DoTileEvents(iUserIndex, UserList(iUserIndex).pos.Map, UserList(iUserIndex).pos.x, UserList(iUserIndex).pos.y)
         
                    '[MaTeO 10]
        If UserList(iUserIndex).pos.Map = 70 Then
            UserList(iUserIndex).Counters.Laberinto = UserList(iUserIndex).Counters.Laberinto + 1
            If UserList(iUserIndex).Counters.Laberinto Mod Fix(5000 / GameTimer.Interval) = 0 Then
                Call SendData(SendTarget.toindex, iUserIndex, 0, "||¡Muevete o volveras al principio!" & FONTTYPE_INFO)
            End If
            If UserList(iUserIndex).Counters.Laberinto = Fix(20000 / GameTimer.Interval) Then
                Call WarpUserChar(iUserIndex, 70, 13, 12)
                UserList(iUserIndex).Counters.Laberinto = 0
            End If
        End If
        '[/MaTeO 10]
         If UserList(iUserIndex).flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
         If UserList(iUserIndex).flags.Ceguera = 1 Or _
            UserList(iUserIndex).flags.Estupidez Then Call EfectoCegueEstu(iUserIndex)
         
          
         If UserList(iUserIndex).flags.Muerto = 0 Then
               
               '[Consejeros]
               If UserList(iUserIndex).flags.Desnudo And UserList(iUserIndex).flags.Privilegios = PlayerType.User Then Call EfectoFrio(iUserIndex)
               If UserList(iUserIndex).flags.Meditando Then Call DoMeditar(iUserIndex)
               If UserList(iUserIndex).flags.Envenenado = 1 And UserList(iUserIndex).flags.Privilegios = PlayerType.User Then Call EfectoVeneno(iUserIndex, bEnviarStats)
               If UserList(iUserIndex).flags.AdminInvisible <> 1 And UserList(iUserIndex).flags.Invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
               If UserList(iUserIndex).flags.Mimetizado = 1 Then Call EfectoMimetismo(iUserIndex)
                
               Call DuracionPociones(iUserIndex)
                
               If Lloviendo Then
                    If Not Intemperie(iUserIndex) Then
                        If Not UserList(iUserIndex).flags.Descansar And (UserList(iUserIndex).flags.Hambre = 0 And UserList(iUserIndex).flags.Sed = 0) Then
                        'No esta descansando
                            
                            Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                            If bEnviarStats Then Call SendData(SendTarget.toindex, iUserIndex, 0, "ASH" & UserList(iUserIndex).Stats.MinHP): bEnviarStats = False
                            Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                            If bEnviarStats Then Call SendData(SendTarget.toindex, iUserIndex, 0, "ASS" & UserList(iUserIndex).Stats.MinSta): bEnviarStats = False
                            
                        ElseIf UserList(iUserIndex).flags.Descansar Then
                        'esta descansando
                            
                            Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                            If bEnviarStats Then Call SendData(SendTarget.toindex, iUserIndex, 0, "ASH" & UserList(iUserIndex).Stats.MinHP): bEnviarStats = False
                            Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                            If bEnviarStats Then Call SendData(SendTarget.toindex, iUserIndex, 0, "ASS" & UserList(iUserIndex).Stats.MinSta): bEnviarStats = False
                                 'termina de descansar automaticamente
                            If UserList(iUserIndex).Stats.MaxHP = UserList(iUserIndex).Stats.MinHP And _
                                UserList(iUserIndex).Stats.MaxSta = UserList(iUserIndex).Stats.MinSta Then
                                    Call SendData(SendTarget.toindex, iUserIndex, 0, "DOK")
                                    Call SendData(SendTarget.toindex, iUserIndex, 0, "||Has terminado de descansar." & FONTTYPE_INFO)
                                    UserList(iUserIndex).flags.Descansar = False
                            End If
                            
                        End If 'Not UserList(UserIndex).Flags.Descansar And (UserList(UserIndex).Flags.Hambre = 0 And UserList(UserIndex).Flags.Sed = 0)
                    End If
               Else
                    If Not UserList(iUserIndex).flags.Descansar And (UserList(iUserIndex).flags.Hambre = 0 And UserList(iUserIndex).flags.Sed = 0) Then
                    'No esta descansando
                        
                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                        If bEnviarStats Then Call SendData(SendTarget.toindex, iUserIndex, 0, "ASH" & UserList(iUserIndex).Stats.MinHP): bEnviarStats = False
                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                        If bEnviarStats Then Call SendData(SendTarget.toindex, iUserIndex, 0, "ASS" & UserList(iUserIndex).Stats.MinSta): bEnviarStats = False
                        
                    ElseIf UserList(iUserIndex).flags.Descansar Then
                    'esta descansando
                        
                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                        If bEnviarStats Then Call SendData(SendTarget.toindex, iUserIndex, 0, "ASH" & UserList(iUserIndex).Stats.MinHP): bEnviarStats = False
                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                        If bEnviarStats Then Call SendData(SendTarget.toindex, iUserIndex, 0, "ASS" & UserList(iUserIndex).Stats.MinSta): bEnviarStats = False
                             'termina de descansar automaticamente
                        If UserList(iUserIndex).Stats.MaxHP = UserList(iUserIndex).Stats.MinHP And _
                            UserList(iUserIndex).Stats.MaxSta = UserList(iUserIndex).Stats.MinSta Then
                                Call SendData(SendTarget.toindex, iUserIndex, 0, "DOK")
                                Call SendData(SendTarget.toindex, iUserIndex, 0, "||Has terminado de descansar." & FONTTYPE_INFO)
                                UserList(iUserIndex).flags.Descansar = False
                        End If
                        
                    End If 'Not UserList(UserIndex).Flags.Descansar And (UserList(UserIndex).Flags.Hambre = 0 And UserList(UserIndex).Flags.Sed = 0)
               End If
               
               If bEnviarAyS Then Call EnviarHambreYsed(iUserIndex)

               If UserList(iUserIndex).NroMacotas > 0 Then Call TiempoInvocacion(iUserIndex)
       End If 'Muerto
     Else 'no esta logeado?
     'UserList(iUserIndex).Counters.IdleCount = 0
     '[Gonzalo]: deshabilitado para el nuevo sistema de tiraje
     'de dados :)
        
        UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
        If UserList(iUserIndex).Counters.IdleCount > IntervaloParaConexion Then
              UserList(iUserIndex).Counters.IdleCount = 0
              Call CloseSocket(iUserIndex)
        End If
        
     End If 'UserLogged

   End If

   Next iUserIndex

'[Alejo]
If Not lPermiteAtacar < IntervaloUserPuedeAtacar Then
    lPermiteAtacar = 0
End If

If Not lPermiteCast < IntervaloUserPuedeCastear Then
    lPermiteCast = 0
End If

If Not lPermiteTrabajar < IntervaloUserPuedeTrabajar Then
     lPermiteTrabajar = 0
End If

Exit Sub
hayerror:
LogError ("Error en GameTimer: " & Err.Description & " UserIndex = " & iUserIndex)
'[/Alejo]
  'DoEvents
End Sub

Private Sub Mascotas_Timer()
Dim Npc1 As Integer
Dim Npc1Pos As WorldPos
Npc1 = RandomNumber(924, 939)
Npc1Pos.Map = 30
Npc1Pos.x = 61
Npc1Pos.y = 38

mariano = mariano + 1
Select Case mariano
Case 475
Call SendData(SendTarget.toall, 0, 0, "||RevivalAo> En 5 minutos se invocara un Domador." & FONTTYPE_GUILD)
Case 479
Call SendData(SendTarget.toall, 0, 0, "||RevivalAo> En 1 minuto se invocara un Domador." & FONTTYPE_GUILD)
Case 480
Call SendData(SendTarget.toall, 0, 0, "||RevivalAo> Se ha invocado un domador en el mapa 30." & FONTTYPE_GUILD)
Call SendData(SendTarget.toall, 0, 0, "TW122")
Call SpawnNpc(val(Npc1), Npc1Pos, True, False)
mariano = 0
End Select

End Sub

Private Sub mnuCerrar_Click()


If MsgBox("¡¡Atencion!! Si cierra el servidor puede provocar la perdida de datos. ¿Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
Inet2.URL = "http://symxsoft.net/revival/online.php?num=OFFLINE"
Inet2.OpenURL
    Dim f
    For Each f In Forms
        Unload f
    Next
    
End If

End Sub

Private Sub mnusalir_Click()
    Call mnuCerrar_Click
End Sub

Public Sub mnuMostrar_Click()
On Error Resume Next
    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0
End Sub

Private Sub KillLog_Timer()
On Error Resume Next
If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\connect.log"
If FileExist(App.Path & "\logs\haciendo.log", vbNormal) Then Kill App.Path & "\logs\haciendo.log"
If FileExist(App.Path & "\logs\stats.log", vbNormal) Then Kill App.Path & "\logs\stats.log"
If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
If Not FileExist(App.Path & "\logs\nokillwsapi.txt") Then
    If FileExist(App.Path & "\logs\wsapi.log", vbNormal) Then Kill App.Path & "\logs\wsapi.log"
End If


End Sub

Private Sub mnuServidor_Click()
frmServidor.Visible = True
End Sub

Private Sub mnuSystray_Click()

Dim i As Integer
Dim s As String
Dim nid As NOTIFYICONDATA

s = "ARGENTUM-ONLINE"
nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, s)
i = Shell_NotifyIconA(NIM_ADD, nid)
    
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False

End Sub

Private Sub npcataca_Timer()

On Error Resume Next
Dim npc As Integer

For npc = 1 To LastNPC
    Npclist(npc).CanAttack = 1
Next npc

End Sub




'[MaTeO 13]
Private Sub RespawnNPC_Timer()
Dim ThePos As WorldPos
ThePos.Map = 113
ThePos.x = 72
ThePos.y = 74
If Hour(Time) = 5 And MazIndex = 0 Then
    Call SendData(SendTarget.toall, 0, 0, "¡El Calamar Gigante ha Emergido de las Profundidades! ~255~255~255~1~0")
    MazIndex = SpawnNpc(954, ThePos, True, False)
ElseIf Hour(Time) = 6 And MazIndex <> 0 Then
    Call SendData(SendTarget.toall, 0, 0, "¡El Calamar Gigante se ha Sumergido en las Profundidades! ~255~255~255~1~0")
    Call QuitarNPC(MazIndex)
    MazIndex = 0
End If
End Sub
'[/MaTeO 13]

Private Sub TIMER_AI_Timer()

On Error GoTo ErrorHandler
Dim NpcIndex As Integer
Dim x As Integer
Dim y As Integer
Dim UseAI As Integer
Dim mapa As Integer
Dim e_p As Integer

'Barrin 29/9/03
If Not haciendoBK And Not EnPausa Then
    'Update NPCs
    For NpcIndex = 1 To LastNPC
        
        If Npclist(NpcIndex).flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
                ''ia comun
                If Npclist(NpcIndex).flags.Paralizado = 1 Then
                      Call EfectoParalisisNpc(NpcIndex)
                Else
                     'Usamos AI si hay algun user en el mapa
                     If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
                        Call EfectoParalisisNpc(NpcIndex)
                     End If
                     mapa = Npclist(NpcIndex).pos.Map
                     If mapa > 0 Then
                          If MapInfo(mapa).NumUsers > 0 Then
                                  If Npclist(NpcIndex).Movement <> TipoAI.ESTATICO Then
                                        Call NPCAI(NpcIndex)
                                  End If
                          End If
                     End If
                     
                End If
            End If
    
    Next NpcIndex

End If


Exit Sub

ErrorHandler:
 Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).name & " mapa:" & Npclist(NpcIndex).pos.Map)
 Call MuereNpc(NpcIndex, 0)

End Sub

Private Sub Timer1_Timer()

On Error Resume Next
Dim i As Integer

For i = 1 To MaxUsers
    If UserList(i).flags.UserLogged Then _
        If UserList(i).flags.Oculto = 1 Then Call DoPermanecerOculto(i)
Next i

End Sub



Private Sub Timer2_Timer()
On Error Resume Next
asdff = asdff + 1

If asdff = 5 Then
frmMain.Inet1.URL = "http://www.symxsoft.net/revival/ranking.php?num=" & Ranking.MaxUsuariosMatados.UserName & " (" & Ranking.MaxUsuariosMatados.value & ")"
frmMain.Inet1.OpenURL
End If

If asdff = 6 Then
frmMain.Inet1.URL = "http://www.symxsoft.net/revival/ranking2.php?num=" & Ranking.MaxTrofeos.UserName & " (" & Ranking.MaxTrofeos.value & ")"
frmMain.Inet1.OpenURL
End If

If asdff = 7 Then

frmMain.Inet1.URL = "http://www.symxsoft.net/revival/ranking3.php?num=" & Ranking.MaxOro.UserName & " (" & Ranking.MaxOro.value & ")"
frmMain.Inet1.OpenURL
End If

If asdff = 8 Then
Inet1.URL = "http://www.symxsoft.net/revival/castillonix.php?num=" & AlmacenaDominadornix
Inet1.OpenURL
End If


If asdff = 9 Then
Inet1.URL = "http://www.symxsoft.net/revival/castilloulla.php?num=" & AlmacenaDominador
Inet1.OpenURL
End If

If asdff = 10 Then

Inet1.URL = "http://www.symxsoft.net/revival/fortaleza.php?num=" & Fortaleza
Inet1.OpenURL

End If
If asdff = 11 Then
Inet1.URL = "http://www.symxsoft.net/revival/castillolemuria.php?num=" & Lemuria
Inet1.OpenURL
End If
If asdff = 12 Then
Inet1.URL = "http://www.symxsoft.net/revival/castillotale.php?num=" & Tale
Inet1.OpenURL
asdff = 0
End If
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
Inet2.URL = "http://www.symxsoft.net/revival/user.php?num=" & NumUsers
Inet2.OpenURL
End Sub

Private Sub tmrResp_Timer()

'CHOTS | Npc de cada 6 horas
Dim Npc1 As Integer
Dim Npc1Pos As WorldPos
Npc1 = 604
Npc1Pos.Map = 56
Npc1Pos.x = 61
Npc1Pos.y = 38

'CHOTS | Npc de cada 8 horas
Dim Npc2 As Integer
Dim Npc2Pos As WorldPos
Npc2 = 607
Npc2Pos.Map = 50
Npc2Pos.x = 34
Npc2Pos.y = 48

ContReSpawnNpc = ContReSpawnNpc + 1

If ContReSpawnNpc = 350 Then
Call SendData(SendTarget.toall, 0, 0, "||RevivalAo> En 10 minutos renacerá el Espectro Infernal del dungeon del mapa 21." & FONTTYPE_SERVER)

ElseIf ContReSpawnNpc = 358 Then
Call SendData(SendTarget.toall, 0, 0, "||RevivalAo> En 2 minutos renacerá el Espectro Infernal del dungeon del mapa 21." & FONTTYPE_SERVER)

ElseIf ContReSpawnNpc = 360 Then
Call SpawnNpc(val(Npc1), Npc1Pos, True, False)
Call SendData(SendTarget.toall, 0, 0, "||RevivalAo> El ESPECTRO INFERNAL del dungeon del mapa 21, HA RENACIDO!!" & FONTTYPE_GUILD)
Call SendData(SendTarget.toall, 0, 0, "TW122")

ElseIf ContReSpawnNpc = 370 Then
Call SendData(SendTarget.toall, 0, 0, "||RevivalAo> En 10 minutos renacerá el Arcangel del dungeon del mapa 29." & FONTTYPE_SERVER)

ElseIf ContReSpawnNpc = 378 Then
Call SendData(SendTarget.toall, 0, 0, "||RevivalAo> En 2 minutos renacerá el Arcangel del dungeon del mapa 29." & FONTTYPE_SERVER)

ElseIf ContReSpawnNpc = 380 Then
Call SpawnNpc(val(Npc2), Npc2Pos, True, False)
Call SendData(SendTarget.toall, 0, 0, "||RevivalAo> El ARCANGEL del dungeon del mapa 29, HA RENACIDO!!" & FONTTYPE_GUILD)
Call SendData(SendTarget.toall, 0, 0, "TW122")
ContReSpawnNpc = 0
End If

End Sub

Private Sub torneos_Timer()

xao = xao + 1
Select Case xao
Case 84
Call SendData(SendTarget.toall, 0, 0, "||Torneo> En 10 minutos se realizará un torneo automatico." & FONTTYPE_GUILD)
Case 89
Call SendData(SendTarget.toall, 0, 0, "||Torneo> En 5 minutos se realizará un torneo automatico." & FONTTYPE_GUILD)
Case 93
Call SendData(SendTarget.toall, 0, 0, "||Torneo> En 1 minutos se realizará un torneo automatico." & FONTTYPE_GUILD)
Case 94
Call torneos_auto(3)
Case 96
If Torneo_Esperando = True Then
Call Torneoauto_Cancela
xao = 2
Else
xao = 2
End If
End Select
End Sub

Private Sub tPiqueteC_Timer()
On Error GoTo errhandler
Static Segundos As Integer
Dim NuevaA As Boolean
Dim NuevoL As Boolean
Dim GI As Integer

Segundos = Segundos + 6

Dim i As Integer

For i = 1 To LastUser
    If UserList(i).flags.UserLogged Then
            
            If MapData(UserList(i).pos.Map, UserList(i).pos.x, UserList(i).pos.y).trigger = eTrigger.ANTIPIQUETE Then
                    UserList(i).Counters.PiqueteC = UserList(i).Counters.PiqueteC + 1
                    Call SendData(SendTarget.toindex, i, 0, "Z39")
                    If UserList(i).Counters.PiqueteC > 23 Then
                            UserList(i).Counters.PiqueteC = 0
                            Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)
                    End If
            Else
                    If UserList(i).Counters.PiqueteC > 0 Then UserList(i).Counters.PiqueteC = 0
            End If

            'ustedes se preguntaran que hace esto aca?
            'bueno la respuesta es simple: el codigo de AO es una mierda y encontrar
            'todos los puntos en los cuales la alineacion puede cambiar es un dolor de
            'huevos, asi que lo controlo aca, cada 6 segundos, lo cual es razonable

            GI = UserList(i).GuildIndex
            If GI > 0 Then
                NuevaA = False
                NuevoL = False
                If Not modGuilds.m_ValidarPermanencia(i, True, NuevaA, NuevoL) Then
                    Call SendData(SendTarget.toindex, i, 0, "||Has sido expulsado del clan. ¡El clan ha sumado un punto de antifacción!" & FONTTYPE_GUILD)
                End If
                If NuevaA Then
                    Call SendData(SendTarget.ToGuildMembers, GI, 0, "||¡El clan ha pasado a tener alineación neutral!" & FONTTYPE_GUILD)
                    Call LogClanes("El clan cambio de alineacion!")
                End If
                If NuevoL Then
                    Call SendData(SendTarget.ToGuildMembers, GI, 0, "||¡El clan tiene un nuevo líder!" & FONTTYPE_GUILD)
                    Call LogClanes("El clan tiene nuevo lider!")
                End If
            End If

            If Segundos >= 18 Then
'                Dim nfile As Integer
'                nfile = FreeFile ' obtenemos un canal
'                Open App.Path & "\logs\maxpasos.log" For Append Shared As #nfile
'                Print #nfile, UserList(i).Counters.Pasos
'                Close #nfile
                If Segundos >= 18 Then UserList(i).Counters.Pasos = 0
            End If
            
    End If
Next i

If Segundos >= 18 Then Segundos = 0
   
Exit Sub

errhandler:
    Call LogError("Error en tPiqueteC_Timer " & Err.Number & ": " & Err.Description)
End Sub





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''USO DEL CONTROL TCPSERV'''''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#If UsarQueSocket = 3 Then

Private Sub TCPServ_Eror(ByVal Numero As Long, ByVal Descripcion As String)
    Call LogError("TCPSERVER SOCKET ERROR: " & Numero & "/" & Descripcion)
End Sub

Private Sub TCPServ_NuevaConn(ByVal ID As Long)
On Error GoTo errorHandlerNC

    ESCUCHADAS = ESCUCHADAS + 1
    Escuch.caption = ESCUCHADAS
    
    Dim i As Integer
    
    Dim NewIndex As Integer
    NewIndex = NextOpenUser
    
    If NewIndex <= MaxUsers Then
        'call logindex(NewIndex, "******> Accept. ConnId: " & ID)
        
        TCPServ.SetDato ID, NewIndex
        
        If aDos.MaxConexiones(TCPServ.GetIP(ID)) Then
            Call aDos.RestarConexion(TCPServ.GetIP(ID))
            Call ResetUserSlot(NewIndex)
            Exit Sub
        End If

        If NewIndex > LastUser Then LastUser = NewIndex

        UserList(NewIndex).ConnID = ID
        UserList(NewIndex).ip = TCPServ.GetIP(ID)
        UserList(NewIndex).ConnIDValida = True
        Set UserList(NewIndex).CommandsBuffer = New CColaArray
        
        For i = 1 To BanIps.Count
            If BanIps.Item(i) = TCPServ.GetIP(ID) Then
                Call ResetUserSlot(NewIndex)
                Exit Sub
            End If
        Next i

    Else
        Call CloseSocket(NewIndex, True)
        LogCriticEvent ("NEWINDEX > MAXUSERS. IMPOSIBLE ALOCATEAR SOCKETS")
    End If

Exit Sub

errorHandlerNC:
Call LogError("TCPServer::NuevaConexion " & Err.Description)
End Sub

Private Sub TCPServ_Close(ByVal ID As Long, ByVal MiDato As Long)
    On Error GoTo eh
    '' No cierro yo el socket. El on_close lo cierra por mi.
    'call logindex(MiDato, "******> Remote Close. ConnId: " & ID & " Midato: " & MiDato)
    Call CloseSocket(MiDato, False)
Exit Sub
eh:
    Call LogError("Ocurrio un error en el evento TCPServ_Close. ID/miDato:" & ID & "/" & MiDato)
End Sub

Private Sub TCPServ_Read(ByVal ID As Long, Datos As Variant, ByVal Cantidad As Long, ByVal MiDato As Long)
Dim t() As String
Dim LoopC As Long
Dim RD As String
On Error GoTo errorh
If UserList(MiDato).ConnID <> UserList(MiDato).ConnID Then
    Call LogError("Recibi un read de un usuario con ConnId alterada")
    Exit Sub
End If

RD = StrConv(Datos, vbUnicode)

'call logindex(MiDato, "Read. ConnId: " & ID & " Midato: " & MiDato & " Dato: " & RD)

UserList(MiDato).RDBuffer = UserList(MiDato).RDBuffer & RD

t = Split(UserList(MiDato).RDBuffer, ENDC)
If UBound(t) > 0 Then
    UserList(MiDato).RDBuffer = t(UBound(t))
    
    For LoopC = 0 To UBound(t) - 1
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        '%%% SI ESTA OPCION SE ACTIVA SOLUCIONA %%%
        '%%% EL PROBLEMA DEL SPEEDHACK          %%%
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        If ClientsCommandsQueue = 1 Then
            If t(LoopC) <> "" Then
                If Not UserList(MiDato).CommandsBuffer.Push(t(LoopC)) Then
                    Call LogError("Cerramos por no encolar. Userindex:" & MiDato)
                    Call CloseSocket(MiDato)
                End If
            End If
        Else ' no encolamos los comandos (MUY VIEJO)
              If UserList(MiDato).ConnID <> -1 Then
                Call HandleData(MiDato, t(LoopC))
              Else
                Exit Sub
              End If
        End If
    Next LoopC
End If
Exit Sub

errorh:
Call LogError("Error socket read: " & MiDato & " dato:" & RD & " userlogged: " & UserList(MiDato).flags.UserLogged & " connid:" & UserList(MiDato).ConnID & " ID Parametro" & ID & " error:" & Err.Description)

End Sub

#End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''FIN  USO DEL CONTROL TCPSERV'''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
