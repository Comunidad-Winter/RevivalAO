VERSION 5.00
Begin VB.Form frmOldPersonaje 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Argentum"
   ClientHeight    =   3765
   ClientLeft      =   5430
   ClientTop       =   2250
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   251
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2265
      TabIndex        =   0
      Top             =   705
      Width           =   4530
   End
   Begin VB.TextBox PasswordTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
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
      ForeColor       =   &H000000FF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2265
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   4530
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   960
      Left            =   510
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   6120
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   4920
      MouseIcon       =   "frmOldPersonaje.frx":0000
      MousePointer    =   99  'Custom
      Top             =   3090
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   1365
      MouseIcon       =   "frmOldPersonaje.frx":0152
      MousePointer    =   99  'Custom
      Top             =   3105
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   2
      Left            =   3120
      MouseIcon       =   "frmOldPersonaje.frx":02A4
      MousePointer    =   99  'Custom
      Top             =   3090
      Width           =   960
   End
End
Attribute VB_Name = "frmOldPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Const textoKeypad = "Utilice el teclado como protecci�n contra keyloggers. Seleccione el password con el mouse y presione <ENTER> al finalizar"
Private Const textoSeguir = "Conectarse al juego" & vbNewLine & "con el usuario y" & vbNewLine & "clave seleccionadas"
Private Const textoSalir = "Volver a la pantalla principal" & vbNewLine & "para crear personajes o recuperar" & vbNewLine & "contrase�as"



Private Sub Form_Load()
Dim j
For Each j In Image1()
    j.Tag = "0"
Next

NameTxt.Text = ""
PasswordTXT.Text = ""
'Me.Picture = LoadPicture(App.Path & "\Graficos\oldcaracter.jpg")
'Image1(1).Picture = LoadPicture(App.Path & "\Graficos\bvolver.jpg")
'Image1(0).Picture = LoadPicture(App.Path & "\Graficos\bsiguiente.jpg")
'Image1(2).Picture = LoadPicture(App.Path & "\Graficos\bteclas.jpg")



End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1(0).Tag = "1" Then
            Me.lblInfo.Visible = False
            Me.lblInfo.Caption = vbNullString
            Image1(0).Tag = "0"
            Image1(0).Picture = LoadPicture(App.Path & "\Graficos\bsiguiente.jpg")
End If
If Image1(1).Tag = "1" Then
            Me.lblInfo.Visible = False
            Me.lblInfo.Caption = vbNullString
            Image1(1).Tag = "0"
            Image1(1).Picture = LoadPicture(App.Path & "\Graficos\bvolver.jpg")
End If
If Image1(2).Tag = "1" Then
            Me.lblInfo.Visible = False
            Me.lblInfo.Caption = vbNullString
            Image1(2).Tag = "0"
            Image1(2).Picture = LoadPicture(App.Path & "\Graficos\bteclas.jpg")
End If

End Sub

Private Sub Image1_Click(index As Integer)

Call Audio.PlayWave(SND_CLICK)



Select Case index
    Case 0
       
#If UsarWrench = 1 Then
        If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
#Else
        If frmMain.Winsock1.State <> sckClosed Then _
            frmMain.Winsock1.Close
#End If
        If frmConnect.MousePointer = 11 Then
            Exit Sub
        End If
        
        
        'update user info
        UserName = NameTxt.Text
        Dim aux As String
        aux = PasswordTXT.Text
#If SeguridadAlkon Then
        UserPassword = md5.GetMD5String(aux)
        Call md5.MD5Reset
#Else
        UserPassword = aux
#End If
        If CheckUserData(False) = True Then
            'SendNewChar = False
            EstadoLogin = Normal
            Me.MousePointer = 11
#If UsarWrench = 1 Then
            frmMain.Socket1.HostAddress = CurServerIp
            frmMain.Socket1.RemotePort = CurServerPort
            frmMain.Socket1.Connect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
            frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If
        End If
        
    Case 1
        Me.Visible = False
    Case 2
        Load frmKeypad
        frmKeypad.Show vbModal
        Unload frmKeypad
        Me.PasswordTXT.SetFocus
        
End Select
End Sub

Private Sub Image1_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case index
    Case 0
        If Image1(0).Tag = "0" Then
            Me.lblInfo.Visible = True
            Me.lblInfo.Caption = textoSeguir
            Image1(0).Tag = "1"
            Call Audio.PlayWave(SND_OVER)
            Image1(0).Picture = LoadPicture(App.Path & "\Graficos\bsiguientea.jpg")
        End If
    Case 1
        If Image1(1).Tag = "0" Then
            Me.lblInfo.Visible = True
            Me.lblInfo.Caption = textoSalir
            Image1(1).Tag = "1"
            Call Audio.PlayWave(SND_OVER)
            Image1(1).Picture = LoadPicture(App.Path & "\Graficos\bvolvera.jpg")
        End If
    Case 2
        If Image1(2).Tag = "0" Then
            Me.lblInfo.Visible = True
            Me.lblInfo.Caption = textoKeypad
            Image1(2).Tag = "1"
            Call Audio.PlayWave(SND_OVER)
            Image1(2).Picture = LoadPicture(App.Path & "\Graficos\bteclasa.jpg")
        End If
        
End Select
End Sub

Private Sub PasswordTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Image1_Click(0)
    End If
End Sub
