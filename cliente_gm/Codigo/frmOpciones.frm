VERSION 5.00
Begin VB.Form frmOpciones 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Opciones"
   ClientHeight    =   5235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3240
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
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   5235
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton FPS 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton FPS 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   4080
      Width           =   255
   End
   Begin VB.OptionButton FPS 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   3720
      Width           =   255
   End
   Begin VB.OptionButton FPS 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "a"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   9120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   240
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   240
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   240
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   240
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   240
      Top             =   840
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   240
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdAlphaB_Click()
If ConAlfaB = True Then
ConAlfaB = False
cmdAlphaB.Caption = "AlphaBlending Desactivado"
Else
ConAlfaB = True
cmdAlphaB.Caption = "AlphaBlending Activado"
End If
End Sub

Private Sub cmdBalance_Click(index As Integer)

End Sub



Private Sub cmdCerrar_Click()

End Sub

Private Sub CmdMapa_Click(index As Integer)

End Sub

Private Sub cmdMsn_Click()

End Sub

Private Sub cmdMusica_Click()
       
End Sub
Private Sub cmdSound_Click()

End Sub

Private Sub CmdUclick_Click()
If Uclickear = True Then
Uclickear = False
CmdUclick.Caption = "U+Click Boton derecho Desactivado"
Else
Uclickear = True
CmdUclick.Caption = "U+Click Boton derecho Activado"
End If
End Sub

Private Sub Command1_Click()

End Sub


Private Sub Command4_Click()

End Sub

Private Sub Form_Load()
Call cargarImagenRes(Me, 128)
'Me.Picture = LoadPicture(App.Path & _
'    "\Graficos\Opciones.jpg")
    '[MaTeO 11]
    FPS(IndexSet).value = True
    '[/MaTeO 11]
'Valores máximos y mínimos para el ScrollBar
  
    If Musica Then
        'cmdMusica.Caption = "Musica Activada"
    Else
        'cmdMusica.Caption = "Musica Desactivada"
    End If
    
    If Sound Then
        'cmdSound.Caption = "Sonidos Activados"
    Else
        'cmdSound.Caption = "Sonidos Desactivados"
    End If
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)

End Sub
'[MaTeO 11] ¡¡Acordarse de agregar los OptionButton!!
Private Sub FPS_Click(index As Integer)
    Call LimitarFPS(index)
    Call SaveClientSetup
End Sub
'[/MaTeO]

Private Sub Label2_Click()
Dim web As Long
web = ShellExecute(Me.hwnd, "open", "http://www.RevivalAo.com.ar", "", "", 1)
End Sub

Private Sub Label3_Click()
Dim web As Long
web = ShellExecute(Me.hwnd, "open", "http://www.RevivalAo.com.ar/", "", "", 1)
End Sub

Private Sub Label4_Click()
Dim web As Long
web = ShellExecute(Me.hwnd, "open", "http://www.RevivalAo.com.ar", "", "", 1)
End Sub

Private Sub Label5_Click()
Dim web As Long
web = ShellExecute(Me.hwnd, "open", "http://www.RevivalAo.com.ar", "", "", 1)
End Sub

Private Sub Label6_Click()
Dim web As Long
web = ShellExecute(Me.hwnd, "open", "http://www.RevivalAo.com.ar", "", "", 1)
End Sub

Private Sub Label7_Click()
Dim web As Long
web = ShellExecute(Me.hwnd, "open", "http://www.RevivalAo.com.ar", "", "", 1)
End Sub

Private Sub Image1_Click()
        If Sound Then
            Sound = False
            MsgBox "Sonidos Desactivados"
            Call Audio.StopWave
            RainBufferIndex = 0
            frmMain.IsPlaying = PlayLoop.plNone
        Else
            Sound = True
            MsgBox "Sonidos Activados"
        End If
        ClientSetup.bNoSound = Not Sound
        Call SaveClientSetup
End Sub

Private Sub Image2_Click()
        If Musica Then
            Musica = False
           MsgBox "Musica Desactivada"
            Audio.StopMidi
        Else
            Musica = True
           MsgBox "Musica Activada"
            Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
        End If
        ClientSetup.bNoMusic = Not Music
End Sub

Private Sub Image3_Click()
Call frmMapa.Show(vbModeless, frmMain)
End Sub

Private Sub Image4_Click()
If Centrada = False Then
frmMain.Top = (Screen.Height - frmMain.Height) / 2
frmMain.Left = (Screen.Width - frmMain.Width) / 2
MsgBox "Vuelve a hacer click para volver a la posicion original"
Centrada = True
Else
frmMain.Top = 0
frmMain.Left = 0
Centrada = False
End If
End Sub

Private Sub Image5_Click()
Call frmContra.Show(vbModeless, frmMain)
End Sub

Private Sub Image6_Click()
Call SaveClientSetup
Unload Me
End Sub
