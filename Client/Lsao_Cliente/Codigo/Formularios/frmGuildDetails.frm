VERSION 5.00
Begin VB.Form frmGuildDetails 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Detalles del Clan"
   ClientHeight    =   7755
   ClientLeft      =   -60
   ClientTop       =   -165
   ClientWidth     =   7455
   ClipControls    =   0   'False
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
   ScaleHeight     =   7755
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   610
      TabIndex        =   10
      Top             =   3970
      Width           =   6250
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   610
      TabIndex        =   9
      Top             =   4380
      Width           =   6255
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   610
      TabIndex        =   8
      Top             =   4790
      Width           =   6250
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   610
      TabIndex        =   7
      Top             =   5190
      Width           =   6255
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   610
      TabIndex        =   6
      Top             =   5600
      Width           =   6255
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   610
      TabIndex        =   5
      Top             =   6000
      Width           =   6255
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   610
      TabIndex        =   4
      Top             =   6400
      Width           =   6255
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   8880
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000009&
      Height          =   1575
      Left            =   610
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   900
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
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
      Index           =   1
      Left            =   4800
      MouseIcon       =   "frmGuildDetails.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
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
      Index           =   0
      Left            =   600
      MouseIcon       =   "frmGuildDetails.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7080
      Width           =   2055
   End
End
Attribute VB_Name = "frmGuildDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private Sub Command1_Click(index As Integer)
Select Case index

Case 0
    Unload Me
Case 1
    Dim fdesc$
    fdesc$ = Replace(txtDesc, vbCrLf, "º", , , vbBinaryCompare)
    
'    If Not AsciiValidos(fdesc$) Then
'        MsgBox "La descripcion contiene caracteres invalidos"
'        Exit Sub
'    End If
    
    Dim k As Integer
    Dim Cont As Integer
    Cont = 0
    For k = 0 To txtCodex1.UBound
'        If Not AsciiValidos(txtCodex1(k)) Then
'            MsgBox "El codex tiene invalidos"
'            Exit Sub
'        End If
        If Len(txtCodex1(k).Text) > 0 Then Cont = Cont + 1
    Next k
    If Cont < 4 Then
            MsgBox "Debes definir al menos cuatro mandamientos."
            Exit Sub
    End If
    
    Dim chunk$
    
    If CreandoClan Then
        chunk$ = "CIG" & fdesc$
        chunk$ = chunk$ & "¬" & ClanName & "¬" & Site & "¬" & Cont
    Else
        chunk$ = "DESCOD" & fdesc$ & "¬" & Cont
    End If
    
    
    
    For k = 0 To txtCodex1.UBound
        chunk$ = chunk$ & "¬" & txtCodex1(k)
    Next k
    
    
    Call SendData(chunk$)
    
    CreandoClan = False
    
    Unload Me
    
End Select



End Sub

Private Sub Form_Deactivate()

'If Not frmGuildLeader.Visible Then
'    Me.SetFocus
'Else
'    'Unload Me
'End If
'
End Sub

Private Sub Form_Load()
Call cargarImagenRes(frmGuildDetails, 118)
'Me.Picture = LoadPicture(App.Path & _
   ' "\Graficos\GuildDetails.jpg")
End Sub

