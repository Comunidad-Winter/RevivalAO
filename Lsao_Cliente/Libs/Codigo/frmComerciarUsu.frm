VERSION 5.00
Begin VB.Form frmComerciarUsu 
   BorderStyle     =   0  'None
   ClientHeight    =   7305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   487
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   463
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optQue 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   5880
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.OptionButton optQue 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   5040
      TabIndex        =   6
      Top             =   1200
      Value           =   -1  'True
      Width           =   195
   End
   Begin VB.TextBox txtCant 
      Height          =   285
      Left            =   4920
      TabIndex        =   5
      Text            =   "1"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Left            =   3915
      TabIndex        =   3
      Top             =   1425
      Width           =   2490
   End
   Begin VB.ListBox List2 
      Height          =   3960
      Left            =   540
      TabIndex        =   2
      Top             =   1425
      Width           =   2490
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   1260
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   0
      Top             =   225
      Width           =   540
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   2520
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   480
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   1800
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4200
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   1440
      TabIndex        =   4
      Top             =   5445
      Width           =   90
   End
   Begin VB.Label lblEstadoResp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Esperando respuesta..."
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
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   2490
   End
End
Attribute VB_Name = "frmComerciarUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()

End Sub

Private Sub cmdOfrecer_Click()


End Sub

Private Sub cmdRechazar_Click()

End Sub

Private Sub Command2_Click()


End Sub

Private Sub Form_Deactivate()
'Me.SetFocus
'Picture1.SetFocus


End Sub


Private Sub Form_Load()
'Valores máximos y mínimos para el ScrollBar
   Call cargarImagenRes(frmComerciarUsu, 111)
'Carga las imagenes...?
lblEstadoResp.Visible = False

 'Me.Picture = LoadPicture(App.Path & "\Graficos\ComerciarUsu.jpg")
End Sub

Private Sub Form_LostFocus()
Me.SetFocus
Picture1.SetFocus

End Sub

Private Sub Image1_Click()

If optQue(0).value = True Then
    If List1.ListIndex < 0 Then Exit Sub
    If List1.ItemData(List1.ListIndex) <= 0 Then Exit Sub
    
'    If Val(txtCant.Text) > List1.ItemData(List1.ListIndex) Or _
'        Val(txtCant.Text) <= 0 Then Exit Sub
ElseIf optQue(1).value = True Then
'    If Val(txtCant.Text) > UserGLD Then
'        Exit Sub
'    End If
End If

If optQue(0).value = True Then
    Call SendData("OFRECER" & List1.ListIndex + 1 & "," & Trim(Val(txtCant.Text)))
ElseIf optQue(1).value = True Then
    Call SendData("OFRECER" & FLAGORO & "," & Trim(Val(txtCant.Text)))
Else
    Exit Sub
End If

lblEstadoResp.Visible = True
End Sub

Private Sub Image2_Click()
Call SendData("COMUSUNO")
End Sub

Private Sub Image3_Click()
Call SendData("COMUSUOK")
End Sub

Private Sub Image4_Click()
Call SendData("FINCOMUSU")
End Sub

Private Sub list1_Click()
DibujaGrh inventario.GrhIndex(List1.ListIndex + 1)

End Sub

Public Sub DibujaGrh(Grh As Integer)
Dim SR As RECT, dr As RECT

SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.Bottom = 32

dr.Left = 0
dr.Top = 0
dr.Right = 32
dr.Bottom = 32

Call DrawGrhtoHdc(Picture1.hwnd, Picture1.hdc, Grh, SR, dr)

End Sub

Private Sub List2_Click()
If List2.ListIndex >= 0 Then
    DibujaGrh OtroInventario(List2.ListIndex + 1).GrhIndex
    Label3.Caption = List2.ItemData(List2.ListIndex)
   Image3.Enabled = True
   Image2.Enabled = True
Else
    Image3.Enabled = False
    Image2.Enabled = False
End If

End Sub

Private Sub optQue_Click(index As Integer)
Select Case index
Case 0
    List1.Enabled = True
Case 1
    List1.Enabled = False
End Select

End Sub

Private Sub txtCant_Change()
    If Val(txtCant.Text) < 1 Then txtCant.Text = "1"
End Sub

Private Sub txtCant_KeyDown(KeyCode As Integer, Shift As Integer)
If Not ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = vbKeyBack Or _
        KeyCode = vbKeyDelete Or (KeyCode >= 37 And KeyCode <= 40)) Then
    'txtCant = KeyCode
    KeyCode = 0
End If

End Sub

Private Sub txtCant_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or _
        KeyAscii = vbKeyDelete Or (KeyAscii >= 37 And KeyAscii <= 40)) Then
    'txtCant = KeyCode
    KeyAscii = 0
End If

End Sub

'[/Alejo]

