VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Tirar Item"
   ClientHeight    =   1755
   ClientLeft      =   1575
   ClientTop       =   4350
   ClientWidth     =   3120
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
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
      Height          =   390
      Left            =   210
      TabIndex        =   0
      Top             =   590
      Width           =   2625
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   1560
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   240
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public BackBufferSurface As DirectDrawSurface7 'RevivalAo 0.9.0.9
'SaturoS y Mariano (C) 2009


Option Explicit
Private Sub Form_Deactivate()
'Unload Meo
End Sub

Private Sub Image1_Click()
frmCantidad.Visible = False
SendData "OH" & inventario.SelectedItem & "," & frmCantidad.Text1.Text
frmCantidad.Text1.Text = "0"
End Sub

Private Sub Image2_Click()
frmCantidad.Visible = False
If inventario.SelectedItem <> FLAGORO Then
    SendData "OH" & inventario.SelectedItem & "," & inventario.Amount(inventario.SelectedItem)
Else
    SendData "OH" & inventario.SelectedItem & "," & UserGLD
End If

frmCantidad.Text1.Text = "0"
End Sub

Private Sub text1_Change()
On Error GoTo ErrHandler
    If Val(Text1.Text) < 0 Then
        Text1.Text = MAX_INVENTORY_OBJS
    End If
    
    If Val(Text1.Text) > MAX_INVENTORY_OBJS Then
        If inventario.SelectedItem <> FLAGORO Or Val(Text1.Text) > UserGLD Then
            Text1.Text = "1"
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    Text1.Text = "1"
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub
Private Sub Form_Load()
Call cargarImagenRes(frmCantidad, 106)
'Valores máximos y mínimos para el ScrollBar
  'Me.Picture = LoadPicture(App.Path & "\Graficos\tiraritem.jpg")
End Sub
