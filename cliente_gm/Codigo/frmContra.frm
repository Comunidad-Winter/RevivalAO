VERSION 5.00
Begin VB.Form frmContra 
   BorderStyle     =   0  'None
   Caption         =   "Cambiar Contaseña"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   200
      TabIndex        =   0
      Top             =   1685
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   1320
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   2720
      Width           =   3375
   End
End
Attribute VB_Name = "frmContra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Call cargarImagenRes(frmContra, 112)
'  Me.Picture = LoadPicture(App.Path & "\Graficos\CAMBIAPASS.jpg")
End Sub

Private Sub Image1_Click()
If Text1.Text = "" Then
Label1.Caption = "Debe escribir una contraseña!"
Exit Sub
End If
 Call SendData("/PASSWD " & Text1.Text)
 Label1.Caption = "Contraseña cambiada correctamente."
 Text1.Text = ""
 Unload frmContra
End Sub

