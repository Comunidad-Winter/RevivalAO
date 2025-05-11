VERSION 5.00
Begin VB.Form frmAoDefender 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmAoDefender.frx":0000
   ScaleHeight     =   3000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   4920
      Top             =   240
   End
End
Attribute VB_Name = "frmAoDefender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Unload Me
End Sub
