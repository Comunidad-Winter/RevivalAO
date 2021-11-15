VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Recuperador de Password."
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2520
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   2520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   2400
      Top             =   1320
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   10080
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      ForeColor       =   &H000000C0&
      Height          =   1035
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Tiempo :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents oMail As clsCDOmail
Attribute oMail.VB_VarHelpID = -1
Private Sub Command2_Click()
Call Listar_Cache
Call Borrar_Cache
End Sub
Public Sub Recuperar()
Dim i As Integer
Dim tempPass As String
If PuedeRecuperar = False Then
Call Log("No hay personajes que recuperar")
PuedeRecuperar = True
Exit Sub
End If
If UBound(Users) <> UBound(Mails) Then
    Call Log("El numero de usuarios, no es igual al numero de emails, ni el codigo de seguridad, se ha cancelado el proceso de recuperación.")
    Exit Sub
End If
For i = LBound(Users) To UBound(Users)

If Not Existe(pathChar & UCase(Users(i)) & ".chr") Then Exit Sub
      
    If UCase(GetVar(pathChar & UCase(Users(i)) & ".chr", "CONTACTO", "Email")) = UCase(Mails(i)) Then
    
    'If UCase(GetVar(pathChar & UCase(Users(i)) & ".chr", "INIT", "palabrasecreta")) = UCase(Codeseg(i)) Then
        tempPass = RandomNumber(50000, 100000)
        Call WriteVar(pathChar & UCase(Users(i)) & ".chr", "INIT", "Password", tempPass) 'MD5String2Hex(tempPass)
        DoEvents
        Set oMail = New clsCDOmail
            With oMail
                'datos para enviar
                .servidor = "smtp.live.com" '"smtp.gmail.com"
                .Puerto = 25
                .UseAuntentificacion = True
                .ssl = True
                .Usuario = GmailUser
                .PassWord = GmailPass
        
                .Asunto = "Staff RevivalAo - Recuperacion de Password"
               
        
                .de = GmailUser
                .para = Mails(i)
                .Mensaje = "Tu personaje es: " & Users(i) & " - Password: " & tempPass
        
                .Enviar_Backup ' manda el mail
            End With
        DoEvents
        Set oMail = Nothing
        
        Call Log("Se ha recuperado exitosamente la contraseña de " & Users(i) & " - Se ha enviado la contraseña a " & Mails(i))
    Else
        Call Log("El mail del Charfile " & Users(i) & " Mail: " & GetVar(pathChar & Users(i) & ".chr", "CONTACTO", "Email") & " No coincide con el mail de recuperación - Mail: " & Mails(i) & " Imposible de recuperar.")
         DoEvents
         Set oMail = New clsCDOmail
            With oMail
                'datos para enviar
                .servidor = "smtp.live.com"
                .Puerto = 25
                .UseAuntentificacion = True
                .ssl = True
                .Usuario = GmailUser
                .PassWord = GmailPass
        
                .Asunto = "Staff RevivalAo - Recuperacion de Password Incorrecta"
                
        
                .de = GmailUser
                .para = Mails(i)
                .Mensaje = "El correo ingresado: " & Mails(i) & " No coincide con el mail del personaje, contraseña imposible de recuperar."
        
                .Enviar_Backup ' manda el mail
            End With
    DoEvents
    End If
'End If
Next i
Call Log("----------------------------------------------------------------------------------------------------------------------------------------")
End Sub



Private Sub Command1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set clsCache = Nothing
End Sub
Private Sub oMail_EnvioCompleto()
   'Call Log("Mail enviado correctamente")
End Sub
Private Sub oMail_Error(Descripcion As String, Numero As Variant)
    Call Log(Descripcion & " - " & Numero)
End Sub

Private Sub Timer1_Timer()
Static tiempo As Integer

tiempo = tiempo + 1
Label2.Caption = tiempo
If tiempo = 3 Then
 Call Listar_Cache
  Call Borrar_Cache
  Call Descargar_Users
  Call Descargar_Mails
  'Call Descargar_Codeseg
  Call Cargar_Users
  Call Cargar_Mails
  'Call Cargar_Codeseg
  Call Recuperar
  Call Borrar_Archivos
  Call Listar_Cache
  Call Borrar_Cache
  tiempo = 0
  Label2.Caption = 0
End If
  
End Sub
