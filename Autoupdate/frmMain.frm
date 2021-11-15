VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   ClientHeight    =   2340
   ClientLeft      =   -60
   ClientTop       =   -165
   ClientWidth     =   7395
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":231D7
   ScaleHeight     =   2340
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Left            =   4680
      Top             =   4080
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   1680
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7440
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lEstado 
      BackStyle       =   0  'Transparent
      Caption         =   "Esperando sincronización."
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   840
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   9360
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tamaño del archivo:"
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "A:"
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1965
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "De:"
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   9120
      TabIndex        =   4
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lURL 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   9360
      TabIndex        =   3
      Top             =   3240
      Width           =   5655
   End
   Begin VB.Line Line1 
      X1              =   2520
      X2              =   2520
      Y1              =   5160
      Y2              =   6765
   End
   Begin VB.Label lDirectorio 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   6120
      Width           =   5655
   End
   Begin VB.Label lSize 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   1080
      Width           =   5655
   End
   Begin VB.Line Line2 
      X1              =   5040
      X2              =   12840
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1575
      Left            =   4080
      Shape           =   4  'Rounded Rectangle
      Top             =   4320
      Visible         =   0   'False
      Width           =   7815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* UpdateInteligente v4.0 *
'**************************
' Contacto: (dudas o cualquier cosa)
'   MSN/MAIL: shedark@live.com.ar
'   GSZone: www.gs-zone.com.ar, mensaje privado a Shed
' Configuracion:
'   Leer manual adjunto al código
' Nuevo:
'   Código reescrito y simplificado, adaptandolo a las únicas necesidades del programa
'   Posibilidad de elegir que se creen los links automaticamente (EJ: http://host/Parche1.zip) o _
    redirigir hacia un link elegido por ustedes, puede ser cualquiera (pero debe ubicarse en EJ: http://host/Link1.txt) _
    Esto se cambia llendo a Proyecto > Propiedades del proyecto > Generar > BuscarLinks = (0 o 1). Por defecto automático (0).
'   Nueva forma de descarga de archivos más efectiva y que nos permite informar, a medida que se realiza la descarga, _
    el tamaño del archivo descargado, su ubicacion, host y nombre.
'   Nueva forma de escritura y lectura de archivos (destinado unicamente a la búsqueda del Integer del número de actualización)
'   La progressbar nos indica un porcentaje preciso del tamaño del archivo
'   Eliminación de elementos que quedaron en deshuso
' Bugs:
'   En caso de encontrar un error enviar un e-mail o MP (ver Contacto) con:
'       - Una imágen del error (en modo depuración si es posible)
'       - Modificaciones del código (incluyendo links modificados)
'   e intentaré responder cuanto antes
' Los créditos del código del programa corresponden a SHEDARK (Shed)
' AVISO: MANTENTE AL TANTO, NUEVAS VERSIONES MÁS AUTOMÁTICAS

Option Explicit

Rem Programado por Shedark

Dim Directory As String, bDone As Boolean, dError As Boolean, F As Integer
        


Public Sub Analizar()
    Dim i As Integer, iX As Integer, tX As Integer, DifX As Integer, dNum As String
    
    lEstado.Caption = "Obteniendo datos..."
    
    iX = Inet1.OpenURL("http://www.revivalao.com/Parches/VEREXE.txt") 'Host
    tX = LeerInt(App.Path & "\Recursos\Update.revival")
    DifX = iX - tX
    
    If Not (DifX = 0) Then
    frmMain.Visible = True
        For i = 1 To DifX
            Inet1.AccessType = icUseDefault
            dNum = i + tX
            
            #If BuscarLinks Then 'Buscamos el link en el host (1)
                Inet1.URL = Inet1.OpenURL("http://tuhost/Link" & dNum & ".txt") 'Host
            #Else                'Generamos Link por defecto (0)
                Inet1.URL = "http://www.revivalao.com/Parches/Parche" & dNum & ".zip" 'Host
            #End If
            
            Directory = App.Path & "\Recursos\Parche" & dNum & ".zip"
            bDone = False
            dError = False
            
            lURL.Caption = Inet1.URL
            lName.Caption = "Parche" & dNum & ".zip"
            lDirectorio.Caption = App.Path & "\"
                
            frmMain.Inet1.Execute , "GET"
            
            Do While bDone = False
            DoEvents
            Loop
            
            If dError Then Exit Sub
            
            Unzip Directory, App.Path & "\"
            Kill Directory
        Next i
    End If
     
    Call GuardarInt(App.Path & "\Recursos\Update.revival", iX)
    
 
    lEstado.Caption = "Cliente actualizado correctamente."
    Shell (App.Path & "\Libs\RevivalAo.exe")
    frmMain.Visible = False
    End
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    Select Case State
        Case icError
            lEstado.Caption = "Error en la coneccion, descarga abortada."
            bDone = True
            dError = True
        Case icResponseCompleted
            Dim vtData As Variant
            Dim tempArray() As Byte
            Dim FileSize As Long
            
            FileSize = Inet1.GetHeader("Content-length")
            ProgressBar1.Max = FileSize
            
            lEstado.Caption = "Descarga iniciada."
            
            Open Directory For Binary Access Write As #1
                vtData = Inet1.GetChunk(1024, icByteArray)
                DoEvents
                
                
                Do While Not Len(vtData) = 0
                    tempArray = vtData
                    Put #1, , tempArray
                    
                vtData = Inet1.GetChunk(1024, icByteArray)
                    
                    ProgressBar1.Value = ProgressBar1.Value + Len(vtData) * 2
                    lSize.Caption = ProgressBar1.Value & "bytes de " & FileSize & "bytes"

                    DoEvents
                Loop
            Close #1
            
            lEstado.Caption = "Descarga finalizada."
            lSize.Caption = FileSize & "bytes"
            ProgressBar1.Value = 0
            
            bDone = True
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Function LeerInt(ByVal Ruta As String) As Integer
    F = FreeFile
    Open Ruta For Input As F
    LeerInt = Input$(LOF(F), #F)
    Close #F
End Function

Private Sub GuardarInt(ByVal Ruta As String, ByVal data As Integer)
    F = FreeFile
    Open Ruta For Output As F
    Print #F, data
    Close #F
End Sub

