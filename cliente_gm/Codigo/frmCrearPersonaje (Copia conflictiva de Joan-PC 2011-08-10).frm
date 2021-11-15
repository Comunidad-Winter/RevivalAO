VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   FillColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   8160
      TabIndex        =   44
      Top             =   6855
      Width           =   3495
   End
   Begin VB.TextBox txtPasswdCheck 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   8160
      MaxLength       =   25
      PasswordChar    =   "*"
      TabIndex        =   37
      Top             =   3000
      Width           =   3480
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   8160
      MaxLength       =   25
      PasswordChar    =   "*"
      TabIndex        =   36
      Top             =   2160
      Width           =   3525
   End
   Begin VB.TextBox txtCorreoCheck 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   8085
      MaxLength       =   50
      TabIndex        =   35
      Top             =   4980
      Width           =   3525
   End
   Begin VB.TextBox txtCorreo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   8205
      MaxLength       =   50
      TabIndex        =   34
      Top             =   3960
      Width           =   3405
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "frmCrearPersonaje.frx":0000
      Left            =   1800
      List            =   "frmCrearPersonaje.frx":0037
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   6120
      Width           =   1890
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":00D1
      Left            =   1800
      List            =   "frmCrearPersonaje.frx":00DB
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   7230
      Width           =   1890
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":00EE
      Left            =   1800
      List            =   "frmCrearPersonaje.frx":0101
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   6690
      Width           =   1890
   End
   Begin VB.ComboBox lstHogar 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":012E
      Left            =   1800
      List            =   "frmCrearPersonaje.frx":0135
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   7785
      Width           =   1890
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   8160
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1200
      Width           =   3525
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8085
      TabIndex        =   43
      Top             =   5925
      Width           =   3495
   End
   Begin VB.Label lblPlusFuerza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   3165
      TabIndex        =   42
      Top             =   1650
      Width           =   330
   End
   Begin VB.Label lblPlusConstitucion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   3165
      TabIndex        =   41
      Top             =   3480
      Width           =   330
   End
   Begin VB.Label lblPlusCarisma 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   3165
      TabIndex        =   40
      Top             =   3000
      Width           =   330
   End
   Begin VB.Label lblPlusInteligencia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   3165
      TabIndex        =   39
      Top             =   2520
      Width           =   330
   End
   Begin VB.Label lblPlusAgilidad 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   3165
      TabIndex        =   38
      Top             =   2085
      Width           =   330
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "+3"
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   5400
      TabIndex        =   33
      Top             =   8640
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   8760
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   6870
      TabIndex        =   32
      Top             =   8220
      Width           =   375
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   3
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":0145
      MousePointer    =   99  'Custom
      Top             =   1320
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   5
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":0297
      MousePointer    =   99  'Custom
      Top             =   1680
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   7
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":03E9
      MousePointer    =   99  'Custom
      Top             =   1920
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   9
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":053B
      MousePointer    =   99  'Custom
      Top             =   2280
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   11
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":068D
      MousePointer    =   99  'Custom
      Top             =   2640
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   13
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":07DF
      MousePointer    =   99  'Custom
      Top             =   2880
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   15
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":0931
      MousePointer    =   99  'Custom
      Top             =   3240
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   17
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":0A83
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   19
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":0BD5
      MousePointer    =   99  'Custom
      Top             =   3960
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   21
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":0D27
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   23
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":0E79
      MousePointer    =   99  'Custom
      Top             =   4560
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   25
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":0FCB
      MousePointer    =   99  'Custom
      Top             =   4920
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   27
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":111D
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   1
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":126F
      MousePointer    =   99  'Custom
      Top             =   960
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   0
      Left            =   6600
      MouseIcon       =   "frmCrearPersonaje.frx":13C1
      MousePointer    =   99  'Custom
      Top             =   960
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   2
      Left            =   6480
      MouseIcon       =   "frmCrearPersonaje.frx":1513
      MousePointer    =   99  'Custom
      Top             =   1320
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   4
      Left            =   6480
      MouseIcon       =   "frmCrearPersonaje.frx":1665
      MousePointer    =   99  'Custom
      Top             =   1680
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   270
      Index           =   6
      Left            =   6480
      MouseIcon       =   "frmCrearPersonaje.frx":17B7
      MousePointer    =   99  'Custom
      Top             =   1920
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   270
      Index           =   8
      Left            =   6600
      MouseIcon       =   "frmCrearPersonaje.frx":1909
      MousePointer    =   99  'Custom
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   10
      Left            =   6480
      MouseIcon       =   "frmCrearPersonaje.frx":1A5B
      MousePointer    =   99  'Custom
      Top             =   2640
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   6600
      MouseIcon       =   "frmCrearPersonaje.frx":1BAD
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   240
      Index           =   14
      Left            =   6600
      MouseIcon       =   "frmCrearPersonaje.frx":1CFF
      MousePointer    =   99  'Custom
      Top             =   3240
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   240
      Index           =   16
      Left            =   6480
      MouseIcon       =   "frmCrearPersonaje.frx":1E51
      MousePointer    =   99  'Custom
      Top             =   3960
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   240
      Index           =   18
      Left            =   6480
      MouseIcon       =   "frmCrearPersonaje.frx":1FA3
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   20
      Left            =   6480
      MouseIcon       =   "frmCrearPersonaje.frx":20F5
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   22
      Left            =   6600
      MouseIcon       =   "frmCrearPersonaje.frx":2247
      MousePointer    =   99  'Custom
      Top             =   4560
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   24
      Left            =   6600
      MouseIcon       =   "frmCrearPersonaje.frx":2399
      MousePointer    =   99  'Custom
      Top             =   4920
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   240
      Index           =   26
      Left            =   6480
      MouseIcon       =   "frmCrearPersonaje.frx":24EB
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   270
      Index           =   28
      Left            =   6600
      MouseIcon       =   "frmCrearPersonaje.frx":263D
      MousePointer    =   99  'Custom
      Top             =   5520
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   29
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":278F
      MousePointer    =   99  'Custom
      Top             =   5520
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   30
      Left            =   6600
      MouseIcon       =   "frmCrearPersonaje.frx":28E1
      MousePointer    =   99  'Custom
      Top             =   5880
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   31
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":2A33
      MousePointer    =   99  'Custom
      Top             =   5880
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   32
      Left            =   6600
      MouseIcon       =   "frmCrearPersonaje.frx":2B85
      MousePointer    =   99  'Custom
      Top             =   6240
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   33
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":2CD7
      MousePointer    =   99  'Custom
      Top             =   6120
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   34
      Left            =   6600
      MouseIcon       =   "frmCrearPersonaje.frx":2E29
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   35
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":2F7B
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   36
      Left            =   6600
      MouseIcon       =   "frmCrearPersonaje.frx":30CD
      MousePointer    =   99  'Custom
      Top             =   6840
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   37
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":321F
      MousePointer    =   99  'Custom
      Top             =   6840
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   38
      Left            =   6600
      MouseIcon       =   "frmCrearPersonaje.frx":3371
      MousePointer    =   99  'Custom
      Top             =   7200
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   39
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":34C3
      MousePointer    =   99  'Custom
      Top             =   7080
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   40
      Left            =   6600
      MouseIcon       =   "frmCrearPersonaje.frx":3615
      MousePointer    =   99  'Custom
      Top             =   7440
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   41
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":3767
      MousePointer    =   99  'Custom
      Top             =   7440
      Width           =   255
   End
   Begin VB.Image boton 
      Height          =   1605
      Index           =   2
      Left            =   1080
      MouseIcon       =   "frmCrearPersonaje.frx":38B9
      MousePointer    =   99  'Custom
      Top             =   4320
      Width           =   1380
   End
   Begin VB.Image boton 
      Height          =   495
      Index           =   1
      Left            =   600
      MouseIcon       =   "frmCrearPersonaje.frx":3A0B
      MousePointer    =   99  'Custom
      Top             =   8400
      Width           =   1965
   End
   Begin VB.Image boton 
      Height          =   570
      Index           =   0
      Left            =   9000
      MouseIcon       =   "frmCrearPersonaje.frx":3B5D
      MousePointer    =   99  'Custom
      Top             =   8400
      Width           =   2400
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   6975
      TabIndex        =   28
      Top             =   7470
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   6975
      TabIndex        =   27
      Top             =   7170
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   6975
      TabIndex        =   26
      Top             =   6825
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   6990
      TabIndex        =   25
      Top             =   6495
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   6990
      TabIndex        =   24
      Top             =   6165
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   6990
      TabIndex        =   23
      Top             =   5850
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   6990
      TabIndex        =   22
      Top             =   5520
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   6990
      TabIndex        =   21
      Top             =   5190
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   6990
      TabIndex        =   20
      Top             =   4875
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   6990
      TabIndex        =   19
      Top             =   4545
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   6990
      TabIndex        =   18
      Top             =   4230
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   6990
      TabIndex        =   17
      Top             =   3885
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   6990
      TabIndex        =   16
      Top             =   3555
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   6990
      TabIndex        =   15
      Top             =   3255
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   6990
      TabIndex        =   14
      Top             =   2925
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   6990
      TabIndex        =   13
      Top             =   2580
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   6990
      TabIndex        =   12
      Top             =   2280
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   6990
      TabIndex        =   11
      Top             =   1950
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   6990
      TabIndex        =   10
      Top             =   1635
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   6990
      TabIndex        =   9
      Top             =   1020
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   6990
      TabIndex        =   8
      Top             =   1305
      Width           =   270
   End
   Begin VB.Image imgHogar 
      Height          =   2850
      Left            =   6120
      Picture         =   "frmCrearPersonaje.frx":3CAF
      Top             =   9000
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.Label lbCarisma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2370
      TabIndex        =   6
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lbSabiduria 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   8640
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2370
      TabIndex        =   4
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2370
      TabIndex        =   3
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2370
      TabIndex        =   2
      Top             =   2070
      Width           =   375
   End
   Begin VB.Label lbFuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2370
      TabIndex        =   1
      Top             =   1650
      Width           =   375
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public SkillPoints As Byte

Function CheckData() As Boolean
If UserRaza = "" Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserSexo = "" Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = "" Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If

If UserHogar = "" Then
    MsgBox "Seleccione el hogar del personaje."
    Exit Function
End If

If SkillPoints > 0 Then
    MsgBox "Asigne los skillpoints del personaje."
    Exit Function
End If

Dim i As Integer
For i = 1 To NUMATRIBUTOS
    If UserAtributos(i) = 0 Then
        MsgBox "Los atributos del personaje son invalidos."
        Exit Function
    End If
Next i

CheckData = True


End Function

Private Sub boton_Click(index As Integer)

Call Audio.PlayWave(SND_CLICK)

Select Case index
    Case 0
        
        Dim i As Integer
        Dim k As Object
        i = 1
        For Each k In Skill
            UserSkills(i) = k.Caption
            i = i + 1
        Next
        
        UserName = txtNombre.Text
        
        If Right$(UserName, 1) = " " Then
                UserName = RTrim$(UserName)
                MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
        End If
        
        UserRaza = lstRaza.List(lstRaza.listIndex)
        UserSexo = lstGenero.List(lstGenero.listIndex)
        UserClase = lstProfesion.List(lstProfesion.listIndex)
        
        UserAtributos(1) = Val(lbFuerza.Caption)
        UserAtributos(2) = Val(lbInteligencia.Caption)
        UserAtributos(3) = Val(lbAgilidad.Caption)
        UserAtributos(4) = Val(lbCarisma.Caption)
        UserAtributos(5) = Val(lbConstitucion.Caption)
        
        UserHogar = lstHogar.List(lstHogar.listIndex)
        
If CheckDatos() Then
#If SeguridadAlkon Then
    UserPassword = MD5.GetMD5String(txtPasswd.Text)
    Call MD5.MD5Reset
#Else
    UserPassword = txtPasswd.Text
#End If
    UserEmail = txtCorreo.Text
    
    If Not CheckMailString(UserEmail) Then
            MsgBox "Direccion de mail invalida."
            Exit Sub
    End If
    If Not Text1.Text = Label1.Caption Then
    MsgBox "Captcha incorrecto, intenta nuevamente", vbCritical, "Captcha"
    Dim Num As Long
    Num = Int(999999 - 100000) * Rnd + 0
    Label1.Caption = Num
    Exit Sub
    End If
    
#If UsarWrench = 1 Then
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
#End If

    'SendNewChar = True
    EstadoLogin = CrearNuevoPj
    
    Me.MousePointer = 11

    EstadoLogin = CrearNuevoPj

#If UsarWrench = 1 Then
    If Not frmMain.Socket1.Connected Then
#Else
    If frmMain.Winsock1.State <> sckConnected Then
#End If
        MsgBox "Error: Se ha perdido la conexion con el server."
        Unload Me
        
    Else
        Call login(RandomCode)
    End If
End If
        
    Case 1
frmMain.Socket1.Disconnect
frmMain.Socket1.Cleanup
frmConnect.MousePointer = 1
Musica = False
Audio.StopMidi
        
        frmConnect.FONDO.Picture = LoadPicture(App.Path & "\Graficos\conectar.jpg")
        Me.Visible = False
        
        
    Case 2
        Call Audio.PlayWave(SND_DICE)
        Call TirarDados
      
End Select


End Sub


Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function
Function intNumeroaleatorio() As Integer
    Dim r As String, s As Integer, T As Integer, seacabo As Boolean
    Dim gletras As String
    Dim gMaxNum As Integer
    seacabo = False
    Do While seacabo = False
        r = CStr(Timer)
        s = Len(r)
        T = mid(r, s, 1)
        intNumeroaleatorio = (T * Int(gletras * Rnd))
        r = CStr(intNumeroaleatorio)
        s = Len(r)
        T = mid(r, s, 1)
        intNumeroaleatorio = T
        If intNumeroaleatorio >= 0 And intNumeroaleatorio < gMaxNum Then
            seacabo = True
        End If
    Loop
End Function


Private Sub TirarDados()
'lbFuerza.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
'lbInteligencia.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
'lbAgilidad.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
'lbCarisma.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
'lbConstitucion.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))

#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
#Else
    If frmMain.Winsock1.State = sckConnected Then
#End If
        Call SendData(Encode64(EncryptStr("TIRDAD", "xaopepe")))
    End If

End Sub

Private Sub Command1_Click(index As Integer)
Call Audio.PlayWave(SND_CLICK)

Dim indice
If index Mod 2 = 0 Then
    If SkillPoints > 0 Then
        indice = index \ 2
        Skill(indice).Caption = Val(Skill(indice).Caption) + 1
        SkillPoints = SkillPoints - 1
    End If
Else
    If SkillPoints < 10 Then
        
        indice = index \ 2
        If Val(Skill(indice).Caption) > 0 Then
            Skill(indice).Caption = Val(Skill(indice).Caption) - 1
            SkillPoints = SkillPoints + 1
        End If
    End If
End If

Puntos.Caption = SkillPoints
End Sub

Private Sub Form_Load()
Dim Num As Long
    Num = Int(999999 - 100000) * Rnd + 0
    Label1.Caption = Num
SkillPoints = 10
Puntos.Caption = SkillPoints
Me.Picture = LoadPicture(App.Path & "\graficos\CP-Interface.jpg")
imgHogar.Picture = LoadPicture(App.Path & "\graficos\CP-Ullathorpe.jpg")

Dim i As Integer
lstProfesion.Clear
For i = LBound(ListaClases) To UBound(ListaClases)
    lstProfesion.AddItem ListaClases(i)
Next i

lstProfesion.listIndex = 1

Call TirarDados
End Sub

Private Sub lstRaza_Click()
    Select Case UCase(lstRaza.List(lstRaza.listIndex))
        Case "HUMANO"
            lblPlusFuerza = "+1"
            lblPlusAgilidad = "+1"
            lblPlusConstitucion = "+2"
            lblPlusInteligencia = "+1"
            lblPlusCarisma = "+0"
        Case "ELFO"
            lblPlusAgilidad = "+4"
            lblPlusInteligencia = "+2"
            lblPlusCarisma = "+2"
            lblPlusConstitucion = "+1"
            lblPlusFuerza = "+1"
        Case "ELFO OSCURO"
            lblPlusFuerza = "+2"
            lblPlusAgilidad = "+2"
            lblPlusInteligencia = "+2"
            lblPlusCarisma = "-3"
            lblPlusConstitucion = "+1"
        Case "ENANO"
            lblPlusFuerza = "+3"
            lblPlusConstitucion = "+3"
            lblPlusInteligencia = "-5"
            lblPlusAgilidad = "+1"
            lblPlusCarisma = "-2"
        Case "GNOMO"
            lblPlusFuerza = "+1"
            lblPlusInteligencia = "+4"
            lblPlusAgilidad = "+3"
            lblPlusCarisma = "+1"
            lblPlusConstitucion = "+0"
        End Select
End Sub

Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_GotFocus()
MsgBox "Sea cuidadoso al seleccionar el nombre de su personaje, Argentum es un juego de rol, un mundo magico y fantastico, si selecciona un nombre obsceno o con connotación politica los administradores borrarán su personaje y no habrá ninguna posibilidad de recuperarlo."
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Function CheckDatos() As Boolean

If txtPasswd.Text <> txtPasswdCheck.Text Then
    MsgBox "Los passwords que tipeo no coinciden, por favor vuelva a ingresarlos."
    Exit Function
End If

If txtCorreo.Text <> txtCorreoCheck.Text Then
    MsgBox "Los Mails que tipeo no coinciden, por favor vuelva a ingresarlos."
    Exit Function
End If

CheckDatos = True

End Function

