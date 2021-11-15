VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{B370EF78-425C-11D1-9A28-004033CA9316}#2.0#0"; "Captura.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   8970
   ClientLeft      =   -1065
   ClientTop       =   -1110
   ClientWidth     =   11970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   6.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00004000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   99  'Custom
   Picture         =   "frmMain.frx":231D7
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   5760
      Top             =   2520
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   999999
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.TextBox Gmtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   195
      TabIndex        =   47
      Top             =   1740
      Width           =   8160
   End
   Begin VB.Timer Ttale 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3000
      Top             =   3240
   End
   Begin VB.Timer Tlemu 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2400
      Top             =   3240
   End
   Begin VB.Timer Tnix 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   3240
   End
   Begin VB.Timer Tulla 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1200
      Top             =   3240
   End
   Begin VB.TextBox SendCMSTXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   195
      TabIndex        =   40
      Top             =   1740
      Width           =   8160
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   195
      TabIndex        =   39
      Top             =   1740
      Visible         =   0   'False
      Width           =   8160
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1200
      Top             =   2520
   End
   Begin VB.Timer tmrAntiSH 
      Interval        =   60000
      Left            =   1800
      Top             =   2520
   End
   Begin MSWinsockLib.Winsock WSAntiSH 
      Left            =   2280
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer EfectosAlpha 
      Interval        =   8
      Left            =   2760
      Top             =   2520
   End
   Begin VB.PictureBox MiniMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   6855
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   98
      TabIndex        =   25
      Top             =   180
      Width           =   1500
   End
   Begin VB.Timer AntiExternos 
      Interval        =   15000
      Left            =   3240
      Top             =   2520
   End
   Begin VB.Timer AntiEngine 
      Interval        =   300
      Left            =   3720
      Top             =   2520
   End
   Begin VB.Timer timerUclick 
      Interval        =   500
      Left            =   4200
      Top             =   2520
   End
   Begin Captura.wndCaptura Foto 
      Left            =   4680
      Top             =   2520
      _ExtentX        =   688
      _ExtentY        =   688
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7680
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   6720
      Top             =   2520
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6240
      Top             =   2520
   End
   Begin VB.Timer FPS 
      Interval        =   1000
      Left            =   7200
      Top             =   2520
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5160
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   30
   End
   Begin VB.PictureBox PanelDer 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8625
      Left            =   8355
      Picture         =   "frmMain.frx":528A7
      ScaleHeight     =   575
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   239
      TabIndex        =   1
      Top             =   195
      Width           =   3585
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   480
         Left            =   510
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   38
         Top             =   7800
         Width           =   480
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   480
         Left            =   2670
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   37
         Top             =   7800
         Width           =   480
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   480
         Left            =   1950
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   36
         Top             =   7800
         Width           =   480
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   480
         Left            =   1230
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   35
         Top             =   7800
         Width           =   480
      End
      Begin VB.PictureBox PicALT 
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   600
         Picture         =   "frmMain.frx":77867
         ScaleHeight     =   1215
         ScaleWidth      =   1695
         TabIndex        =   30
         Top             =   2880
         Visible         =   0   'False
         Width           =   1695
         Begin VB.Label LInfoItem 
            Alignment       =   2  'Center
            BackColor       =   &H00400000&
            BackStyle       =   0  'Transparent
            Caption         =   "Label8"
            ForeColor       =   &H0000C0C0&
            Height          =   375
            Index           =   0
            Left            =   75
            TabIndex        =   34
            Top             =   45
            Width           =   1575
         End
         Begin VB.Label LInfoItem 
            BackColor       =   &H00400000&
            BackStyle       =   0  'Transparent
            Caption         =   "Label8"
            ForeColor       =   &H0000C0C0&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label LInfoItem 
            BackColor       =   &H00400000&
            BackStyle       =   0  'Transparent
            Caption         =   "Label8"
            ForeColor       =   &H0000C0C0&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   32
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label LInfoItem 
            BackColor       =   &H00400000&
            BackStyle       =   0  'Transparent
            Caption         =   "Label8"
            ForeColor       =   &H0000C0C0&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   31
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.CommandButton DespInv 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   3360
         MouseIcon       =   "frmMain.frx":7FFFD
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   4440
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.CommandButton DespInv 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   3360
         MouseIcon       =   "frmMain.frx":8014F
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   4080
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2280
         Left            =   630
         ScaleHeight     =   152
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   160
         TabIndex        =   6
         Top             =   2205
         Width           =   2400
      End
      Begin VB.ListBox hlst 
         BackColor       =   &H00000000&
         ForeColor       =   &H000080FF&
         Height          =   2760
         ItemData        =   "frmMain.frx":802A1
         Left            =   420
         List            =   "frmMain.frx":802A3
         MousePointer    =   99  'Custom
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2040
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Label ManaBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "945/945"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   300
         TabIndex        =   15
         Top             =   6525
         Width           =   1425
      End
      Begin VB.Image Image14 
         Height          =   255
         Left            =   2040
         Top             =   7200
         Width           =   1095
      End
      Begin VB.Label Tale 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1500
         TabIndex        =   46
         Top             =   7080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Lemu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   1125
         TabIndex        =   45
         Top             =   7080
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Nix 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   735
         TabIndex        =   44
         Top             =   7080
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Ulla 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   375
         TabIndex        =   42
         Top             =   7080
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "M"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3150
         TabIndex        =   41
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image13 
         Height          =   255
         Left            =   1920
         MouseIcon       =   "frmMain.frx":802A5
         MousePointer    =   99  'Custom
         Top             =   6240
         Width           =   1335
      End
      Begin VB.Label Labelcasco 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Index           =   3
         Left            =   3480
         TabIndex        =   29
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "55"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   2520
         TabIndex        =   28
         Top             =   165
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label 3"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblFuerza 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "35"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   1320
         TabIndex        =   24
         Top             =   8370
         Width           =   255
      End
      Begin VB.Label lblAgi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "35"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   23
         Top             =   8370
         Width           =   255
      End
      Begin VB.Label lblEscudo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "30/30"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2655
         TabIndex        =   22
         Top             =   7545
         Width           =   525
      End
      Begin VB.Label LblCasc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "30/30"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1935
         TabIndex        =   21
         Top             =   7545
         Width           =   525
      End
      Begin VB.Label lblArma 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "30/30"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1215
         TabIndex        =   20
         Top             =   7545
         Width           =   525
      End
      Begin VB.Label lblArmor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "30/30"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   495
         TabIndex        =   19
         Top             =   7545
         Width           =   525
      End
      Begin VB.Label lblPorcLvl 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "33.33%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   1560
         TabIndex        =   13
         Top             =   1035
         Width           =   660
      End
      Begin VB.Shape ExpShp 
         BackColor       =   &H0000FFFF&
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   225
         Left            =   495
         Top             =   1020
         Width           =   2745
      End
      Begin VB.Label AguBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100/100"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   1680
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label HamBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100/100"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   1560
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label HpBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "396/396"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   300
         TabIndex        =   16
         Top             =   6135
         Width           =   1425
      End
      Begin VB.Label StaBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "715/715"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   300
         TabIndex        =   14
         Top             =   5745
         Width           =   1425
      End
      Begin VB.Image cmdMoverHechi 
         Height          =   375
         Index           =   0
         Left            =   3000
         MouseIcon       =   "frmMain.frx":80F6F
         MousePointer    =   99  'Custom
         Top             =   2100
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Image cmdMoverHechi 
         Height          =   375
         Index           =   1
         Left            =   3000
         MouseIcon       =   "frmMain.frx":81C39
         MousePointer    =   99  'Custom
         Top             =   2520
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Image CmdInfo 
         Height          =   375
         Left            =   1890
         MouseIcon       =   "frmMain.frx":82903
         MousePointer    =   99  'Custom
         Top             =   4920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Image CmdLanzar 
         Height          =   375
         Left            =   360
         MouseIcon       =   "frmMain.frx":835CD
         MousePointer    =   99  'Custom
         Top             =   4920
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "."
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
         Height          =   495
         Left            =   2910
         MouseIcon       =   "frmMain.frx":84297
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label exp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "350/350"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   135
         Left            =   960
         TabIndex        =   10
         Top             =   5520
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Image Image3 
         Height          =   300
         Index           =   0
         Left            =   1800
         MouseIcon       =   "frmMain.frx":84F61
         MousePointer    =   99  'Custom
         Top             =   5520
         Width           =   1605
      End
      Begin VB.Label GldLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   210
         Left            =   2400
         TabIndex        =   9
         Top             =   5610
         Width           =   90
      End
      Begin VB.Image Image1 
         Height          =   300
         Index           =   2
         Left            =   2040
         MouseIcon       =   "frmMain.frx":85C2B
         MousePointer    =   99  'Custom
         Top             =   6840
         Width           =   1170
      End
      Begin VB.Image Image1 
         Height          =   300
         Index           =   1
         Left            =   1920
         MouseIcon       =   "frmMain.frx":868F5
         MousePointer    =   99  'Custom
         Top             =   6480
         Width           =   1410
      End
      Begin VB.Image Image1 
         Height          =   345
         Index           =   0
         Left            =   1920
         MouseIcon       =   "frmMain.frx":875BF
         MousePointer    =   99  'Custom
         Top             =   5880
         Width           =   1410
      End
      Begin VB.Shape AGUAsp 
         BackColor       =   &H00C0C000&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00C0C000&
         Height          =   150
         Left            =   1920
         Top             =   1440
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Shape COMIDAsp 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000C000&
         Height          =   150
         Left            =   600
         Top             =   1440
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Shape MANShp 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00FF0000&
         Height          =   150
         Left            =   300
         Top             =   6555
         Width           =   1425
      End
      Begin VB.Shape STAShp 
         BackColor       =   &H0000C0C0&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000C0C0&
         Height          =   150
         Left            =   300
         Top             =   5775
         Width           =   1425
      End
      Begin VB.Shape Hpshp 
         BorderColor     =   &H80000006&
         FillColor       =   &H000000C0&
         FillStyle       =   0  'Solid
         Height          =   150
         Left            =   300
         Top             =   6165
         Width           =   1425
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1920
         MouseIcon       =   "frmMain.frx":88289
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1440
         Width           =   1485
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   240
         MouseIcon       =   "frmMain.frx":88F53
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Image InvEqu 
         Height          =   4050
         Left            =   210
         Picture         =   "frmMain.frx":89C1D
         Top             =   1380
         Width           =   3240
      End
      Begin VB.Label lbCRIATURA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   5.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   120
         Left            =   555
         TabIndex        =   3
         Top             =   1965
         Width           =   30
      End
      Begin VB.Label LvlLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Height          =   195
         Left            =   3480
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin RichTextLib.RichTextBox RecTxt 
      CausesValidation=   0   'False
      Height          =   1500
      Left            =   195
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   180
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2646
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":95BBA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9000
      TabIndex        =   43
      Top             =   7320
      Width           =   255
   End
   Begin VB.Image Image11 
      Height          =   135
      Left            =   11640
      Top             =   0
      Width           =   135
   End
   Begin VB.Image Image12 
      Height          =   135
      Left            =   11400
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label 3"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8520
      TabIndex        =   27
      Top             =   600
      Width           =   3015
   End
   Begin VB.Image Image9 
      Height          =   255
      Left            =   7920
      MouseIcon       =   "frmMain.frx":95C3F
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   495
   End
   Begin VB.Image Image8 
      Height          =   255
      Left            =   5400
      MouseIcon       =   "frmMain.frx":96909
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Image Image7 
      Height          =   255
      Left            =   3960
      MouseIcon       =   "frmMain.frx":975D3
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   3120
      MouseIcon       =   "frmMain.frx":9829D
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   855
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   2040
      MouseIcon       =   "frmMain.frx":98F67
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   975
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   1320
      MouseIcon       =   "frmMain.frx":99C31
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   240
      MouseIcon       =   "frmMain.frx":9A8FB
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   975
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      Height          =   6225
      Left            =   210
      Top             =   2205
      Width           =   8175
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
' NUNCA OLVIDAR, TAMAÑO DE VISION 545 415
'posicion 14, 147
Public ActualSecond As Long
Public LastSecond As Long
Public tX As Integer
Public tY As Integer
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Public SelM As Integer
Public MapMapa As Integer
Dim gDSB As DirectSoundBuffer
Dim gD As DSBUFFERDESC
Dim gW As WAVEFORMATEX
Dim gFileName As String
Dim dsE As DirectSoundEnum
Dim Pos(0) As DSBPOSITIONNOTIFY
Public IsPlaying As Byte
Dim endEvent As Long
Private TiempoActual As Long
Private Contador As Integer
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim PuedeMacrear As Boolean

'Anti Engine By NicoNZ
Private ElDeAhora As Double
Private Diferencia As Double
Private ElDeAntes As Double
Private Empezo As Boolean
Private Minimo As Double
Private Maximo As Double
Private Cont As Byte
Private EstuboDesbalanceado As Long
Private ContEngine As Byte
'/Anti Engine By NicoNZ

Implements DirectXEvent

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


'Funciones del Api
'-------------------------------------------------------------
Private Declare Function CreateRoundRectRgn Lib "gdi32" ( _
    ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long, _
    ByVal X3 As Long, _
    ByVal Y3 As Long) As Long

Private Declare Function SetWindowRgn Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hRgn As Long, _
    ByVal bRedraw As Boolean) As Long
 'AntiMacros
Dim Macros As AntiMacros

Private Sub Check1_Click()

End Sub

Private Sub CmdLanzar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Macros.ClickRatonDown
End Sub

Private Sub CmdLanzar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'AntiMacros
    'AntiMacros
    If Macros.ClickRatonUP Then
          If hlst.List(hlst.ListIndex) <> "(Vacío)" Then
          If Autoaim = True Then
           Call SendData("PIC" & Magia)
           End If
        Call SendData("VB" & hlst.ListIndex + 1)
        Call SendData("UK" & Magia)
        'frmMain.MousePointer = 2
        'UsaMacro = True
    End If
    Else
       'Call AddtoRichTextBox(frmMain.RecTxt, "Mouse->No se permiten macros externos", 255, 255, 255, False, False, False)
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Macros.ClickKeyDown KeyCode
End Sub
'/AntiMacros


Private Sub Command1_Click()
    Dim ret As Long
    Dim L As Long
    Dim Ancho_form As Long
    Dim Alto_form As Long
    Dim OldScale As Integer
    
    ' guarda el scale del form
    OldScale = ScaleMode
    
    ' cambia la escala ya que el api trabaja con pixeles
    ScaleMode = vbPixels
    
    'Ancho y alto del form en pixeles
    Ancho_form = Me.ScaleWidth
    Alto_form = Me.ScaleHeight
    
    'Crea la región
    ret = CreateRoundRectRgn(10, 35, Ancho_form, Alto_form + 25, 0, 0)
    
    'Aplica la nueva región al formulario
    L = SetWindowRgn(Me.hwnd, ret, True)
    ' reestablece la escala que tenia el formulario
    ScaleMode = OldScale

End Sub



Private Sub AntiEngine_Timer()
'CHOTS | Revisa Intervalos para Engine (Gracias NiconZ)
If Not logged Then Exit Sub

    If GetTickCount - TiempoActual > 350 Or GetTickCount - TiempoActual < 250 Then
        Contador = Contador + 1
    Else
        Contador = 0
    End If
    
    If FramesPerSec < 5 Then
    Contador = Contador + 1
    End If
    
    If Contador > 30 Then
        Call MsgBox("Has Sido Echado por uso de SH", vbCritical, "Chitero")
        Call SendData("BANEAMESpeedHack(Cliente)")
        End
    End If
    
TiempoActual = GetTickCount()
'CHOTS | Revisa Intervalos para Engine (Gracias NiconZ)
End Sub

Private Sub AntiExternos_Timer()
If FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.1.1")) Then
    Call HayExterno("CHEAT ENGINE 5.1.1")
ElseIf FindWindow(vbNullString, UCase$("ART-MONEY")) Then
    Call HayExterno("Art Money")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.0")) Then
    Call HayExterno("CHEAT ENGINE 5.0")
ElseIf FindWindow(vbNullString, UCase$("CROWN MAKRO")) Then
    Call HayExterno("CROWN MAKRO")
ElseIf FindWindow(vbNullString, UCase$("A TRABAJAR...")) Then
    Call HayExterno("A TRABAJAR...")
ElseIf FindWindow(vbNullString, UCase$("ews")) Then
    Call HayExterno("ews")
ElseIf FindWindow(vbNullString, UCase$("Pts")) Then
    Call HayExterno("Auto Pots")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.2")) Then
    Call HayExterno("CHEAT ENGINE 5.2")
ElseIf FindWindow(vbNullString, UCase$("SOLOCOVO?")) Then
    Call HayExterno("SOLOCOVO?")
ElseIf FindWindow(vbNullString, UCase$("-=[ANUBYS RADAR]=-")) Then
    Call HayExterno("-=[ANUBYS RADAR]=-")
ElseIf FindWindow(vbNullString, UCase$("CRAZY SPEEDER 1.05")) Then
    Call HayExterno("CRAZY SPEEDER 1.05")
ElseIf FindWindow(vbNullString, UCase$("SET !XSPEED.NET")) Then
    Call HayExterno("SET !XSPEED.NET")
ElseIf FindWindow(vbNullString, UCase$("SPEEDERXP V1.80 - UNREGISTERED")) Then
    Call HayExterno("SPEEDERXP V1.80 - UNREGISTERED")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.3")) Then
    Call HayExterno("CHEAT ENGINE 5.3")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.4")) Then
    Call HayExterno("CHEAT ENGINE 5.4")
ElseIf FindWindow(vbNullString, UCase$("MACROCRACK <GONZA_VI@HOTMAIL.COM>")) Then
    Call HayExterno("MACRO CRACK")
ElseIf FindWindow(vbNullString, UCase$("MACROCRACK <GONZA_VJ@HOTMAIL.COM>")) Then
    Call HayExterno("MACRO CRACK")
ElseIf FindWindow(vbNullString, UCase$("MACRO CRACK <GONZA_VI@HOTMAIL.COM>")) Then
    Call HayExterno("MACRO CRACK")
ElseIf FindWindow(vbNullString, UCase$("MACRO CRACK <GONZA_VJ@HOTMAIL.COM>")) Then
    Call HayExterno("MACRO CRACK")
ElseIf FindWindow(vbNullString, UCase$("CHITS")) Then
    Call HayExterno("EL CHEAT DE GERI")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.1")) Then
    Call HayExterno("CHEAT ENGINE 5.1")
ElseIf FindWindow(vbNullString, UCase$("A SPEEDER")) Then
    Call HayExterno("A SPEEDER")
ElseIf FindWindow(vbNullString, UCase$("MEMO :P")) Then
    Call HayExterno("MEMO :P")
ElseIf FindWindow(vbNullString, UCase$("ORK4M VERSION 1.5")) Then
    Call HayExterno("ORK4M VERSION 1.5")
ElseIf FindWindow(vbNullString, UCase$("ORKAM")) Then
    Call HayExterno("ORK4M VERSION 1.5")
ElseIf FindWindow(vbNullString, UCase$("MACRO")) Then
    Call HayExterno("Macro")
ElseIf FindWindow(vbNullString, UCase$("BY FEDEX")) Then
    Call HayExterno("By Fedex")
ElseIf FindWindow(vbNullString, UCase$("!XSPEED.NET +4.59")) Then
    Call HayExterno("!Xspeeder")
ElseIf FindWindow(vbNullString, UCase$("CAMBIA TITULOS DE CHEATS BY FEDEX")) Then
    Call HayExterno("Cambia titulos")
ElseIf FindWindow(vbNullString, UCase$("NEWENG OCULTO")) Then
    Call HayExterno("Cambia titulos")
ElseIf FindWindow(vbNullString, UCase$("SERBIO ENGINE")) Then
    Call HayExterno("Serbio Engine")
ElseIf FindWindow(vbNullString, UCase$("REYMIX ENGINE 5.3 PUBLIC")) Then
    Call HayExterno("ReyMix Engine")
ElseIf FindWindow(vbNullString, UCase$("REY ENGINE 5.2")) Then
    Call HayExterno("ReyMix Engine")
ElseIf FindWindow(vbNullString, UCase$("AUTOCLICK - BY NIO_SHOOTER")) Then
    Call HayExterno("AutoClick")
ElseIf FindWindow(vbNullString, UCase$("TONNER MINER! :D [REG][SKLOV] 2.0")) Then
    Call HayExterno("Tonner")
ElseIf FindWindow(vbNullString, UCase$("Buffy The vamp Slayer")) Then
    Call HayExterno("Buffy The vamp Slayer")
ElseIf FindWindow(vbNullString, UCase$("Blorb Slayer 1.12.552 (BETA)")) Then
    Call HayExterno("Blorb Slayer 1.12.552 (BETA)")
ElseIf FindWindow(vbNullString, UCase$("PumaEngine3.0")) Then
    Call HayExterno("PumaEngine3.0")
ElseIf FindWindow(vbNullString, UCase$("Vicious Engine 5.0")) Then
    Call HayExterno("Vicious Engine 5.0")
ElseIf FindWindow(vbNullString, UCase$("AkumaEngine33")) Then
    Call HayExterno("AkumaEngine33")
ElseIf FindWindow(vbNullString, UCase$("Spuc3ngine")) Then
    Call HayExterno("Spuc3ngine")
ElseIf FindWindow(vbNullString, UCase$("Ultra Engine")) Then
    Call HayExterno("Ultra Engine")
ElseIf FindWindow(vbNullString, UCase$("Engine")) Then
    Call HayExterno("Engine")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V5.4")) Then
    Call HayExterno("Cheat Engine V5.4")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.4")) Then
    Call HayExterno("Cheat Engine V4.4")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.4 German Add-On")) Then
    Call HayExterno("Cheat Engine V4.4 German Add-On")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.3")) Then
    Call HayExterno("Cheat Engine V4.3")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.2")) Then
    Call HayExterno("Cheat Engine V4.2")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.1.1")) Then
    Call HayExterno("Cheat Engine V4.1.1")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.3")) Then
    Call HayExterno("Cheat Engine V3.3")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.2")) Then
    Call HayExterno("Cheat Engine V3.2")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.1")) Then
    Call HayExterno("Cheat Engine V3.1")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine")) Then
    Call HayExterno("Cheat Engine")
ElseIf FindWindow(vbNullString, UCase$("danza engine 5.2.150")) Then
    Call HayExterno("danza engine 5.2.150")
ElseIf FindWindow(vbNullString, UCase$("zenx engine")) Then
    Call HayExterno("zenx engine")
ElseIf FindWindow(vbNullString, UCase$("MACROMAKER")) Then
    Call HayExterno("Macro Maker")
ElseIf FindWindow(vbNullString, UCase$("MACREOMAKER - EDIT MACRO")) Then
    Call HayExterno("Macro Maker")
ElseIf FindWindow(vbNullString, UCase$("By Fedex")) Then
    Call HayExterno("Macro Fedex")
ElseIf FindWindow(vbNullString, UCase$("Macro Mage 1.0")) Then
    Call HayExterno("Macro Mage")
ElseIf FindWindow(vbNullString, UCase$("Auto* v0.4 (c) 2001 Pete Powa")) Then
    Call HayExterno("Macro Fisher")
ElseIf FindWindow(vbNullString, UCase$("Kizsada")) Then
    Call HayExterno("Macro K33")
ElseIf FindWindow(vbNullString, UCase$("Makro K33")) Then
    Call HayExterno("Macro K33")
ElseIf FindWindow(vbNullString, UCase$("Super Saiyan")) Then
    Call HayExterno("El Chit del Geri")
ElseIf FindWindow(vbNullString, UCase$("Makro-Piringulete")) Then
    Call HayExterno("Piringulete")
ElseIf FindWindow(vbNullString, UCase$("Makro-Piringulete 2003")) Then
    Call HayExterno("Piringulete 2003")
ElseIf FindWindow(vbNullString, UCase$("TUKY2005")) Then
    Call HayExterno("Makro Tuky")
ElseIf FindWindow(vbNullString, UCase$("Countach")) Then
    Call HayExterno("Countach")
    ElseIf FindWindow(vbNullString, UCase$("MacroRecorder")) Then
    Call HayExterno("MacroRecorder")
ElseIf FindWindow(vbNullString, UCase$("Ultimatemacros")) Then
    Call HayExterno("Ultimatemacros")
ElseIf FindWindow(vbNullString, UCase$("MacroLauncher")) Then
    Call HayExterno("MacroLauncher")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine 5.5")) Then
    Call HayExterno("Cheat Engine 5.5")
ElseIf FindWindow(vbNullString, UCase$("Auto Remo- TheFrank^")) Then
 Call HayExterno("Auto Remo- TheFrank^")
 ElseIf FindWindow(vbNullString, UCase$("WPE PRO")) Then
 Call HayExterno("WPE PRO")
 ElseIf FindWindow(vbNullString, UCase$("WPE PRO - revivalao.exe")) Then
 Call HayExterno("WPE PRO - revivalao.exe")
  ElseIf FindWindow(vbNullString, UCase$("WPE PRO - [WPEPRO2]")) Then
 Call HayExterno("WPE PRO - [WPEPRO2]")
  ElseIf FindWindow(vbNullString, UCase$("WPE PRO [WPEPRO2]")) Then
 Call HayExterno("WPE PRO [WPEPRO2]]")
 ElseIf FindWindow(vbNullString, UCase$("WPE PRO - revivalao.exe - [WPEPRO2]")) Then
 Call HayExterno("WPE PRO - revivalao.exe - [WPEPRO2]")
End If


End Sub

Private Sub cmdMoverHechi_Click(index As Integer)
If hlst.ListIndex = -1 Then Exit Sub

Select Case index
Case 0 'subir
    If hlst.ListIndex = 0 Then Exit Sub
Case 1 'bajar
    If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
End Select

Call SendData("DESPHE" & index + 1 & "," & hlst.ListIndex + 1)

Select Case index
Case 0 'subir
    hlst.ListIndex = hlst.ListIndex - 1
Case 1 'bajar
    hlst.ListIndex = hlst.ListIndex + 1
End Select

End Sub

Private Sub Command12_Click()

End Sub

Private Sub DirectXEvent_DXCallback(ByVal eventid As Long)

End Sub

Private Sub CreateEvent()
     endEvent = DirectX.CreateEvent(Me)
End Sub



Private Sub EfectosAlpha_Timer()
If Desbanecimiento1 = True Then
    If Val(AlphaX) = 20 Then
        Desbanecimiento1 = False
        Desbanecimiento2 = True
    Else
        AlphaX = Val(AlphaX) - 5
    End If
End If
If Desbanecimiento2 = True Then
    If Val(AlphaX) = 250 Then
        Desbanecimiento1 = True
        Desbanecimiento2 = False
    Else
        AlphaX = Val(AlphaX) + 5
    End If
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Macros.ClickRatonDown
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'AntiMacros
    'AntiMacros
    If Macros.ClickRatonUP Then
        'Text1.Text = "Mouse Sin Macro" & vbCrLf & Text1.Text
            If Cartel Then Cartel = False

#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(MouseBoton, True)
#End If

    If Not Comerciando Then
        Call ConvertCPtoTP(MainViewShp.Left, MainViewShp.Top, MouseX, MouseY, tX, tY)

        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                If UsingSkill = 0 Then
                    SendData "LC" & tX & "," & tY
                Else
                    frmMain.MousePointer = vbCustom
                    If (UsingSkill = Magia Or UsingSkill = Proyectiles) And UserCanAttack = 0 Then Exit Sub
                    SendData "WLC" & tX & "," & tY & "," & UsingSkill
                    If UsingSkill = Magia Or UsingSkill = Proyectiles Then UserCanAttack = 0
                    UsingSkill = 0
                End If
            Else
                Call AbrirMenuViewPort
            End If
        ElseIf (MouseShift And 1) = 1 Then
            If MouseShift = vbLeftButton Then
                Call SendData("/TELEP YO " & UserMap & " " & tX & " " & tY)
                puedehablar = True
        End If
        End If
    End If
    
    Else
       'Call AddtoRichTextBox(frmMain.RecTxt, "Mouse->No se permiten macros externos", 255, 255, 255, False, False, False)
        Exit Sub
    End If
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If endEvent Then
        DirectX.DestroyEvent endEvent
    End If
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set Macros = Nothing
End Sub
Private Sub FPS_Timer()

If logged And Not frmMain.Visible Then
    Unload frmConnect
    frmMain.Show
End If
    
End Sub

Private Sub lblBlues_Click()
End Sub


Private Sub LblCasco_Click()

End Sub

Private Sub hlst_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Macros.ClickRatonDown
End Sub

Private Sub hlst_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'AntiMacros
    'AntiMacros
    If Macros.ClickRatonUP Then
        'Text1.Text = "Mouse Sin Macro" & vbCrLf & Text1.Text
    Else
       'Call AddtoRichTextBox(frmMain.RecTxt, "Mouse->No se permiten macros externos", 255, 255, 255, False, False, False)
        Exit Sub
    End If
End Sub

Private Sub Image10_Click()
Dim web As Long
web = ShellExecute(Me.hwnd, "open", "http://www.revivalao.com.ar/", "", "", 1)

End Sub

Private Sub Image11_Click()
Call SendData("/SALIR")
End Sub

Private Sub Image12_Click()
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False
End Sub

Private Sub Image13_Click()
Call SendData("/CCCC")
End Sub

Private Sub Image14_Click()
Call SendData("/TIEMPOS")
End Sub

Private Sub Image2_Click()
Call SendData("/ISLA")
End Sub

Private Sub Image4_Click()
Call SendData("/SHOW SOS")
End Sub

Private Sub Image5_Click()
 Call SendData("/ONLINEMAP")
End Sub

Private Sub Image6_Click()
Call SendData("/PANELDEGMS")
End Sub

Private Sub Image7_Click()
Call SendData("/INVISIBLE")
End Sub

Private Sub Image8_Click()
Call SendData("/ATORNEO")
End Sub

Private Sub Image9_Click()
Call SendData("/SALIR")
End Sub

Private Sub InvEqu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)



PicALT.Visible = False
End Sub





Private Sub Label10_Click()
SendData "/VERS"
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Macros.ClickRatonDown
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'AntiMacros
    'AntiMacros
    If Macros.ClickRatonUP Then
        PicALT.Visible = False
    Call Audio.PlayWave(SND_CLICK)
Call cargarImagenRes(frmMain.InvEqu, 137)
    'InvEqu.Picture = LoadPicture(App.Path & "\Graficos\Centronuevohechizos.jpg")
    '%%%%%%OCULTAMOS EL INV&&&&&&&&&&&&
    'DespInv(0).Visible = False
    'DespInv(1).Visible = False
    picInv.Visible = False
    hlst.Visible = True
    CmdInfo.Visible = True
    CmdLanzar.Visible = True

    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    Else
       'Call AddtoRichTextBox(frmMain.RecTxt, "Mouse->No se permiten macros externos", 255, 255, 255, False, False, False)
        Exit Sub
    End If
End Sub

Private Sub LInfoItem_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
PicALT.Visible = False
End Sub

Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    SendData "LC" & tX & "," & tY
    SendData "/COMERCIAR"
End Sub

Private Sub mnuNpcDesc_Click()
    SendData "LC" & tX & "," & tY
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub
Private Sub picALT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicALT.Visible = False
End Sub

Private Sub picInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Macros.ClickRatonDown
End Sub

Private Sub picInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo tronco
If inventario.ItemName(tempitem) = "(Vacío)" Or inventario.ItemName(tempitem) = vbNullString Then
PicALT.Visible = False
Exit Sub
Else
PicALT.Visible = True
PicALT.Move (picInv.Left + X - 50), (picInv.Top + Y + 30)
    LInfoItem(0).Caption = inventario.ItemName(tempitem)
    LInfoItem(1).Caption = "Golpe Mínimo: " & inventario.MinHit(tempitem)
    LInfoItem(2).Caption = "Golpe Máximo: " & inventario.MaxHit(tempitem)
    LInfoItem(3).Caption = "Defensa: " & inventario.Def(tempitem)

    End If
tronco:
End Sub

Private Sub SpoofCheck_Timer()

Dim IPMMSB As Byte
Dim IPMSB As Byte
Dim IPLSB As Byte
Dim IPLLSB As Byte

IPLSB = 3 + 15
IPMSB = 32 + 15
IPMMSB = 200 + 15
IPLLSB = 74 + 15

If IPdelServidor <> ((IPMMSB - 15) & "." & (IPMSB - 15) & "." & (IPLSB - 15) _
& "." & (IPLLSB - 15)) Then End

End Sub

Private Sub Second_Timer()
    ActualSecond = Mid(Time, 7, 2)
    ActualSecond = ActualSecond + 1
    If ActualSecond = LastSecond Then End
    LastSecond = ActualSecond
    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
End Sub

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
    If (inventario.SelectedItem > 0 And inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (inventario.SelectedItem = FLAGORO) Then
        If inventario.Amount(inventario.SelectedItem) = 1 Then
            SendData "OH" & inventario.SelectedItem & "," & 1
        Else
           If inventario.Amount(inventario.SelectedItem) > 1 Then
            frmCantidad.Show , frmMain
           End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    SendData "AG"
End Sub

Private Sub UsarItem()
SendData "HDP"
    If (inventario.SelectedItem > 0) And (inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then SendData "USA" & inventario.SelectedItem
End Sub

Private Sub EquiparItem()
    If (inventario.SelectedItem > 0) And (inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        SendData "EQUI" & inventario.SelectedItem
        
        
End Sub
Private Sub CmdInfo_Click()
    Call SendData("INFS" & hlst.ListIndex + 1)
End Sub

''''''''''''''''''''''''''''''''''''''
'     OTROS                          '
''''''''''''''''''''''''''''''''''''''

Private Sub DespInv_Click(index As Integer)
    inventario.ScrollInventory (index = 0)
End Sub
Private Sub panelder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)



PicALT.Visible = False
End Sub
Private Sub Form_DblClick()
    If Not frmForo.Visible Then
        SendData "RC" & tX & "," & tY
        Call SendData("/MOV")
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(KeyCode, False)
#End If
        
    If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) And (Not Gmtxt.Visible) And _
       ((KeyCode >= 65 And KeyCode <= 90) Or _
       (KeyCode >= 48 And KeyCode <= 57)) Then
        
        If Macros.ClickKeyUP(KeyCode) Then
    'Text1.Text = "Tecla Sin Macro: " & Chr(KeyCode) & "(" & KeyCode & ")" & vbCrLf & Text1.Text
Else
    
Exit Sub
End If
            Select Case KeyCode
                Case vbKeyM:
                    If Not Audio.PlayingMusic Then
                        Musica = True
                        Audio.PlayMIDI CStr(currentMidi) & ".mid"
                    Else
                        Musica = False
                        Audio.StopMidi
                    End If
                    
                Case vbKeyA:
                    Call AgarrarItem
                Case vbKeyE:
                    Call EquiparItem
                Case vbKeyN:
                    If Nombres = True Then
                    Nombres = False
                    Else
                    Nombres = True
                    End If
                Case vbKeyD
                    Call SendData("UK" & Domar)
                Case vbKeyR:
                    Call SendData("UK" & Robar)
                Case vbKeyS:
                    Call SendData("/SEG")
                Case vbKeyW:
                    Call SendData("/SEGCLAN")
                    Case vbKeyZ:
                     frmMain.RecTxt.Text = ""
                Case vbKeyO:
                    Call SendData("UK" & Ocultarse)
                Case vbKeyT:
                    Call TirarItem
        Case vbKeyU:
                    If Not NoPuedeUsar Then
                        NoPuedeUsar = True
                        Call UsarItem
                    End If
        
        Case vbKeyP:
                    If Not NoPuedeUsar Then
                        NoPuedeUsar = True
                        Call UsarItem
                    End If
                Case vbKeyL:
                    If UserPuedeRefrescar Then
                        Call SendData("RPU")
                        UserPuedeRefrescar = False
                        Beep
                    End If
            End Select
        End If
        
        Select Case KeyCode
            Case vbKeyReturn:
                If SendCMSTXT.Visible Or Gmtxt.Visible = True Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendTxt.Visible = True
                    SendTxt.SetFocus
                End If
            Case vbKeyDelete:
                If SendTxt.Visible Or Gmtxt.Visible = True Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendCMSTXT.Visible = True
                    SendCMSTXT.SetFocus
                End If
                Case vbKeyShift:
                If SendTxt.Visible Or SendCMSTXT.Visible = True Then Exit Sub
                If puedehablar = True Then
                puedehablar = False
                Exit Sub
                End If
                If Not frmCantidad.Visible Then
                    Gmtxt.Visible = True
                    Gmtxt.SetFocus
                End If
            Case vbKeyF2:
                Call SendData("/COLAPAJA23")
                Call SendData("/COLAPINCHADA32")
                Call SendData("/ONLINEGM")
            Case vbKeyF3:
                Call SendData("/COLADESHURA11")
                Call SendData("/SEMANTICOZ23")
            Case vbKeyF4:
               Call SendData("/SALIR")
            Case vbKeyControl:
                If (Not UserDescansar) And _
                   (Not UserMeditar) Then
                        SendData "KC"
                End If
                Case vbKeyF12:
                If Timer1.Enabled = True Then
                Call AddtoRichTextBox(frmMain.RecTxt, "Macro herramientas desactivado!", 255, 255, 255, False, False, False)
                Timer1.Enabled = False
                Exit Sub
                End If
      
        If Not inventario.OBJType(inventario.SelectedItem) = 18 Then
         Call AddtoRichTextBox(frmMain.RecTxt, "Debes tener seleccionada la herramienta!", 255, 255, 255, False, False, False)
         Exit Sub
        End If
         If inventario.Equipped(inventario.SelectedItem) = False Then
       Call AddtoRichTextBox(frmMain.RecTxt, "Debes equiparte la herramienta!", 255, 255, 255, False, False, False)
       Exit Sub
       End If
        If Timer1.Enabled = False Then
        Timer1.Enabled = True
        Call AddtoRichTextBox(frmMain.RecTxt, "Macro herramientas activado!", 255, 255, 255, False, False, False)
        Else
        Timer1.Enabled = False
        Call AddtoRichTextBox(frmMain.RecTxt, "Macro herramientas desactivado!", 255, 255, 255, False, False, False)
        End If
            Case vbKeyF5:
                Call frmOpciones.Show(vbModeless, frmMain)
            Case vbKeyF1:
        Call frmMapa.Show(vbModeless, frmMain)
            Case vbKeyF6:
                Call SendData("/SOBAMELA441")
            Case vbKeyF7:
                Call SendData("/HACEME1PT3")
            Case vbKeyF8:
                Call SendData("/COMERCIAR")
            Case vbKeyF9:
             If Autoaim = False Then
                 Call AddtoRichTextBox(frmMain.RecTxt, "AutoAim Activado!", 255, 255, 255, False, False, False)
             Autoaim = True
             Else
                 Call AddtoRichTextBox(frmMain.RecTxt, "AutoAim Desactivado!", 255, 255, 255, False, False, False)
             Autoaim = False
             End If
                        
        End Select
 
End Sub

Private Sub Form_Load()
Detectar RecTxt.hwnd, Me.hwnd
'Antish.GetNistTime
Set Macros = New AntiMacros
SendTxt.Visible = False
SendCMSTXT.Visible = False
Gmtxt.Visible = False
TiempoActual = GetTickCount()

    Label3.Caption = UserName
    Label6.Caption = UserLvl
    

   ' PanelDer.Picture = LoadPicture(App.Path & _
   ' "\Graficos\Principalnuevo_sin_energia.jpg")
    
    'InvEqu.Picture = LoadPicture(App.Path & _
   ' "\Graficos\Centronuevoinventario.jpg")
   Me.Left = 0
   Me.Top = 0
   
    If AntiEngine.Interval <> 300 Or AntiEngine.Enabled = False Then
        Call CliEditado
    ElseIf AntiExternos.Interval <> 15000 Or AntiExternos.Enabled = False Then
        Call CliEditado
    End If


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
    PicALT.Visible = False

End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub
Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub
Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub Image1_Click(index As Integer)
    Call Audio.PlayWave(SND_CLICK)
PicALT.Visible = False
    Select Case index
        Case 0
            '[MatuX] : 01 de Abril del 2002
                Call frmOpciones.Show(vbModeless, frmMain)
             
            '[END]
        Case 1
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            SendData "ATRI"
            SendData "ESKI"
            SendData "FEST"
            SendData "FAMA"
            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
        Case 2
            If Not frmGuildLeader.Visible Then _
                Call SendData("GLINFO")
    End Select
End Sub

Private Sub Image3_Click(index As Integer)
    Select Case index
        Case 0
            inventario.SelectGold
            If UserGLD > 0 Then
             Call FrmTransferir.Show(vbModeless, frmMain)
            End If
    End Select
End Sub

Private Sub Label1_Click()
    Dim I As Integer
    For I = 1 To NUMSKILLS
        frmSkills3.Text1(I).Caption = UserSkills(I)
    Next I
    Alocados = SkillPoints
    frmSkills3.Puntos.Caption = SkillPoints
    frmSkills3.Show , frmMain
End Sub

Private Sub Label4_Click()
    Call Audio.PlayWave(SND_CLICK)
Call cargarImagenRes(frmMain.InvEqu, 136)
    'InvEqu.Picture = LoadPicture(App.Path & "\Graficos\Centronuevoinventario.jpg")

    'DespInv(0).Visible = True
    'DespInv(1).Visible = True
    picInv.Visible = True

    hlst.Visible = False
    CmdInfo.Visible = False
    CmdLanzar.Visible = False
   
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
End Sub

Private Sub picInv_DblClick()
    'AntiMacros

  If ALaMierda = True Then
  Call UsarItem
  ALaMierda = False
  End If
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'AntiMacros
    'AntiMacros
    If Macros.ClickRatonUP Then
     ALaMierda = True
      Call Audio.PlayWave(SND_CLICK)
        'Text1.Text = "Mouse Sin Macro" & vbCrLf & Text1.Text
    Else
    ALaMierda = False
       'Call AddtoRichTextBox(frmMain.RecTxt, "Mouse->No se permiten macros externos", 255, 255, 255, False, False, False)
        Exit Sub
    End If
   
End Sub

Private Sub RecTxt_Change()
    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf Me.SendCMSTXT.Visible Then
        SendCMSTXT.SetFocus
        ElseIf Gmtxt.Visible Then
        Gmtxt.SetFocus
    Else
      If (Not frmComerciar.Visible) And _
         (Not frmSkills3.Visible) And _
         (Not frmMSG.Visible) And _
         (Not frmForo.Visible) And _
         (Not frmEstadisticas.Visible) And _
         (Not frmCantidad.Visible) And _
         (picInv.Visible) Then
            picInv.SetFocus
      End If
    End If
    On Error GoTo 0
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
    'If Len(SendTxt.Text) > 99999 Then
     '   stxtbuffer = "Soy un cheater, avisenle a un gm"
    'Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim I As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For I = 1 To Len(SendTxt.Text)
            CharAscii = Asc(Mid$(SendTxt.Text, I, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next I
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
   ' End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If Left$(stxtbuffer, 1) = "/" Then
            If UCase(Left$(stxtbuffer, 8)) = "/PASSWD " Then
                    Dim j As String
#If SeguridadAlkon Then
                    j = md5.GetMD5String(Right$(stxtbuffer, Len(stxtbuffer) - 8))
                    Call md5.MD5Reset
#Else
                    j = Right$(stxtbuffer, Len(stxtbuffer) - 8)
#End If
                    stxtbuffer = "/PASSWD " & j
            ElseIf UCase$(stxtbuffer) = "/FUNDARCLAN" Then
                frmEligeAlineacion.Show vbModeless, Me
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
                
                Exit Sub
            End If
            Call SendData(stxtbuffer)
    
       'Shout
        ElseIf Left$(stxtbuffer, 1) = "-" Then
            Call SendData("-" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

        'Whisper
        ElseIf Left$(stxtbuffer, 1) = "\" Then
            Call SendData("\" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

        'Say
        ElseIf stxtbuffer <> "" Then
            Call SendData(";" & stxtbuffer)
If Not EsMalapalabra(stxtbuffer) Then
'[Gabriel Mellace]
Call SendData(";" & stxtbuffer)
frmMain.SendTxt.Text = ""
stxtbuffer = ""
KeyCode = 0
frmMain.SendTxt.Visible = False
End If
        End If
        

        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
    End If
End Sub


Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbuffercmsg <> "" Then
            Call SendData("/CMSG " & stxtbuffercmsg)
        End If

        stxtbuffercmsg = ""
        SendCMSTXT.Text = ""
        KeyCode = 0
        Me.SendCMSTXT.Visible = False
    End If
End Sub
Private Sub Gmtxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If textogm <> "" Then
            Call SendData("/r " & textogm)
        End If

        textogm = ""
        Gmtxt.Text = ""
        KeyCode = 0
        Me.Gmtxt.Visible = False
    End If
End Sub


Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub
Private Sub Gmtxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub


Private Sub SendCMSTXT_Change()
  '  If Len(SendCMSTXT.Text) > 999 Then
   '     stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
    'Else
        stxtbuffercmsg = SendCMSTXT.Text
    'End If
End Sub

Private Sub Gmtxt_Change()
  '  If Len(SendCMSTXT.Text) > 999 Then
   '     stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
    'Else
       textogm = Gmtxt.Text
    'End If
End Sub



''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''
#If UsarWrench = 1 Then

Private Sub Socket1_Connect()
    Dim ServerIp As String
    Dim Temporal1 As Long
    Dim Temporal As Long
    
    
    ServerIp = Socket1.PeerAddress
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = ((Mid$(ServerIp, 1, Temporal - 1) Xor &H65) And &H7F) * 16777216
    ServerIp = Mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (Mid$(ServerIp, 1, Temporal - 1) Xor &HF6) * 65536
    ServerIp = Mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (Mid$(ServerIp, 1, Temporal - 1) Xor &H4B) * 256
    ServerIp = Mid$(ServerIp, Temporal + 1, Len(ServerIp)) Xor &H42
    MixedKey = (Temporal1 + ServerIp)
    
    Second.Enabled = True
Clavenueva = Encode64(RandomNumber(1, 2123424))
    Clavefija = "xaopepe"
    Call SendData("CLAVE" & Clavenueva)
    Clavefija = Clavenueva

    'If frmCrearPersonaje.Visible Then
    If EstadoLogin = E_MODO.CrearNuevoPj Then
        Call SendData("gIvEmEvAlcOde")
#If SegudidadAlkon Then
        Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
    'ElseIf Not frmRecuperar.Visible Then
    ElseIf EstadoLogin = E_MODO.Normal Then
    
        Call SendData("gIvEmEvAlcOde")
        
#If SegudidadAlkon Then
        Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
    ElseIf EstadoLogin = E_MODO.Dados Then
        Call SendData("gIvEmEvAlcOde")
#If SegudidadAlkon Then
        Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
    'Else
    ElseIf EstadoLogin = E_MODO.RecuperarPass Then
        Dim cmd As String
        cmd = "PASSRECO" & frmRecuperar.txtNombre.Text & "~" & frmRecuperar.txtCorreo
        frmMain.Socket1.Write cmd, Len(cmd)
    End If
End Sub

Private Sub Socket1_Disconnect()
    Dim I As Long
    
    LastSecond = 0
    Second.Enabled = False
    logged = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConnect.MousePointer = vbNormal
    
    If frmPasswdSinPadrinos.Visible = True Then frmPasswdSinPadrinos.Visible = False
    frmCrearPersonaje.Visible = False
    frmConnect.Visible = True
    
    On Local Error Resume Next
    For I = 0 To Forms.count - 1
        If Forms(I).Name <> Me.Name And Forms(I).Name <> frmConnect.Name Then
            Unload Forms(I)
        End If
    Next I
    On Local Error GoTo 0
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False
    
#If SegudidadAlkon Then
    LOGGING = False
    LOGSTRING = False
    LastPressed = 0
    LastMouse = False
    LastAmount = 0
#End If

    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    
    For I = 1 To NUMSKILLS
        UserSkills(I) = 0
    Next I

    For I = 1 To NUMATRIBUTOS
        UserAtributos(I) = 0
    Next I

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    End If
    
    Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Response = 0
    LastSecond = 0
    Second.Enabled = False

    frmMain.Socket1.Disconnect
    
    If frmOldPersonaje.Visible Then
        frmOldPersonaje.Visible = False
    End If

    If Not frmCrearPersonaje.Visible Then
        If Not frmBorrar.Visible And Not frmRecuperar.Visible Then
            frmConnect.Show
        End If
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
    Dim loopc As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim echar As Integer
    Dim aux$
    Dim nfile As Integer
    
    Socket1.Read RD, DataLength
    
    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    'Check for more than one line
    sChar = 1
    For loopc = 1 To Len(RD)

        tChar = Mid$(RD, loopc, 1)

        If tChar = ENDC Then
            CR = CR + 1
            echar = loopc - sChar
            rBuffer(CR) = Mid$(RD, sChar, echar)
            sChar = loopc + 1
        End If

    Next loopc

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = Mid$(RD, sChar, Len(RD))
    End If

    'Send buffer to Handle data
    For loopc = 1 To CR
        'Call LogCustom("HandleData: " & rBuffer(loopc))
        Call HandleData(rBuffer(loopc))
    Next loopc
End Sub


#End If

Private Sub AbrirMenuViewPort()
#If (ConMenuseConextuales = 1) Then

If tX >= MinXBorder And tY >= MinYBorder And _
    tY <= MaxYBorder And tX <= MaxXBorder Then
    If MapData(tX, tY).charindex > 0 Then
        If charlist(MapData(tX, tY).charindex).invisible = False Then
        
            Dim I As Long
            Dim m As New frmMenuseFashion
            
            Load m
            m.SetCallback Me
            m.SetMenuId 1
            m.ListaInit 2, False
            
            If charlist(MapData(tX, tY).charindex).Nombre <> "" Then
                m.ListaSetItem 0, charlist(MapData(tX, tY).charindex).Nombre, True
            Else
                m.ListaSetItem 0, "<NPC>", True
            End If
            m.ListaSetItem 1, "Comerciar"
            
            m.ListaFin
            m.Show , Me

        End If
    End If
End If

#End If
End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)
Select Case MenuId

Case 0 'Inventario
    Select Case Sel
    Case 0
    Case 1
    Case 2 'Tirar
        Call TirarItem
     Case 3 'Usar
        If Not NoPuedeUsar Then
            NoPuedeUsar = True
            Call UsarItem
        End If
    Case 3 'equipar
        Call EquiparItem
    End Select
    
Case 1 'Menu del ViewPort del engine
    Select Case Sel
    Case 0 'Nombre
        SendData "LC" & tX & "," & tY
    Case 1 'Comerciar
        Call SendData("LC" & tX & "," & tY)
        Call SendData("/COMERCIAR")
    End Select
End Select
End Sub


Private Sub textgm_Change()

End Sub

Private Sub Timer1_Timer()
If inventario.OBJType(inventario.SelectedItem) = 18 Then
Call UsarItem
'Form_click
If Cartel Then Cartel = False

#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(MouseBoton, True)
#End If

    If Not Comerciando Then
        Call ConvertCPtoTP(MainViewShp.Left, MainViewShp.Top, MouseX, MouseY, tX, tY)

        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                If UsingSkill = 0 Then
                    SendData "LC" & tX & "," & tY
                Else
                    frmMain.MousePointer = vbCustom
                    If (UsingSkill = Magia Or UsingSkill = Proyectiles) And UserCanAttack = 0 Then Exit Sub
                    SendData "WLC" & tX & "," & tY & "," & UsingSkill
                    If UsingSkill = Magia Or UsingSkill = Proyectiles Then UserCanAttack = 0
                    UsingSkill = 0
                End If
            Else
                Call AbrirMenuViewPort
            End If
        ElseIf (MouseShift And 1) = 1 Then
            If MouseShift = vbLeftButton Then
                Call SendData("/TELEP YO " & UserMap & " " & tX & " " & tY)
        End If
        End If
    End If
Else
Call AddtoRichTextBox(frmMain.RecTxt, "Debes equiparte y seleccionar la herramienta!", 255, 255, 255, False, False, False)
End If
End Sub

Private Sub timerUclick_Timer()
PuedeUclickear = True
frmMain.timerUclick.Enabled = False
End Sub

Private Sub Tlemu_Timer()
frmMain.Lemu.Visible = False
Tlemu.Enabled = False
End Sub

Private Sub tmrAntiSH_Timer()
Dim minuto As Integer
minuto = minuto + 1
If minuto = 5 Then
'Antish.GetNistTime
minuto = 0
End If
End Sub

Private Sub Tnix_Timer()
frmMain.Nix.Visible = False
Tnix.Enabled = False
End Sub

Private Sub Ttale_Timer()
frmMain.Tale.Visible = False
Ttale.Enabled = False
End Sub

Private Sub Tulla_Timer()
frmMain.Ulla.Visible = False
Tulla.Enabled = False
End Sub

'
' -------------------
'    W I N S O C K
' -------------------
'

#If UsarWrench <> 1 Then

Private Sub Winsock1_Close()
    Dim I As Long
    
    Debug.Print "WInsock Close"
    
    LastSecond = 0
    Second.Enabled = False
    logged = False
    Connected = False
    
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    frmConnect.MousePointer = vbNormal
    
    If frmPasswdSinPadrinos.Visible = True Then frmPasswdSinPadrinos.Visible = False
    frmCrearPersonaje.Visible = False
    frmConnect.Visible = True
    
    On Local Error Resume Next
    For I = 0 To Forms.count - 1
        If Forms(I).Name <> Me.Name And Forms(I).Name <> frmConnect.Name Then
            Unload Forms(I)
        End If
    Next I
    On Local Error GoTo 0
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False

    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    
    For I = 1 To NUMSKILLS
        UserSkills(I) = 0
    Next I

    For I = 1 To NUMATRIBUTOS
        UserAtributos(I) = 0
    Next I

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Winsock1_Connect()
    Dim ServerIp As String
    Dim Temporal1 As Long
    Dim Temporal As Long
    
    Debug.Print "Winsock Connect"
    
    ServerIp = Winsock1.RemoteHostIP
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = ((Mid$(ServerIp, 1, Temporal - 1) Xor &H65) And &H7F) * 16777216
    ServerIp = Mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (Mid$(ServerIp, 1, Temporal - 1) Xor &HF6) * 65536
    ServerIp = Mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (Mid$(ServerIp, 1, Temporal - 1) Xor &H4B) * 256
    ServerIp = Mid$(ServerIp, Temporal + 1, Len(ServerIp)) Xor &H42
    MixedKey = (Temporal1 + ServerIp)
    
    Second.Enabled = True
    
    'If frmCrearPersonaje.Visible Then
    If EstadoLogin = E_MODO.CrearNuevoPj Then
        Call SendData("gIvEmEvAlcOde")
    'ElseIf Not frmRecuperar.Visible Then
    ElseIf EstadoLogin = E_MODO.Normal Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = E_MODO.Dados Then
        Call SendData("gIvEmEvAlcOde")
    'Else
    ElseIf EstadoLogin = E_MODO.RecuperarPass Then
        Dim cmd As String
        cmd = "PASSRECO" & frmRecuperar.txtNombre.Text & "~" & frmRecuperar.txtCorreo
        'frmMain.Socket1.Write cmd$, Len(cmd$)
        'Call SendData(cmd$)
    End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim loopc As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim echar As Integer
    Dim aux$
    Dim nfile As Integer

    Debug.Print "Winsock DataArrival"
    
    'Socket1.Read RD, DataLength
    Winsock1.GetData RD

    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    'Check for more than one line
    sChar = 1
    For loopc = 1 To Len(RD)

        tChar = Mid$(RD, loopc, 1)

        If tChar = ENDC Then
            CR = CR + 1
            echar = loopc - sChar
            rBuffer(CR) = Mid$(RD, sChar, echar)
            sChar = loopc + 1
        End If

    Next loopc

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = Mid$(RD, sChar, Len(RD))
    End If

    'Send buffer to Handle data
    For loopc = 1 To CR
        Call HandleData(rBuffer(loopc))
    Next loopc
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    LastSecond = 0
    Second.Enabled = False

    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    If frmOldPersonaje.Visible Then
        frmOldPersonaje.Visible = False
    End If

    If Not frmCrearPersonaje.Visible Then
        If Not frmBorrar.Visible And Not frmRecuperar.Visible Then
            frmConnect.Show
        End If
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub
#End If
Private Sub WsAntiSH_DataArrival(ByVal bytesTotal As Long)
Dim hora As String
Dim Horas As Long
Dim Minutos As Long
Dim Segundos As Long
Dim ret As Long
 On Error Resume Next
WSAntiSH.GetData hora, vbString
'List1.AddItem hora
WSAntiSH.Close
Horas = Val(Mid$(hora, 17, 2))
Minutos = Val(Mid$(hora, 20, 2))
Segundos = Val(Mid$(hora, 23, 2))
ret = Val(Mid$(hora, 33, 3)) + (Segundos + (Minutos + Horas * 60) * 60) * 1000 - GetTickCount
'AddTime Ret
 
End Sub
Private Sub Minimap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Call SendData("/TELEP YO " & UserMap & " " & CByte(X) & " " & CByte(Y))
End Sub
