VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Compresor de recursos graficos"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11910
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   199
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   794
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Comprimir MIDI"
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Extraer MIDI"
      Height          =   375
      Left            =   8040
      TabIndex        =   13
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Extraer WAV"
      Height          =   375
      Left            =   8040
      TabIndex        =   12
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Comprimir WAV"
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Comprimir MAPAS"
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Extraer MAPAS"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Extraer INIT"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Comprimir INIT"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Frame StatusFrame 
      Caption         =   "StatusFrame"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   5415
      Begin MSComctlLib.ProgressBar StatusBar 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin VB.CommandButton cmdPatch 
      Caption         =   "Parchear"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extraer"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtVersion 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3720
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdCompress 
      Caption         =   "Comprimir"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Working Version :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCompress_Click()
    Dim SourcePath As String
    Dim OutputPath As String
    
    SourcePath = App.Path & GRAPHIC_PATH
    OutputPath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
    
    'Check if the version already exists
    If FileExist(OutputPath & GRH_RESOURCE_FILE, vbNormal) Then
        If MsgBox("La versión ya se encuentra comprimida. ¿Desea reemplazarla?", vbYesNo, "Atencion") = vbNo Then _
            Exit Sub
    Else
        If Not FileExist(OutputPath, vbDirectory) Then
            'Create this version folder
            MkDir OutputPath
        End If
    End If
    
    'Show status
    Me.Height = 2880
    StatusFrame.Caption = "Comprimiendo..."
    
    'Compress!
    If Compress_Files(SourcePath, OutputPath, txtVersion.Text, StatusBar, modCompression.GRH_RESOURCE_FILE) Then
        'Show we finished
        MsgBox "Operación terminada con éxito"
    Else
        'Show we finished
        MsgBox "Operación abortada"
    End If
    
    'Hide status
    Me.Height = 2055
End Sub

Private Sub cmdExtract_Click()
    Dim ResourcePath As String
    Dim OutputPath As String

    ResourcePath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
    OutputPath = App.Path & EXTRACT_PATH & txtVersion.Text & "\"
    
    'Check if the resource file exists
    If Not FileExist(ResourcePath & GRH_RESOURCE_FILE, vbNormal) Then
        MsgBox "No se encontraron los recursos a extraer." & vbCrLf & ResourcePath, , "Error"
        Exit Sub
    End If
    
    'Check if the version is already extracted
    If FileExist(OutputPath, vbDirectory) Then
        If MsgBox("La versión ya se encuentra extraida. ¿Desea reextraerla?", vbYesNo, "Atencion") = vbNo Then _
            Exit Sub
    Else
        'Create this version folder
        MkDir OutputPath
    End If
    
    'Show the status bar
    Me.Height = 2880
    StatusFrame.Caption = "Extrayendo..."
    
    'Extract!
    If Extract_Files(ResourcePath, OutputPath, StatusBar, GRH_RESOURCE_FILE) Then
        'Show we finished
        MsgBox "Operación terminada con éxito"
    Else
        'Show we finished
        MsgBox "Operación abortada"
    End If
    
    'Hide status
    Me.Height = 2055
End Sub

Private Sub cmdPatch_Click()
    Dim NewResourcePath As String
    Dim OldResourcePath As String
    Dim OutputPath As String
    
    Dim NewVersion As Long
    Dim OldVersion As Long
    
    NewVersion = CLng(txtVersion.Text)
    OldVersion = NewVersion - 1 'we patch from the last version
    
    NewResourcePath = App.Path & RESOURCE_PATH & NewVersion & "\"
    OldResourcePath = App.Path & RESOURCE_PATH & OldVersion & "\"
    OutputPath = App.Path & PATCH_PATH & OldVersion & " to " & NewVersion & "\"
    
    'Check if the new resource file exists
    If Not FileExist(NewResourcePath & GRH_RESOURCE_FILE, vbNormal) Then
        MsgBox "No se encontraron los recursos de la version actual." & vbCrLf & NewResourcePath, , "Error"
        Exit Sub
    End If
    
    'Check if the old resource file exists
    If Not FileExist(OldResourcePath & GRH_RESOURCE_FILE, vbNormal) Then
        MsgBox "No se encontraron los recursos de la version anterior." & vbCrLf & OldResourcePath, , "Error"
        Exit Sub
    End If
    
    'Check if the version is already extracted
    If FileExist(OutputPath, vbDirectory) Then
        If MsgBox("El parche ya se encuentra realizado. ¿Desea reparchear?", vbYesNo, "Atencion") = vbNo Then _
            Exit Sub
    Else
        'Create this version folder
        MkDir OutputPath
    End If
    
    'Show the status bar
    Me.Height = 2880
    StatusFrame.Caption = "Armando el parche de " & OldVersion & " a " & NewVersion
    
    'Patch!
    If Make_Patch(NewResourcePath, OldResourcePath, OutputPath, StatusBar) Then
        'Show we finished
        MsgBox "Operación terminada con éxito"
    Else
        'Show we finished
        MsgBox "Operación abortada"
    End If
    
    'Hide status
    Me.Height = 2055
End Sub


Private Sub command1_Click()
    Dim SourcePath As String
    Dim OutputPath As String
    
    SourcePath = App.Path & INIT_PATH
    OutputPath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
    
    'Check if the version already exists
    If FileExist(OutputPath & INIT_RESOURCE_FILE, vbNormal) Then
        If MsgBox("La versión ya se encuentra comprimida. ¿Desea reemplazarla?", vbYesNo, "Atencion") = vbNo Then _
            Exit Sub
    Else
        If Not FileExist(OutputPath, vbDirectory) Then
            'Create this version folder
            MkDir OutputPath
        End If
    End If
    
    'Show status
    Me.Height = 2880
    StatusFrame.Caption = "Comprimiendo..."
    
    'Compress!
    If Compress_Files(SourcePath, OutputPath, txtVersion.Text, StatusBar, INIT_RESOURCE_FILE, ".*") Then
        'Show we finished
        MsgBox "Operación terminada con éxito"
    Else
        'Show we finished
        MsgBox "Operación abortada"
    End If
    
    'Hide status
    Me.Height = 2055
End Sub

Private Sub Command2_Click()
    Dim ResourcePath As String
    Dim OutputPath As String

    ResourcePath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
    OutputPath = App.Path & EXTRACT_PATH & txtVersion.Text & "-Init\"
    
    'Check if the resource file exists
    If Not FileExist(ResourcePath & INIT_RESOURCE_FILE, vbNormal) Then
        MsgBox "No se encontraron los recursos a extraer." & vbCrLf & ResourcePath, , "Error"
        Exit Sub
    End If
    
    'Check if the version is already extracted
    If FileExist(OutputPath, vbDirectory) Then
        If MsgBox("La versión ya se encuentra extraida. ¿Desea reextraerla?", vbYesNo, "Atencion") = vbNo Then _
            Exit Sub
    Else
        'Create this version folder
        MkDir OutputPath
    End If
    
    'Show the status bar
    Me.Height = 2880
    StatusFrame.Caption = "Extrayendo..."
    
    'Extract!
    If Extract_Files(ResourcePath, OutputPath, StatusBar, INIT_RESOURCE_FILE) Then
        'Show we finished
        MsgBox "Operación terminada con éxito"
    Else
        'Show we finished
        MsgBox "Operación abortada"
    End If
    
    'Hide status
    Me.Height = 2055
End Sub

Private Sub Command3_Click()
    Dim ResourcePath As String
    Dim OutputPath As String

    ResourcePath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
    OutputPath = App.Path & EXTRACT_PATH & txtVersion.Text & "-Mapas\"
    
    'Check if the resource file exists
    If Not FileExist(ResourcePath & MIDI_RESOURCE_FILE, vbNormal) Then
        MsgBox "No se encontraron los recursos a extraer." & vbCrLf & ResourcePath, , "Error"
        Exit Sub
    End If
    
    'Check if the version is already extracted
    If FileExist(OutputPath, vbDirectory) Then
        If MsgBox("La versión ya se encuentra extraida. ¿Desea reextraerla?", vbYesNo, "Atencion") = vbNo Then _
            Exit Sub
    Else
        'Create this version folder
        MkDir OutputPath
    End If
    
    'Show the status bar
    Me.Height = 2880
    StatusFrame.Caption = "Extrayendo..."
    
    'Extract!
    If Extract_Files(ResourcePath, OutputPath, StatusBar, MAPAS_RESOURCE_FILE) Then
        'Show we finished
        MsgBox "Operación terminada con éxito"
    Else
        'Show we finished
        MsgBox "Operación abortada"
    End If
    
    'Hide status
    Me.Height = 2055
End Sub

Private Sub Command4_Click()
    Dim SourcePath As String
    Dim OutputPath As String
    
    SourcePath = App.Path & MAP_PATH
    OutputPath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
    
    'Check if the version already exists
    If FileExist(OutputPath & MAPAS_RESOURCE_FILE, vbNormal) Then
        If MsgBox("La versión ya se encuentra comprimida. ¿Desea reemplazarla?", vbYesNo, "Atencion") = vbNo Then _
            Exit Sub
    Else
        If Not FileExist(OutputPath, vbDirectory) Then
            'Create this version folder
            MkDir OutputPath
        End If
    End If
    
    'Show status
    Me.Height = 2880
    StatusFrame.Caption = "Comprimiendo..."
    
    'Compress!
    If Compress_Files(SourcePath, OutputPath, txtVersion.Text, StatusBar, MAPAS_RESOURCE_FILE, ".MAP") Then
        'Show we finished
        MsgBox "Operación terminada con éxito"
    Else
        'Show we finished
        MsgBox "Operación abortada"
    End If
    
    'Hide status
    Me.Height = 2055
End Sub

Private Sub Command5_Click()
    Dim SourcePath As String
    Dim OutputPath As String
    
    SourcePath = App.Path & WAV_PATH
    OutputPath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
    
    'Check if the version already exists
    If FileExist(OutputPath & WAV_RESOURCE_FILE, vbNormal) Then
        If MsgBox("La versión ya se encuentra comprimida. ¿Desea reemplazarla?", vbYesNo, "Atencion") = vbNo Then _
            Exit Sub
    Else
        If Not FileExist(OutputPath, vbDirectory) Then
            'Create this version folder
            MkDir OutputPath
        End If
    End If
    
    'Show status
    Me.Height = 2880
    StatusFrame.Caption = "Comprimiendo..."
    
    'Compress!
    If Compress_Files(SourcePath, OutputPath, txtVersion.Text, StatusBar, WAV_RESOURCE_FILE, ".WAV") Then
        'Show we finished
        MsgBox "Operación terminada con éxito"
    Else
        'Show we finished
        MsgBox "Operación abortada"
    End If
    
    'Hide status
    Me.Height = 2055
End Sub

Private Sub Command6_Click()
    Dim ResourcePath As String
    Dim OutputPath As String

    ResourcePath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
    OutputPath = App.Path & EXTRACT_PATH & txtVersion.Text & "-WAV\"
    
    'Check if the resource file exists
    If Not FileExist(ResourcePath & WAV_RESOURCE_FILE, vbNormal) Then
        MsgBox "No se encontraron los recursos a extraer." & vbCrLf & ResourcePath, , "Error"
        Exit Sub
    End If
    
    'Check if the version is already extracted
    If FileExist(OutputPath, vbDirectory) Then
        If MsgBox("La versión ya se encuentra extraida. ¿Desea reextraerla?", vbYesNo, "Atencion") = vbNo Then _
            Exit Sub
    Else
        'Create this version folder
        MkDir OutputPath
    End If
    
    'Show the status bar
    Me.Height = 2880
    StatusFrame.Caption = "Extrayendo..."
    
    'Extract!
    If Extract_Files(ResourcePath, OutputPath, StatusBar, WAV_RESOURCE_FILE) Then
        'Show we finished
        MsgBox "Operación terminada con éxito"
    Else
        'Show we finished
        MsgBox "Operación abortada"
    End If
    
    'Hide status
    Me.Height = 2055
End Sub

Private Sub Command7_Click()
    Dim ResourcePath As String
    Dim OutputPath As String

    ResourcePath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
    OutputPath = App.Path & EXTRACT_PATH & txtVersion.Text & "-Midi\"
    
    'Check if the resource file exists
    If Not FileExist(ResourcePath & MIDI_RESOURCE_FILE, vbNormal) Then
        MsgBox "No se encontraron los recursos a extraer." & vbCrLf & ResourcePath, , "Error"
        Exit Sub
    End If
    
    'Check if the version is already extracted
    If FileExist(OutputPath, vbDirectory) Then
        If MsgBox("La versión ya se encuentra extraida. ¿Desea reextraerla?", vbYesNo, "Atencion") = vbNo Then _
            Exit Sub
    Else
        'Create this version folder
        MkDir OutputPath
    End If
    
    'Show the status bar
    Me.Height = 2880
    StatusFrame.Caption = "Extrayendo..."
    
    'Extract!
    If Extract_Files(ResourcePath, OutputPath, StatusBar, MIDI_RESOURCE_FILE) Then
        'Show we finished
        MsgBox "Operación terminada con éxito"
    Else
        'Show we finished
        MsgBox "Operación abortada"
    End If
    
    'Hide status
    Me.Height = 2055
End Sub

Private Sub Command8_Click()
    Dim SourcePath As String
    Dim OutputPath As String
    
    SourcePath = App.Path & MIDI_PATH
    OutputPath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
    
    'Check if the version already exists
    If FileExist(OutputPath & MIDI_RESOURCE_FILE, vbNormal) Then
        If MsgBox("La versión ya se encuentra comprimida. ¿Desea reemplazarla?", vbYesNo, "Atencion") = vbNo Then _
            Exit Sub
    Else
        If Not FileExist(OutputPath, vbDirectory) Then
            'Create this version folder
            MkDir OutputPath
        End If
    End If
    
    'Show status
    Me.Height = 2880
    StatusFrame.Caption = "Comprimiendo..."
    
    'Compress!
    If Compress_Files(SourcePath, OutputPath, txtVersion.Text, StatusBar, MIDI_RESOURCE_FILE, ".mid") Then
        'Show we finished
        MsgBox "Operación terminada con éxito"
    Else
        'Show we finished
        MsgBox "Operación abortada"
    End If
    
    'Hide status
    Me.Height = 2055
End Sub

Private Sub Command9_Click()
'**************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 12/02/2007
'Loads all surfaces in random order and then sorts them
'**************************************************************
'On Error Resume Next
    Dim surfaceIndex As Long
    Dim bmpInfo As BITMAPINFO
    Dim data() As Byte
    Dim I As Long
    Dim ResourcePath As String
    ResourcePath = App.Path & RESOURCE_PATH & txtVersion.Text & "\"
    While GetNext_Bitmap(ResourcePath, I, bmpInfo, data(), surfaceIndex)
        DoEvents
    Wend
    
    Debug.Print "Listo."
End Sub

