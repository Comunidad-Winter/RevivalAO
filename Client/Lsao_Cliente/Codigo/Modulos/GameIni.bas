Attribute VB_Name = "GameIni"



Option Explicit
Public MapaBicho As Integer
Public Type tCabecera 'Cabecera de los con
    desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Type tGameIni
    Puerto As Long
    Musica As Byte
    Fx As Byte
    tip As Byte
    Password As String
    Name As String
    DirGraficos As String
    DirSonidos As String
    DirMusica As String
    DirMapas As String
    NumeroDeBMPs As Long
    NumeroMapas As Integer
End Type

Public Type tSetupMods
    bDinamic    As Boolean
    byMemory    As Byte
    bUseVideo   As Boolean
    bNoMusic    As Boolean
    bNoSound    As Boolean
    bFPS        As Byte
End Type

Public ClientSetup As tSetupMods

Public MiCabecera As tCabecera
Public Config_Inicio As tGameIni


Public Sub IniciarCabecera(ByRef Cabecera As tCabecera)
    Cabecera.desc = "RevivalAo by Noland Studios. Copyright Noland-Studios 2001, pablomarquez@noland-studios.com.ar"
    Cabecera.CRC = Rnd * 100
    Cabecera.MagicWord = Rnd * 10
End Sub

Public Function LeerGameIni() As tGameIni
    Dim n As Integer
    Dim GameIni As tGameIni
    

    n = FreeFile
    Open App.Path & "\..\Recursos\GameIni.revival" For Binary As #n
    Get #n, , MiCabecera
    
    Get #n, , GameIni
    
    Close #n
    LeerGameIni = GameIni
End Function

Public Sub EscribirGameIni(ByRef GameIniConfiguration As tGameIni)
On Local Error Resume Next

Dim n As Integer
n = FreeFile
Open App.Path & "\..\Recursos\GameIni.revival" For Binary As #n
Put #n, , MiCabecera
Put #n, , GameIniConfiguration
Close #n

End Sub

