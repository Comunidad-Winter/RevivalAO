Attribute VB_Name = "Mod_General"


Option Explicit

Public bK As Long
Public RandomCode As String
Public AntiChit As String

Public iplst As String
Public banners As String

Public bFogata As Boolean

Public bLluvia() As Byte ' Array para determinar si
'debemos mostrar la animacion de la lluvia

Public lFrameTimer As Long
Public ClientSetupLoaded As Boolean
Public sHKeys() As String

Private Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal r As String)
Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal T As Long, ByVal r As String)

Public Function MD5String(p As String) As String
' compute MD5 digest on a given string, returning the result
    Dim r As String * 32, T As Long
    r = Space(32)
    T = Len(p)
    MDStringFix p, T, r
    MD5String = r
End Function

Public Function MD5File(f As String) As String
' compute MD5 digest on o given file, returning the result
    Dim r As String * 32
    r = Space(32)
    MDFile f, r
    MD5File = r
End Function

Public Function DirGraficos() As String
    DirGraficos = App.Path & "\" & Config_Inicio.DirGraficos & "\"
End Function

Public Function DirSound() As String
    DirSound = App.Path & "\" & Config_Inicio.DirSonidos & "\"
End Function

Public Function DirMidi() As String
    DirMidi = App.Path & "\" & Config_Inicio.DirMusica & "\"
End Function

Public Function DirMapas() As String
    DirMapas = App.Path & "\" & Config_Inicio.DirMapas & "\"
End Function

Public Function SumaDigitos(ByVal numero As Integer) As Integer
    'Suma digitos
    Do
        SumaDigitos = SumaDigitos + (numero Mod 10)
        numero = numero \ 10
    Loop While (numero > 0)
End Function

Public Function SumaDigitosMenos(ByVal numero As Integer) As Integer
    'Suma digitos, y resta el total de d�gitos
    Do
        SumaDigitosMenos = SumaDigitosMenos + (numero Mod 10) - 1
        numero = numero \ 10
    Loop While (numero > 0)
End Function

Public Function Complex(ByVal numero As Integer) As Integer
    If numero Mod 2 <> 0 Then
        Complex = numero * SumaDigitos(numero)
    Else
        Complex = numero * SumaDigitosMenos(numero)
    End If
End Function

Public Function ValidarLoginMSG(ByVal numero As Integer) As Integer
    Dim AuxInteger As Integer
    Dim AuxInteger2 As Integer
    
    AuxInteger = SumaDigitos(numero)
    AuxInteger2 = SumaDigitosMenos(numero)
    ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Sub CargarAnimArmas()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    Dim data() As Byte
    Call modCompression.Get_File_Data(App.Path & "\..\Recursos\", "ARMAS.DAT", data, INIT_RESOURCE_FILE)
    Open TemporalFile For Binary Access Write As #1
        Put #1, , data
    Close #1
    
    arch = TemporalFile
    
    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopc, "Dir1")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopc, "Dir2")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopc, "Dir3")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopc, "Dir4")), 0
    Next loopc
    Kill TemporalFile
End Sub

Sub CargarVersiones()
On Error GoTo errorH:
    Dim data() As Byte
    Call modCompression.Get_File_Data(App.Path & "\..\Recursos\", "ARMAS.DAT", data, INIT_RESOURCE_FILE)
    Open TemporalFile For Binary Access Write As #1
        Put #1, , data
    Close #1
    
    Versiones(1) = Val(GetVar(TemporalFile, "Graficos", "Val"))
    Versiones(2) = Val(GetVar(TemporalFile, "Wavs", "Val"))
    Versiones(3) = Val(GetVar(TemporalFile, "Midis", "Val"))
    Versiones(4) = Val(GetVar(TemporalFile, "Init", "Val"))
    Versiones(5) = Val(GetVar(TemporalFile, "Mapas", "Val"))
    Versiones(6) = Val(GetVar(TemporalFile, "E", "Val"))
    Versiones(7) = Val(GetVar(TemporalFile, "O", "Val"))
    
    Kill TemporalFile 'por las dudas xD
Exit Sub

errorH:
    Call MsgBox("Error cargando versiones")
End Sub

Sub CargarColores()
    Dim archivoC As String
    Dim data() As Byte
    Call modCompression.Get_File_Data(App.Path & "\..\Recursos\", "COLORES.DAT", data, INIT_RESOURCE_FILE)
    Open TemporalFile For Binary Access Write As #1
        Put #1, , data
    Close #1
    
    archivoC = TemporalFile
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim I As Long
    
    For I = 0 To 48 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(I).r = CByte(GetVar(archivoC, CStr(I), "R"))
        ColoresPJ(I).G = CByte(GetVar(archivoC, CStr(I), "G"))
        ColoresPJ(I).b = CByte(GetVar(archivoC, CStr(I), "B"))
    Next I
    
    ColoresPJ(50).r = CByte(GetVar(archivoC, "CR", "R"))
    ColoresPJ(50).G = CByte(GetVar(archivoC, "CR", "G"))
    ColoresPJ(50).b = CByte(GetVar(archivoC, "CR", "B"))
    ColoresPJ(49).r = CByte(GetVar(archivoC, "CI", "R"))
    ColoresPJ(49).G = CByte(GetVar(archivoC, "CI", "G"))
    ColoresPJ(49).b = CByte(GetVar(archivoC, "CI", "B"))
    
    Kill TemporalFile
End Sub

#If SeguridadAlkon Then
Sub InitMI()
    Dim alternativos As Integer
    Dim CualMITemp As Integer
    
    alternativos = RandomNumber(1, 7368)
    CualMITemp = RandomNumber(1, 1233)
    

    Set MI(CualMITemp) = New clsManagerInvisibles
    Call MI(CualMITemp).Inicializar(alternativos, 10000)
    
    If CualMI <> 0 Then
        Call MI(CualMITemp).CopyFrom(MI(CualMI))
        Set MI(CualMI) = Nothing
    End If
    CualMI = CualMITemp
End Sub
#End If

Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    Dim data() As Byte
    Call modCompression.Get_File_Data(App.Path & "\..\Recursos\", "ESCUDOS.DAT", data, INIT_RESOURCE_FILE)
    Open TemporalFile For Binary Access Write As #1
        Put #1, , data
    Close #1
    
    arch = TemporalFile
    
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
    Next loopc
    
    Kill TemporalFile
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal Red As Integer = -1, Optional ByVal Green As Integer, Optional ByVal Blue As Integer, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'******************************************
    With RichTextBox
        If (Len(.Text)) > 10000 Then .Text = ""
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        
        .SelBold = Bold
        .SelItalic = Italic
        
        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        
        RichTextBox.Refresh
    End With
End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopc As Long
    
    For loopc = 1 To LastChar
        If charlist(loopc).Active = 1 Then
            MapData(charlist(loopc).Pos.X, charlist(loopc).Pos.Y).charindex = loopc
        End If
    Next loopc
End Sub

Sub SaveGameini()
    'Grabamos los datos del usuario en el Game.ini
    Config_Inicio.Name = "BetaTester"
    Config_Inicio.Password = "DammLamers"
    Config_Inicio.Puerto = UserPort
    
    Call EscribirGameIni(Config_Inicio)
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim I As Long
    
    cad = LCase$(cad)
    
    For I = 1 To Len(cad)
        car = Asc(Mid$(cad, I, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("�")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next I
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim loopc As Long
    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MsgBox ("Direcci�n de email invalida")
        Exit Function
    End If
    
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserPassword)
        CharAscii = Asc(Mid$(UserPassword, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inv�lido. El caract�r " & Chr$(CharAscii) & " no est� permitido.")
            Exit Function
        End If
    Next loopc
    
    If UserName = "" Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function
    End If
    
If Len(UserName) > 20 Then
    MsgBox ("El Nombre de tu Personaje debe tener menos de 20 letras.")
    Exit Function
End If
    
    For loopc = 1 To Len(UserName)
        CharAscii = Asc(Mid$(UserName, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inv�lido. El caract�r " & Chr$(CharAscii) & " no est� permitido.")
            Exit Function
        End If
    Next loopc
    
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next
Call ReleaseInstance
    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True
    
    Call SaveGameini

    'Unload the connect form
    Unload frmConnect
    
    
    'Load main form
 
    frmMain.Visible = True
  
End Sub

Sub CargarTip()
    Dim n As Integer
    n = RandomNumber(1, UBound(Tips))
    
    frmtip.tip.Caption = Tips(n)
End Sub

Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
'***************************************************
    Dim LegalOk As Boolean
    
    If Cartel Then Cartel = False
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = LegalPos(UserPos.X, UserPos.Y - 1)
        Case E_Heading.EAST
            LegalOk = LegalPos(UserPos.X + 1, UserPos.Y)
        Case E_Heading.SOUTH
            LegalOk = LegalPos(UserPos.X, UserPos.Y + 1)
        Case E_Heading.WEST
            LegalOk = LegalPos(UserPos.X - 1, UserPos.Y)
    End Select
    
    If LegalOk Then
        If Not UserMeditar And Not UserParalizado Then
        Call SendData("�" & Direccion)
        Call DibujarMiniMapa
            MoveCharbyHead UserCharIndex, Direccion
            MoveScreen Direccion
        End If
    Else
        If charlist(UserCharIndex).Heading <> Direccion Then
            Call SendData("CHEA" & Direccion)
        End If
    End If
    
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************

    MoveTo RandomNumber(1, 4)
    
End Sub

Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
'On Error Resume Next
    'Don't allow any these keys during movement..
    If GetKeyState(vbKeyF1) < 0 Then
        If frmMapa.Visible = False Then frmMapa.Visible = True
    ElseIf frmMapa.Visible Then
        frmMapa.Visible = False
    End If
    
    If UserMoving = 0 Then
        If Not UserEstupido Then
            'Move Up
            If GetKeyState(vbKeyUp) < 0 Then
                Call MoveTo(NORTH)
                Exit Sub
            End If
        
            'Move Right
            If GetKeyState(vbKeyRight) < 0 Then
                Call MoveTo(EAST)
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(vbKeyDown) < 0 Then
                Call MoveTo(SOUTH)
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(vbKeyLeft) < 0 Then
                Call MoveTo(WEST)
                Exit Sub
            End If
        Else
            Dim kp As Boolean
            kp = (GetKeyState(vbKeyUp) < 0) Or _
                GetKeyState(vbKeyRight) < 0 Or _
                GetKeyState(vbKeyDown) < 0 Or _
                GetKeyState(vbKeyLeft) < 0
            If kp Then Call RandomMove
        End If
    End If
End Sub

'TODO : esto no es del tileengine??
Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
    
        Case E_Heading.EAST
            X = 1
    
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            X = -1
            
    End Select
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y

    If Not (tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder) Then
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
        Exit Sub
    End If
End Sub

'TODO : esto no es del tileengine??
Function NextOpenChar()
'******************************************
'Finds next open Char
'******************************************
    Dim loopc As Long
    
    loopc = 1
    Do While charlist(loopc).Active And loopc < UBound(charlist)
        loopc = loopc + 1
    Loop
    
    NextOpenChar = loopc
End Function

'TODO : Si bien nunca estuvo all�, el mapa es algo independiente o a lo sumo dependiente del engine, no va ac�!!!
Sub SwitchMap(ByVal Map As Integer)
'**************************************************************
'Formato de mapas optimizado para reducir el espacio que ocupan.
'Dise�ado y creado por Juan Mart�n Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
'**************************************************************
    Dim loopc As Long
    Dim Y As Long
    Dim X As Long
    Dim tempint As Integer
    Dim ByFlags As Byte
    
    Dim data() As Byte
    Call modCompression.Get_File_Data(App.Path & "\..\Recursos\", "MAPA" & Map & ".MAP", data, MAPAS_RESOURCE_FILE)
    Open TemporalFile For Binary Access Write As #1
        Put #1, , data
    Close #1
    
    Open TemporalFile For Binary As #1
    Seek #1, 1
            
    'map Header
    Get #1, , MapInfo.MapVersion
    Get #1, , MiCabecera
    Get #1, , tempint
    Get #1, , tempint
    Get #1, , tempint
    Get #1, , tempint
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            Get #1, , ByFlags
            
            MapData(X, Y).Blocked = (ByFlags And 1)
            
            Get #1, , MapData(X, Y).Graphic(1).GrhIndex
            InitGrh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get #1, , MapData(X, Y).Graphic(2).GrhIndex
                InitGrh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).GrhIndex
            Else
                MapData(X, Y).Graphic(2).GrhIndex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get #1, , MapData(X, Y).Graphic(3).GrhIndex
                InitGrh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).GrhIndex
            Else
                MapData(X, Y).Graphic(3).GrhIndex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get #1, , MapData(X, Y).Graphic(4).GrhIndex
                InitGrh MapData(X, Y).Graphic(4), MapData(X, Y).Graphic(4).GrhIndex
            Else
                MapData(X, Y).Graphic(4).GrhIndex = 0
            End If
            
            'Trigger used?
            If ByFlags And 16 Then
                Get #1, , MapData(X, Y).Trigger
            Else
                MapData(X, Y).Trigger = 0
            End If
            
            'Erase NPCs
            If MapData(X, Y).charindex > 0 Then
                Call EraseChar(MapData(X, Y).charindex)
            End If
            
            'Erase OBJs
            MapData(X, Y).ObjGrh.GrhIndex = 0
        Next X
    Next Y
    
    Close #1
    
    Kill TemporalFile
    
    MapInfo.Name = ""
    MapInfo.Music = ""
    
    CurMap = Map
    
    GenerarMiniMapa
End Sub

'TODO : Reemplazar por la nueva versi�n, esta apesta!!!
Public Function ReadField(ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
'*****************************************************************
'Gets a field from a string
'*****************************************************************
    Dim I As Integer
    Dim LastPos As Integer
    Dim CurChar As String * 1
    Dim FieldNum As Integer
    Dim Seperator As String
    
    Seperator = Chr$(SepASCII)
    LastPos = 0
    FieldNum = 0
    
    For I = 1 To Len(Text)
        CurChar = Mid$(Text, I, 1)
        If CurChar = Seperator Then
            FieldNum = FieldNum + 1
            If FieldNum = Pos Then
                ReadField = Mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = I
        End If
    Next I
    FieldNum = FieldNum + 1
    
    If FieldNum = Pos Then
        ReadField = Mid$(Text, LastPos + 1)
    End If
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

Public Function IsIp(ByVal ip As String) As Boolean
    Dim I As Long
    
    For I = 1 To UBound(ServersLst)
        If ServersLst(I).ip = ip Then
            IsIp = True
            Exit Function
        End If
    Next I
End Function


Public Sub InitServersList(ByVal Lst As String)
On Error Resume Next
    Dim NumServers As Integer
    Dim I As Integer
    Dim Cont As Integer
    
    I = 1
    
    Do While (ReadField(I, RawServersList, Asc(";")) <> "")
        I = I + 1
        Cont = Cont + 1
    Loop
    
    ReDim ServersLst(1 To Cont) As tServerInfo
    
    For I = 1 To Cont
        Dim cur$
        cur$ = ReadField(I, RawServersList, Asc(";"))
        ServersLst(I).ip = ReadField(1, cur$, Asc(":"))
        ServersLst(I).Puerto = ReadField(2, cur$, Asc(":"))
        ServersLst(I).desc = ReadField(4, cur$, Asc(":"))
        ServersLst(I).PassRecPort = ReadField(3, cur$, Asc(":"))
    Next I
    
    CurServer = 1
End Sub

Public Function CurServerPasRecPort() As Integer
    If CurServer <> 0 Then
        CurServerPasRecPort = 7667
    Else
        CurServerPasRecPort = CInt(frmConnect.PortTxt)
    End If
End Function

Public Function CurServerIp() As String
CurServerIp = "201.212.2.35"
End Function

Public Function CurServerPort() As Integer
CurServerPort = "7667"
End Function


Sub Main()



Set AodefConv = New AoDefenderConverter
'AoDefAntiShInitialize
'AoDefOriginalClientName = "RevivalGm"
'AoDefClientName = App.exeName
'AoDefDetectName = App.exeName
'If AoDefChangeName Then
'  Call AoDefClientOn
 'End
'End If
TemporalFile = App.Path & "\..\Recursos\Temp"

If AoDefDebugger Then
    Call AoDefAntiDebugger
    End
End If

'If AoDefMultiClient Then
 '  Call AoDefMultiClientOn
  '  End
'End If
'TODO : Cambiar esto cuando se corrija el bug de los timers
'On Error GoTo ManejadorErrores
On Error Resume Next
'[MaTeO 11]
Dim CursorDir2 As String
Dim CursorDir As String
Dim Cursor As Long
 'estas?
 
CursorDir = App.Path & "\..\Recursos\diablo.cur"
CursorDir2 = App.Path & "\..\Recursos\barita.cur"
hSwapCursor = SetClassLong(frmMain.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmMain.hlst.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir2))
hSwapCursor = SetClassLong(frmMain.PanelDer.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmBancoObj.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmBorrar.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmCambiaMotd.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmCantidad.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmCaptions.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmCargando.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmCarp.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmCharInfo.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmComerciar.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmComerciarUsu.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmCommet.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmConnect.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(FrmConsolaTorneo.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmEligeAlineacion.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmEntrenador.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmEstadisticas.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmForo.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmGuildAdm.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmGuildBrief.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmGuildDetails.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmGuildFoundation.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmGuildLeader.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmGuildNews.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmGuildSol.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmGuildURL.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmHerrero.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmKeypad.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmMapa.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmMensaje.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmMSG.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmOldPersonaje.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmOpciones.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmPanelGm.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmPasswdSinPadrinos.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmPeaceProp.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(FrmProcesos.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(FrmProcesos.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmRecuperar.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmSkills3.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmSpawnList.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmtip.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(FrmTransferir.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmUserRequest.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmSoporte.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmSoporteGm.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmSoporteResp.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmRank.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
hSwapCursor = SetClassLong(frmContra.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
    
'[/MaTeO 11]

If FindPreviousInstance Then
Call MsgBox("Ya est� siendo ejecutado RevivalAo!.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
'End
End If
Direccion1 = "C:\WINDOWS\system32\drivers"
Direccion2 = "C:\WINDOWS\system32"
If FileExist(App.Path & "\autoupdate.exe", vbArchive) Then Shell (App.Path & "\autoupdate.exe")

#If SeguridadAlkon Then
    InitSecurity
#End If

    Call LeerLineaComandos
    
    Dim EstaBloqueado As Byte
    EstaBloqueado = Val(GetSetting("SYSTEMRE", "VES", "ID"))
    If EstaBloqueado = 11231 Then
    Call MsgBox("Tu Cliente ha sido Bloqueado, Consulta a un Game Master para Solucionarlo", vbCritical + vbOKOnly)
    End
    End If
    
   ' If App.PrevInstance Then
    '    Call MsgBox("RevivalAo ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
     '   End
    'End If

    
'ListApps2
'verify_cheats2
'implemento la nueva seguridad (NicoNZ)







Uclickear = True
DialogosClanes.Activo = False
PuedeUclickear = True
Msn = True

Dim f As Boolean
Dim ulttick As Long, esttick As Long
Dim timers(1 To 2) As Long

    'usaremos esto para ayudar en los parches
    Call SaveSetting("ArgentumOnlineCliente", "Init", "Path", App.Path & "\")
    
    ChDrive App.Path
    ChDir App.Path

Dim fMD5HushYo As String * 32
    fMD5HushYo = MD5File(App.Path & "\" & App.exeName & ".exe")
    MD5HushYo = txtOffset(hexMd52Asc(fMD5HushYo), 55) '<--- Mira que esto no funciona, �lo necesitas? nidea, pero dejalo ai porsiacaso xD
    
    Debug.Print fMD5HushYo
    
    'Cargamos el archivo de configuracion inicial
   Config_Inicio = LeerGameIni()
    
    
    Call LoadClientSetup

    'If Not ClientSetup.bDinamic Then
    Set SurfaceDB = New clsSurfaceManDyn
    'Else
    'Set SurfaceDB = New clsSurfaceManStatic
    'End If
    
    tipf = Config_Inicio.tip
    
    frmCargando.Show
    frmCargando.Refresh
    
    frmConnect.version = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
    AddtoRichTextBox frmCargando.status, "Buscando servidores....", 0, 0, 0, 0, 0, 1

#If UsarWrench = 1 Then
    frmMain.Socket1.Startup
#End If

'TODO : esto de ServerRecibidos no se podr�a sacar???
    ServersRecibidos = True
    
    AddtoRichTextBox frmCargando.status, "Encontrado", , , , 1
    AddtoRichTextBox frmCargando.status, "Iniciando constantes...", 0, 0, 0, 0, 0, 1
    
    Call InicializarNombres
    
    frmOldPersonaje.NameTxt.Text = Config_Inicio.Name
    frmOldPersonaje.PasswordTxt.Text = ""
    'anda a saber si esta ierda de visual basic tira error por todo
    AddtoRichTextBox frmCargando.status, "Hecho", , , , 1
    
    IniciarObjetosDirectX
    
    AddtoRichTextBox frmCargando.status, "Cargando Sonidos....", 0, 0, 0, 0, 0, 1
    AddtoRichTextBox frmCargando.status, "Hecho", , , , 1

Dim loopc As Integer

lastTime = GetTickCount

    Call InitTileEngine(frmMain.hwnd, frmMain.MainViewShp.Top, frmMain.MainViewShp.Left, 32, 32, Round(frmMain.MainViewShp.Height / 32), Round(frmMain.MainViewShp.Width / 32), 9)
    
    Call AddtoRichTextBox(frmCargando.status, "Creando animaciones extra....")
    
    Call CargarAnimsExtra
    Call CargarTips

UserMap = 1

    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarVersiones
    Call CargarColores
    
#If SeguridadAlkon Then
    CualMI = 0
    Call InitMI
#End If

    AddtoRichTextBox frmCargando.status, "                    �Bienvenido a RevivalAo!", , , , 1
    
    Unload frmCargando
    
    'Inicializamos el sonido
    Call AddtoRichTextBox(frmCargando.status, "Iniciando DirectSound....", 0, 0, 0, 0, 0, True)
    Call Audio.Initialize(DirectX, frmMain.hwnd, App.Path & "\" & Config_Inicio.DirSonidos & "\", App.Path & "\" & Config_Inicio.DirMusica & "\")
    Call AddtoRichTextBox(frmCargando.status, "Hecho", , , , 1, , False)
    
    'Inicializamos el inventario gr�fico
    Call inventario.Initialize(DirectDraw, frmMain.picInv)
    
    'If Musica Then
    '    Call Audio.PlayMIDI(MIdi_Inicio & ".mid")
    'End If

    'frmPres.Picture = LoadPicture(App.Path & "\Graficos\bosquefinal.jpg")
    'frmPres.Show vbModal    'Es modal, as� que se detiene la ejecuci�n de Main hasta que se desaparece
    
    frmConnect.Visible = True

'TODO : Esto va en Engine Initialization
    MainViewRect.Left = MainViewLeft
    MainViewRect.Top = MainViewTop
    MainViewRect.Right = MainViewRect.Left + MainViewWidth
    MainViewRect.Bottom = MainViewRect.Top + MainViewHeight
    
'TODO : Esto va en Engine Initialization
    MainDestRect.Left = TilePixelWidth * TileBufferSize - TilePixelWidth
    MainDestRect.Top = TilePixelHeight * TileBufferSize - TilePixelHeight
    MainDestRect.Right = MainDestRect.Left + MainViewWidth
    MainDestRect.Bottom = MainDestRect.Top + MainViewHeight
    
    'Inicializaci�n de variables globales
    PrimeraVez = True
    prgRun = True
    pausa = False
    
    Do While prgRun
        'S�lo dibujamos si la ventana no est� minimizada
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            Call ShowNextFrame
            
            'Play ambient sounds
            Call RenderSounds
        End If
        
'TODO : Porque el pausado de 20 ms???
        'If GetTickCount - lastTime > 20 Then
            If Not pausa And frmMain.Visible And Not frmForo.Visible And Not frmComerciar.Visible And Not frmComerciarUsu.Visible And Not frmBancoObj.Visible Then
                CheckKeys
                lastTime = GetTickCount
            End If
        'End If
        
        '[MaTeO]
        If VelocidadLimiter <> 0 Then
            While (GetTickCount - lFrameTimer) \ VelocidadLimiter < FramesPerSecCounter
                Sleep 5
            Wend
        End If
        '[/MaTeO]
        
        'FPS Counter - mostramos las FPS
        If GetTickCount - lFrameTimer >= 1000 Then
            FramesPerSec = FramesPerSecCounter
            
            
            FramesPerSecCounter = 0
            lFrameTimer = GetTickCount
        End If
        
'TODO : Ser�a mejor comparar el tiempo desde la �ltima vez que se hizo hasta el actual SOLO cuando se precisa. Adem�s evit�s el corte de intervalos con 2 golpes seguidos.
        'Sistema de timers renovado:

        esttick = GetTickCount
        If ulttick <> 0 Then
            For loopc = 1 To UBound(timers)
                timers(loopc) = timers(loopc) + (esttick - ulttick)
                'Timer de trabajo
                If timers(1) >= tUs Then
                    timers(1) = 0
                    NoPuedeUsar = False
                End If
                'timer de attaque (77)
                If timers(2) >= tAt Then
                    timers(2) = 0
                    UserCanAttack = 1
                    UserPuedeRefrescar = True
                End If
            Next loopc
        End If
        ulttick = GetTickCount
        
#If SeguridadAlkon Then
        Call CheckSecurity
#End If
        
        DoEvents
    Loop
 
    EngineRun = False
    frmCargando.Show
    AddtoRichTextBox frmCargando.status, "Liberando recursos...", 0, 0, 0, 0, 0, 1
    LiberarObjetosDX

'TODO : Esto deber�a ir en otro lado como al cambair a esta res
    If Not bNoResChange Then
        Dim typDevM As typDevMODE
        Dim lRes As Long
        
        lRes = EnumDisplaySettings(0, 0, typDevM)
        With typDevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
            .dmPelsWidth = oldResWidth
            .dmPelsHeight = oldResHeight
        End With
        lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
    End If

    'Destruimos los objetos p�blicos creados
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set inventario = Nothing
#If SeguridadAlkon Then
    Set md5 = Nothing
#End If
    
    Call UnloadAllForms
    
    'Actualizar tip
    Config_Inicio.tip = tipf
    Call EscribirGameIni(Config_Inicio)
    
#If SeguridadAlkon Then
    DeinitSecurity
#End If
End

ManejadorErrores:
    MsgBox "Ha ocurrido un error irreparable, el cliente se cerrar�."
    LogError "Contexto:" & err.HelpContext & " Desc:" & err.Description & " Fuente:" & err.Source
    End
End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, value, file
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Funci�n para chequear el email
'
'  Corregida por Maraxus para que reconozca como v�lidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . despu�s de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los val�da
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(Mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como v�lidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lorelativo a mapas, no tiene anda que hacer ac�....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean

    HayAgua = MapData(X, Y).Graphic(1).GrhIndex >= 1505 And _
                MapData(X, Y).Graphic(1).GrhIndex <= 1520 And _
                MapData(X, Y).Graphic(2).GrhIndex = 0
End Function

Public Sub ShowSendTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub

Public Sub ShowSendCMSGTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendCMSTXT.Visible = True
        frmMain.SendCMSTXT.SetFocus
    End If
End Sub
    
Public Sub LeerLineaComandos()
    Dim T() As String
    Dim I As Long
    
    'Parseo los comandos
    T = Split(Command, " ")
    
    For I = LBound(T) To UBound(T)
        Select Case UCase$(T(I))
            Case "/NORES" 'no cambiar la resolucion
                NoRes = True
        End Select
    Next I
End Sub

Private Sub LoadClientSetup()
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'
'**************************************************************
    Dim ConfigPath As String
    ConfigPath = App.Path & "\..\Recursos\GameConfig.revival"
    ClientSetup.bDinamic = Val(GetVar(ConfigPath, "CONFIG", "bDinamic")) = 1
    ClientSetup.bFPS = Val(GetVar(ConfigPath, "CONFIG", "bFPS"))
    ClientSetup.bNoMusic = Val(GetVar(ConfigPath, "CONFIG", "bNoMusic")) = 1
    ClientSetup.bNoSound = Val(GetVar(ConfigPath, "CONFIG", "bNoSound")) = 1
    ClientSetup.bUseVideo = Val(GetVar(ConfigPath, "CONFIG", "bUseVideo")) = 1
    ClientSetup.byMemory = Val(GetVar(ConfigPath, "CONFIG", "byMemory"))

    Musica = Not ClientSetup.bNoMusic
    Sound = Not ClientSetup.bNoSound
    
    ClientSetupLoaded = True
    Call LimitarFPS(CByte(ClientSetup.bFPS))
End Sub
Public Sub SaveClientSetup()
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'
'**************************************************************
    If Not ClientSetupLoaded Then Exit Sub
    Dim ConfigPath As String
    ConfigPath = App.Path & "\..\Recursos\GameConfig.revival"
    With ClientSetup
        Call WriteVar(ConfigPath, "CONFIG", "bDinamic", IIf(.bDinamic, 1, 0))
        Call WriteVar(ConfigPath, "CONFIG", "bFPS", .bFPS)
        Call WriteVar(ConfigPath, "CONFIG", "bNoMusic", IIf(.bNoMusic, 1, 0))
        Call WriteVar(ConfigPath, "CONFIG", "bNoSound", IIf(.bNoSound, 1, 0))
        Call WriteVar(ConfigPath, "CONFIG", "bUseVideo", IIf(.bUseVideo, 1, 0))
        Call WriteVar(ConfigPath, "CONFIG", "byMemory", .byMemory)
    End With
End Sub

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
    Ciudades(1) = "Ullathorpe"
    Ciudades(2) = "Nix"
    Ciudades(3) = "Banderbill"

    CityDesc(1) = "Ullathorpe est� establecida en el medio de los grandes bosques de Argentum, es principalmente un pueblo de campesinos y le�adores. Su ubicaci�n hace de Ullathorpe un punto de paso obligado para todos los aventureros ya que se encuentra cerca de los lugares m�s legendarios de este mundo."
    CityDesc(2) = "Nix es una gran ciudad. Edificada sobre la costa oeste del principal continente de Argentum."
    CityDesc(3) = "Banderbill se encuentra al norte de Ullathorpe y Nix, es una de las ciudades m�s importantes de todo el imperio."

    ListaRazas(1) = "Humano"
    ListaRazas(2) = "Elfo"
    ListaRazas(3) = "Elfo Oscuro"
    ListaRazas(4) = "Gnomo"
    ListaRazas(5) = "Enano"

    ListaClases(1) = "Mago"
    ListaClases(2) = "Clerigo"
    ListaClases(3) = "Guerrero"
    ListaClases(4) = "Asesino"
    ListaClases(5) = "Bardo"
    ListaClases(6) = "Paladin"
    ListaClases(7) = "Cazador"

    SkillsNames(Skills.Suerte) = "Suerte"
    SkillsNames(Skills.Magia) = "Magia"
    SkillsNames(Skills.Robar) = "Robar"
    SkillsNames(Skills.Tacticas) = "Tacticas de combate"
    SkillsNames(Skills.Armas) = "Combate con armas"
    SkillsNames(Skills.Meditar) = "Meditar"
    SkillsNames(Skills.Apu�alar) = "Apu�alar"
    SkillsNames(Skills.Ocultarse) = "Ocultarse"
    SkillsNames(Skills.Supervivencia) = "Supervivencia"
    SkillsNames(Skills.Talar) = "Talar �rboles"
    SkillsNames(Skills.Comerciar) = "Comercio"
    SkillsNames(Skills.Defensa) = "Defensa con escudos"
    SkillsNames(Skills.Pesca) = "Pesca"
    SkillsNames(Skills.Mineria) = "Mineria"
    SkillsNames(Skills.Carpinteria) = "Carpinteria"
    SkillsNames(Skills.Herreria) = "Herreria"
    SkillsNames(Skills.Liderazgo) = "Liderazgo"
    SkillsNames(Skills.Domar) = "Domar animales"
    SkillsNames(Skills.Proyectiles) = "Armas de proyectiles"
    SkillsNames(Skills.Wresterling) = "Wresterling"
    SkillsNames(Skills.Navegacion) = "Navegacion"

    AtributosNames(1) = "Fuerza"
    AtributosNames(2) = "Agilidad"
    AtributosNames(3) = "Inteligencia"
    AtributosNames(4) = "Carisma"
    AtributosNames(5) = "Constitucion"
End Sub
