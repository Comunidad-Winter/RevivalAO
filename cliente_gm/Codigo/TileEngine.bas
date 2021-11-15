Attribute VB_Name = "Mod_TileEngine"



Option Explicit

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'    C       O       N       S      T
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'Map sizes in tiles
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef source As Any, ByVal byteCount As Long)

Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

Public Const GrhFogata = 1521

'bltbit constant
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source


'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'    T       I      P      O      S
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'Encabezado bmp
Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    X As Integer
    Y As Integer
End Type

'[MaTeO 11]
'Posicion en un mapa
Public Type Position2
    X As Single
    Y As Single
End Type
'[/MaTeO 11]

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh
'tamaño y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    NumFrames As Integer
    Frames(1 To 25) As Integer
    Speed As Integer
    Active As Boolean
    MiniMap_color As Long
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
        '[MaTeO 11]
    FrameCounter As Single
    '[/MaTeO 11]
   
    SpeedCounter As Byte
    Started As Byte
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(1 To 4) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(1 To 4) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(1 To 4) As Grh
    '[ANIM ATAK]
    WeaponAttack As Single
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(1 To 4) As Grh
    'ANIM ATAK HELIOS
     ShieldAttack As Single
     'ANIM ATAK HELIOS
End Type


'Lista de cuerpos
Public Type FxData
    Fx As Grh
    OffsetX As Long
    OffsetY As Long
End Type

'Apariencia del personaje
Public Type Char
    Active As Byte
    Heading As Byte ' As E_Heading ?
    Pos As Position
       '[MaTeO 9]
    Alas As BodyData
    '[/MaTeO 9]
    BodyNum As Integer
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    EscudoEqu As Boolean
    UsandoArma As Boolean
    Fx As Integer
    FxLoopTimes As Integer
    Criminal As Byte
    
    Nombre As String
    
    Moving As Byte
        
    '[MaTeO 11]
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    MoveOffset As Position2
    '[/MaTeO 11]
    
   
    ServerIndex As Integer
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    priv As Byte
    
End Type

'Info de un objeto
Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    charindex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    
    'ME Only
    Changed As Byte
End Type


Public MapPath As String


'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap As Integer 'Mapa actual
Public Userindex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public UserMaxAGU As Integer
Public UserMinAGU As Integer
Public UserMaxHAM As Integer
Public UserMinHAM As Integer

Public EngineRun As Boolean
Public FramesPerSec As Integer
Public FramesPerSecCounter As Long

'Tamaño del la vista en Tiles
Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

'Offset del desde 0,0 del main view
Public MainViewTop As Integer
Public MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

'Handle to where all the drawing is going to take place
Public DisplayFormhWnd As Long

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Totales?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public lastTime As Long 'Para controlar la velocidad


'[CODE]:MatuX'
Public MainDestRect   As RECT
'[END]'
Public MainViewRect   As RECT
Public BackBufferRect As RECT

Public MainViewWidth As Integer
Public MainViewHeight As Integer




'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As FxData
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public Grh() As Grh 'Animaciones publicas
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Usuarios?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'
'epa ;)
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿API?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'Blt
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?


'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'       [CODE 000]: MatuX
'
Public bRain        As Boolean 'está raineando?
Public bTecho       As Boolean 'hay techo?
Public brstTick     As Long

Private RLluvia(7)  As RECT  'RECT de la lluvia
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private LTLluvia(4) As Integer

Public charlist(1 To 10000) As Char

#If SeguridadAlkon Then

Public MI(1 To 1233) As clsManagerInvisibles
Public CualMI As Integer

#End If

'estados internos del surface (read only)
Public Enum TextureStatus
    tsOriginal = 0
    tsNight = 1
    tsFog = 2
End Enum

'[CODE 001]:MatuX
Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum
'[END]'
'
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

#If ConAlfaB Then

Private Declare Function BltAlphaFast Lib "vbabdx" (ByRef lpDDSDest As Any, ByRef lpDDSSource As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchSrc As Long, ByVal pitchDst As Long, ByVal dwMode As Long) As Long
Private Declare Function BltEfectoNoche Lib "vbabdx" (ByRef lpDDSDest As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchDst As Long, ByVal dwMode As Long) As Long
        Public Declare Function vbDABLalphablend16 Lib "vbDABL" (ByVal iMode As Integer, ByVal bColorKey As Integer, _
ByRef sPtr As Any, ByRef dPtr As Any, ByVal iAlphaVal As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer, _
ByVal isPitch As Integer, ByVal idPitch As Integer, ByVal iColorKey As Integer) As Integer
Public Declare Function vbDABLcolorblend16555 Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
Public Declare Function vbDABLcolorblend16565 Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
Public Declare Function vbDABLcolorblend16555ck Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
Public Declare Function vbDABLcolorblend16565ck Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long

#End If

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Sub CargarCabezas()
'On Error Resume Next
Dim n As Integer, I As Integer, Numheads As Integer, index As Integer

Dim Miscabezas() As tIndiceCabeza

Dim data() As Byte
Call modCompression.Get_File_Data(App.Path & "\..\Recursos\", "CABEZAS.IND", data, INIT_RESOURCE_FILE)
Open TemporalFile For Binary Access Write As #1
    Put #1, , data
Close #1

n = FreeFile
Open TemporalFile For Binary Access Read As #n

'cabecera
Get #n, , MiCabecera

'num de cabezas
Get #n, , Numheads

'ara prueba ya se actualizo

'Resize array
ReDim HeadData(0 To Numheads + 1) As HeadData
ReDim Miscabezas(0 To Numheads + 1) As tIndiceCabeza

For I = 1 To Numheads
    Get #n, , Miscabezas(I)
    InitGrh HeadData(I).Head(1), Miscabezas(I).Head(1), 0
    InitGrh HeadData(I).Head(2), Miscabezas(I).Head(2), 0
    InitGrh HeadData(I).Head(3), Miscabezas(I).Head(3), 0
    InitGrh HeadData(I).Head(4), Miscabezas(I).Head(4), 0
Next I

Close #n
Kill TemporalFile
End Sub

Sub CargarCascos()
On Error Resume Next
Dim n As Integer, I As Integer, NumCascos As Integer, index As Integer

Dim Miscabezas() As tIndiceCabeza
Dim data() As Byte
Call modCompression.Get_File_Data(App.Path & "\..\Recursos\", "CASCOS.IND", data, INIT_RESOURCE_FILE)
Open TemporalFile For Binary Access Write As #1
    Put #1, , data
Close #1

n = FreeFile
Open TemporalFile For Binary Access Read As #n

'cabecera
Get #n, , MiCabecera

'num de cabezas
Get #n, , NumCascos

'Resize array
ReDim CascoAnimData(0 To NumCascos + 1) As HeadData
ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza

For I = 1 To NumCascos
    Get #n, , Miscabezas(I)
    InitGrh CascoAnimData(I).Head(1), Miscabezas(I).Head(1), 0
    InitGrh CascoAnimData(I).Head(2), Miscabezas(I).Head(2), 0
    InitGrh CascoAnimData(I).Head(3), Miscabezas(I).Head(3), 0
    InitGrh CascoAnimData(I).Head(4), Miscabezas(I).Head(4), 0
Next I

Close #n
Kill TemporalFile
End Sub

Sub CargarCuerpos()
On Error Resume Next
Dim n As Integer, I As Integer
Dim NumCuerpos As Integer
Dim MisCuerpos() As tIndiceCuerpo

Dim data() As Byte
Call modCompression.Get_File_Data(App.Path & "\..\Recursos\", "PERSONAJES.IND", data, INIT_RESOURCE_FILE)
Open TemporalFile For Binary Access Write As #1
    Put #1, , data
Close #1

n = FreeFile
Open TemporalFile For Binary Access Read As #n

'cabecera
Get #n, , MiCabecera

'num de cabezas
Get #n, , NumCuerpos

'Resize array
ReDim BodyData(0 To NumCuerpos + 1) As BodyData
ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpo

For I = 1 To NumCuerpos
    Get #n, , MisCuerpos(I)
    InitGrh BodyData(I).Walk(1), MisCuerpos(I).Body(1), 0
    InitGrh BodyData(I).Walk(2), MisCuerpos(I).Body(2), 0
    InitGrh BodyData(I).Walk(3), MisCuerpos(I).Body(3), 0
    InitGrh BodyData(I).Walk(4), MisCuerpos(I).Body(4), 0
    BodyData(I).HeadOffset.X = MisCuerpos(I).HeadOffsetX
    BodyData(I).HeadOffset.Y = MisCuerpos(I).HeadOffsetY
Next I

Close #n
Kill TemporalFile
End Sub
Sub CargarFxs()
On Error Resume Next
Dim n As Integer, I As Integer
Dim NumFxs As Integer
Dim MisFxs() As tIndiceFx

Dim data() As Byte
Call modCompression.Get_File_Data(App.Path & "\..\Recursos\", "FXS.IND", data, INIT_RESOURCE_FILE)
Open TemporalFile For Binary Access Write As #1
    Put #1, , data
Close #1

n = FreeFile
Open TemporalFile For Binary Access Read As #n

'cabecera
Get #n, , MiCabecera

'num de cabezas
Get #n, , NumFxs

'Resize array
ReDim FxData(0 To NumFxs + 1) As FxData
ReDim MisFxs(0 To NumFxs + 1) As tIndiceFx

For I = 1 To NumFxs
    Get #n, , MisFxs(I)
    Call InitGrh(FxData(I).Fx, MisFxs(I).Animacion, 1)
    FxData(I).OffsetX = MisFxs(I).OffsetX
    FxData(I).OffsetY = MisFxs(I).OffsetY
Next I

Close #n
Kill TemporalFile
End Sub

Sub CargarTips()
On Error Resume Next
Dim n As Integer, I As Integer
Dim NumTips As Integer

Dim data() As Byte
Call modCompression.Get_File_Data(App.Path & "\..\Recursos\", "TIPS.AYU", data, INIT_RESOURCE_FILE)
Open TemporalFile For Binary Access Write As #1
    Put #1, , data
Close #1

n = FreeFile
Open TemporalFile For Binary Access Read As #n

'cabecera
Get #n, , MiCabecera

'num de cabezas
Get #n, , NumTips

'Resize array
ReDim Tips(1 To NumTips) As String * 255

For I = 1 To NumTips
    Get #n, , Tips(I)
Next I

Close #n
Kill TemporalFile
End Sub

Sub CargarArrayLluvia()
On Error Resume Next
Dim n As Integer, I As Integer
Dim Nu As Integer

Dim data() As Byte
Call modCompression.Get_File_Data(App.Path & "\..\Recursos\", "FK.IND", data, INIT_RESOURCE_FILE)
Open TemporalFile For Binary Access Write As #1
    Put #1, , data
Close #1

n = FreeFile
Open TemporalFile For Binary Access Read As #n

'cabecera
Get #n, , MiCabecera

'num de cabezas
Get #n, , Nu

'Resize array
ReDim bLluvia(1 To Nu) As Byte

For I = 1 To Nu
    Get #n, , bLluvia(I)
Next I

Close #n
Kill TemporalFile
End Sub
Sub ConvertCPtoTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal cx As Single, ByVal cy As Single, tX As Integer, tY As Integer)
'******************************************
'Converts where the user clicks in the main window
'to a tile position
'******************************************
Dim HWindowX As Integer
Dim HWindowY As Integer

cx = cx - StartPixelLeft
cy = cy - StartPixelTop

HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)

'Figure out X and Y tiles
cx = (cx \ TilePixelWidth)
cy = (cy \ TilePixelHeight)

If cx > HWindowX Then
    cx = (cx - HWindowX)

Else
    If cx < HWindowX Then
        cx = (0 - (HWindowX - cx))
    Else
        cx = 0
    End If
End If

If cy > HWindowY Then
    cy = (0 - (HWindowY - cy))
Else
    If cy < HWindowY Then
        cy = (cy - HWindowY)
    Else
        cy = 0
    End If
End If

tX = UserPos.X + cx
tY = UserPos.Y + cy

End Sub
'[MaTeO 9]
Sub MakeChar(ByVal charindex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer, ByVal Alas As Integer)
'[/MaTeO 9]

On Error Resume Next

'Apuntamos al ultimo Char
If charindex > LastChar Then LastChar = charindex

NumChars = NumChars + 1

If Arma = 0 Then Arma = 2
If Escudo = 0 Then Escudo = 2
If Casco = 0 Then Casco = 2
'anim Helios

   'anim Helios
charlist(charindex).iHead = Head
charlist(charindex).iBody = Body
charlist(charindex).Head = HeadData(Head)
charlist(charindex).Body = BodyData(Body)
charlist(charindex).Arma = WeaponAnimData(Arma)
'[ANIM ATAK]
charlist(charindex).Arma.WeaponAttack = 0
charlist(charindex).Escudo.ShieldAttack = 0 'Anim Atak Helios
charlist(charindex).Escudo = ShieldAnimData(Escudo)
charlist(charindex).Casco = CascoAnimData(Casco)
'[MaTeO 9]
charlist(charindex).Alas = BodyData(Alas)

'[/MaTeO 9]
charlist(charindex).Heading = Heading

'Reset moving stats
charlist(charindex).Moving = 0
charlist(charindex).MoveOffset.X = 0
charlist(charindex).MoveOffset.Y = 0

'Update position
charlist(charindex).Pos.X = X
charlist(charindex).Pos.Y = Y

'Make active
charlist(charindex).Active = 1

'Plot on map
MapData(X, Y).charindex = charindex

End Sub

Sub ResetCharInfo(ByVal charindex As Integer)

    charlist(charindex).Active = 0
    charlist(charindex).Criminal = 0
    charlist(charindex).Fx = 0
    charlist(charindex).FxLoopTimes = 0
    charlist(charindex).invisible = False

#If SeguridadAlkon Then
    Call MI(CualMI).ResetInvisible(charindex)
#End If

    charlist(charindex).Moving = 0
    charlist(charindex).muerto = False
    charlist(charindex).Nombre = ""
    charlist(charindex).pie = False
    charlist(charindex).Pos.X = 0
    charlist(charindex).Pos.Y = 0
    charlist(charindex).UsandoArma = False
charlist(charindex).EscudoEqu = False 'anim helios
End Sub


Sub EraseChar(ByVal charindex As Integer)
On Error Resume Next

'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************

charlist(charindex).Active = 0

'Update lastchar
If charindex = LastChar Then
    Do Until charlist(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If


MapData(charlist(charindex).Pos.X, charlist(charindex).Pos.Y).charindex = 0

Call ResetCharInfo(charindex)

'Update NumChars
NumChars = NumChars - 1

End Sub

Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************

Grh.GrhIndex = GrhIndex

If Started = 2 Then
    If GrhData(Grh.GrhIndex).NumFrames > 1 Then
        Grh.Started = 1
    Else
        Grh.Started = 0
    End If
Else
    Grh.Started = Started
End If

Grh.FrameCounter = 1
'[CODE 000]:MatuX
'
'  La linea generaba un error en la IDE, (no ocurría debido al
' on error)
'
'   Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
'
If Grh.GrhIndex <> 0 Then Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
'
'[END]'

End Sub

Sub MoveCharbyHead(ByVal charindex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
Dim addX As Integer
Dim addY As Integer
Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

X = charlist(charindex).Pos.X
Y = charlist(charindex).Pos.Y

'Figure out which way to move
Select Case nHeading

    Case E_Heading.NORTH
        addY = -1

    Case E_Heading.EAST
        addX = 1

    Case E_Heading.SOUTH
        addY = 1
    
    Case E_Heading.WEST
        addX = -1
        
End Select

nX = X + addX
nY = Y + addY

MapData(nX, nY).charindex = charindex
charlist(charindex).Pos.X = nX
charlist(charindex).Pos.Y = nY
MapData(X, Y).charindex = 0

charlist(charindex).MoveOffset.X = -1 * (TilePixelWidth * addX)
charlist(charindex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

charlist(charindex).Moving = 1
charlist(charindex).Heading = nHeading
'[MaTeO 11]
charlist(charindex).scrollDirectionX = addX
charlist(charindex).scrollDirectionY = addY
'[/MaTeO 11]
If UserEstado <> 1 Then Call DoPasosFx(charindex)

'areas viejos
If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    Debug.Print UserCharIndex
    Call EraseChar(charindex)
End If

End Sub

Public Sub DoFogataFx()
If Sound Then
    If bFogata Then
        bFogata = HayFogata()
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata()
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", LoopStyle.Enabled)
    End If
End If
End Sub

Function EstaPCarea(ByVal Index2 As Integer) As Boolean

Dim X As Integer, Y As Integer

For Y = UserPos.Y - MinYBorder + 1 To UserPos.Y + MinYBorder - 1
  For X = UserPos.X - MinXBorder + 1 To UserPos.X + MinXBorder - 1
            
            If MapData(X, Y).charindex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
  Next X
Next Y

EstaPCarea = False

End Function


Sub DoPasosFx(ByVal charindex As Integer)
Static pie As Boolean

If Not Sound Then Exit Sub

If Not UserNavegando Then
    If Not charlist(charindex).muerto And EstaPCarea(charindex) Then
        charlist(charindex).pie = Not charlist(charindex).pie
        If charlist(charindex).pie Then
            Call Audio.PlayWave(SND_PASOS1)
        Else
            Call Audio.PlayWave(SND_PASOS2)
        End If
    End If
Else
    Call Audio.PlayWave(SND_NAVEGANDO)
End If

End Sub


Sub MoveCharbyPos(ByVal charindex As Integer, ByVal nX As Integer, ByVal nY As Integer)

On Error Resume Next

Dim X As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nHeading As E_Heading



X = charlist(charindex).Pos.X
Y = charlist(charindex).Pos.Y

MapData(X, Y).charindex = 0

addX = nX - X
addY = nY - Y

If Sgn(addX) = 1 Then
    nHeading = E_Heading.EAST
End If

If Sgn(addX) = -1 Then
    nHeading = E_Heading.WEST
End If

If Sgn(addY) = -1 Then
    nHeading = E_Heading.NORTH
End If

If Sgn(addY) = 1 Then
    nHeading = E_Heading.SOUTH
End If

MapData(nX, nY).charindex = charindex


charlist(charindex).Pos.X = nX
charlist(charindex).Pos.Y = nY

charlist(charindex).MoveOffset.X = -1 * (TilePixelWidth * addX)
charlist(charindex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

charlist(charindex).Moving = 1
charlist(charindex).Heading = nHeading

'[MaTeO 11]
charlist(charindex).scrollDirectionX = Sgn(addX)
charlist(charindex).scrollDirectionY = Sgn(addY)
'[/MaTeO 11]

'parche para que no medite cuando camina
Dim fxCh As Integer
fxCh = charlist(charindex).Fx
If fxCh = FxMeditar.CHICO Or fxCh = FxMeditar.GRANDE Or fxCh = FxMeditar.MEDIANO Or fxCh = FxMeditar.XGRANDE Then
    charlist(charindex).Fx = 0
    charlist(charindex).FxLoopTimes = 0
End If

If Not EstaPCarea(charindex) Then Dialogos.QuitarDialogo (charindex)

If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    Call EraseChar(charindex)
End If

End Sub

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

'Check to see if its out of bounds
If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
    Exit Sub
Else
    'Start moving... MainLoop does the rest
    AddtoUserPos.X = X
    UserPos.X = tX
    AddtoUserPos.Y = Y
    UserPos.Y = tY
    UserMoving = 1
   
End If


    

End Sub


Function HayFogata() As Boolean
Dim j As Integer, k As Integer
For j = UserPos.X - 8 To UserPos.X + 8
    For k = UserPos.Y - 6 To UserPos.Y + 6
        If InMapBounds(j, k) Then
            If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    HayFogata = True
                    Exit Function
            End If
        End If
    Next k
Next j
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
Dim loopc As Integer
Dim Dale As Boolean

loopc = 1
Do While charlist(loopc).Active And Dale
    loopc = loopc + 1
    Dale = (loopc <= UBound(charlist))
Loop

NextOpenChar = loopc

End Function


Sub LoadGrhData()
'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim Grh As Integer
Dim Frame As Integer
Dim tempint As Integer


Dim data() As Byte
Call modCompression.Get_File_Data(App.Path & "\..\Recursos\", "GRAFICOS.IND", data, INIT_RESOURCE_FILE)
Open TemporalFile For Binary Access Write As #1
    Put #1, , data
Close #1

'Resize arrays
ReDim GrhData(1 To Config_Inicio.NumeroDeBMPs) As GrhData

Dim Pointer As Long

'Open files
Open TemporalFile For Binary Access Read As #1
Seek #1, 1

Get #1, , MiCabecera


Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint

'Fill Grh List

'Get first Grh Number
Get #1, , Grh

Do Until Grh <= 0
GrhData(Grh).Active = True
        
    'Get number of frames
    Get #1, , GrhData(Grh).NumFrames
    
    If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
    
    If GrhData(Grh).NumFrames > 1 Then
    
        'Read a animation GRH set
        For Frame = 1 To GrhData(Grh).NumFrames
         
            Get #1, , GrhData(Grh).Frames(Frame)
            If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > Config_Inicio.NumeroDeBMPs Then
                GoTo ErrorHandler
            End If
        
        Next Frame
    
        Get #1, , GrhData(Grh).Speed

        If GrhData(Grh).Speed <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
        If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
        If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
    
    Else
    
        'Read in normal GRH data
        Get #1, , GrhData(Grh).FileNum
        
        If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sX
        If GrhData(Grh).sX < 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sY
        If GrhData(Grh).sY < 0 Then GoTo ErrorHandler
            
        Get #1, , GrhData(Grh).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
        GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth
        
        GrhData(Grh).Frames(1) = Grh
            
    End If

    'Get Next Grh Number
    Get #1, , Grh


Loop
'************************************************

Close #1
Dim count As Long

ReDim data(0)
Call modCompression.Get_File_Data(App.Path & "\..\Recursos\", "MINIMAP.DAT", data, INIT_RESOURCE_FILE)
Open TemporalFile For Binary Access Write As #1
    Put #1, , data
Close #1

Open TemporalFile For Binary As #1
    Seek #1, 1
    For count = 1 To 32000
        If GrhData(count).Active Then
            Get #1, , GrhData(count).MiniMap_color
        End If
    Next count
Close #1

Kill TemporalFile
Exit Sub

ErrorHandler:
Close #1
If FileExist(TemporalFile, vbArchive) Then Kill TemporalFile
MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh & "-" & err.Description

End Sub

Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************

'Limites del mapa
If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    LegalPos = False
    Exit Function
End If

    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        LegalPos = False
        Exit Function
    End If
    
    '¿Hay un personaje?
If MapData(X, Y).charindex > 0 Then
LegalPos = False
Exit Function
End If
   
    If Not UserNavegando Then
        If HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    Else
        If Not HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    End If
    
LegalPos = True

End Function




Function InMapLegalBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps
'LEGAL/Walkable bounds
'*****************************************************************

If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************

If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function

Sub DDrawGrhtoSurface(Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte)

Dim CurrentGrh As Grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + (1 / (8 / Velocidad)) '[MaTeO]
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                End If
            End If
        End If
    End If
End If
'Figure out what frame to draw (always 1 if not animated)
CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
'Center Grh over X,Y pos
If center Then
    If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If
With SourceRect
        .Left = GrhData(CurrentGrh.GrhIndex).sX
        .Top = GrhData(CurrentGrh.GrhIndex).sY
        .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
End With
Surface.BltFast X, Y, SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum), SourceRect, DDBLTFAST_WAIT
End Sub

Sub DDrawTransGrhIndextoSurface(Surface As DirectDrawSurface7, Grh As Integer, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte)
Dim CurrentGrh As Grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

With destRect
    .Left = X
    .Top = Y
    .Right = .Left + GrhData(Grh).pixelWidth
    .Bottom = .Top + GrhData(Grh).pixelHeight
End With

Surface.GetSurfaceDesc SurfaceDesc

'Draw
If destRect.Left >= 0 And destRect.Top >= 0 And destRect.Right <= SurfaceDesc.lWidth And destRect.Bottom <= SurfaceDesc.lHeight Then
    With SourceRect
        .Left = GrhData(Grh).sX
        .Top = GrhData(Grh).sY
        .Right = .Left + GrhData(Grh).pixelWidth
        .Bottom = .Top + GrhData(Grh).pixelHeight
    End With
    
    Surface.BltFast destRect.Left, destRect.Top, SurfaceDB.Surface(GrhData(Grh).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

End Sub

'Sub DDrawTransGrhtoSurface(surface As DirectDrawSurface7, Grh As Grh, X As Integer, Y As Integer, Center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'[CODE 000]:MatuX
    Sub DDrawTransGrhtoSurface(Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'[END]'
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
'[CODE]:MatuX
'
'  CurrentGrh.GrhIndex = iGrhIndex
'
'[END]

'Dim CurrentGrh As Grh
Dim iGrhIndex As Integer
'Dim destRect As RECT
Dim SourceRect As RECT
'Dim SurfaceDesc As DDSURFACEDESC2
Dim QuitarAnimacion As Boolean



If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + (1 / (8 / Velocidad)) '[MaTeO]
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If charlist(KillAnim).FxLoopTimes <> LoopAdEternum Then
                            
                            If charlist(KillAnim).FxLoopTimes > 0 Then charlist(KillAnim).FxLoopTimes = charlist(KillAnim).FxLoopTimes - 1
                            If charlist(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                                charlist(KillAnim).Fx = 0
                                Exit Sub
                            End If
                            
                        End If
                    End If
               End If
            End If
        End If
    End If
End If

If Grh.GrhIndex = 0 Then Exit Sub

'Figure out what frame to draw (always 1 if not animated)
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
If iGrhIndex = 0 Then Exit Sub
'Center Grh over X,Y pos
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX
    .Top = GrhData(iGrhIndex).sY
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With


Surface.BltFast X, Y, SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

End Sub

#If ConAlfaB = 1 Then
    Sub DDrawTransGrhtoSurfaceAlpha(Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'[END]'
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
'[CODE]:MatuX
'
'  CurrentGrh.GrhIndex = iGrhIndex
'
'[END]

'Dim CurrentGrh As Grh
Dim iGrhIndex As Integer
'Dim destRect As RECT
Dim SourceRect As RECT
'Dim SurfaceDesc As DDSURFACEDESC2
Dim QuitarAnimacion As Boolean


If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + (1 / (8 / Velocidad)) '[MaTeO]
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If charlist(KillAnim).FxLoopTimes <> LoopAdEternum Then

                            If charlist(KillAnim).FxLoopTimes > 0 Then charlist(KillAnim).FxLoopTimes = charlist(KillAnim).FxLoopTimes - 1
                            If charlist(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                                charlist(KillAnim).Fx = 0
                                Exit Sub
                            End If

                        End If
                    End If
               End If
            End If
        End If
    End If
End If

If Grh.GrhIndex = 0 Then Exit Sub

'Figure out what frame to draw (always 1 if not animated)
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX + IIf(X < 0, Abs(X), 0)
    .Top = GrhData(iGrhIndex).sY + IIf(Y < 0, Abs(Y), 0)
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With

'surface.BltFast X, Y, SurfaceDB.surface(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

Dim Src As DirectDrawSurface7
Dim rDest As RECT
Dim dArray() As Byte, sArray() As Byte
Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
Dim Modo As Long

Set Src = SurfaceDB.Surface(GrhData(iGrhIndex).FileNum)

Src.GetSurfaceDesc ddsdSrc
Surface.GetSurfaceDesc ddsdDest

With rDest
    .Left = X
    .Top = Y
    .Right = X + GrhData(iGrhIndex).pixelWidth
    .Bottom = Y + GrhData(iGrhIndex).pixelHeight
    
    If .Right > ddsdDest.lWidth Then
        .Right = ddsdDest.lWidth
    End If
    If .Bottom > ddsdDest.lHeight Then
        .Bottom = ddsdDest.lHeight
    End If
End With

' 0 -> 16 bits 555
' 1 -> 16 bits 565
' 2 -> 16 bits raro (Sin implementar)
' 3 -> 24 bits
' 4 -> 32 bits

If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H3E0 Then
    Modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 1
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 3
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = 65280 And ddsdSrc.ddpfPixelFormat.lGBitMask = 65280 Then
    Modo = 4
Else
    'Modo = 2 '16 bits raro ?
    Surface.BltFast X, Y, Src, SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    Exit Sub
End If

Dim SrcLock As Boolean, DstLock As Boolean
SrcLock = False: DstLock = False

On Local Error GoTo HayErrorAlpha

Src.Lock SourceRect, ddsdSrc, DDLOCK_WAIT, 0
SrcLock = True
Surface.Lock rDest, ddsdDest, DDLOCK_WAIT, 0
DstLock = True

Surface.GetLockedArray dArray()
Src.GetLockedArray sArray()

Call BltAlphaFast(ByVal VarPtr(dArray(X + X, Y)), ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, Modo)

Surface.Unlock rDest
DstLock = False
Src.Unlock SourceRect
SrcLock = False


Exit Sub

HayErrorAlpha:
If SrcLock Then Src.Unlock SourceRect
If DstLock Then Surface.Unlock rDest

End Sub
#End If 'ConAlfaB = 1

Sub DrawBackBufferSurface()
    PrimarySurface.Blt MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT
End Sub

Function GetBitmapDimensions(BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
Dim BMHeader As BITMAPFILEHEADER
Dim BINFOHeader As BITMAPINFOHEADER

Open BmpFile For Binary Access Read As #1
Get #1, , BMHeader
Get #1, , BINFOHeader
Close #1
bmWidth = BINFOHeader.biWidth
bmHeight = BINFOHeader.biHeight
End Function

Sub DrawGrhtoHdc(hwnd As Long, hdc As Long, Grh As Integer, SourceRect As RECT, destRect As RECT)
    If Grh <= 0 Then Exit Sub
    
    SecundaryClipper.SetHWnd hwnd
    SurfaceDB.Surface(GrhData(Grh).FileNum).BltToDC hdc, SourceRect, destRect
End Sub
Sub RenderScreen(tilex As Integer, tiley As Integer, PixelOffsetX As Single, PixelOffsetY As Single)
On Error Resume Next

If UserCiego Then Exit Sub

Dim Y        As Integer 'Keeps track of where on map we are
Dim X        As Integer 'Keeps track of where on map we are
Dim minY     As Integer 'Start Y pos on current map
Dim maxY     As Integer 'End Y pos on current map
Dim minX     As Integer 'Start X pos on current map
Dim maxX     As Integer 'End X pos on current map
Dim ScreenX  As Integer 'Keeps track of where to place tile on screen
Dim ScreenY  As Integer 'Keeps track of where to place tile on screen
Dim Moved    As Byte
Dim Grh      As Grh     'Temp Grh for show tile and blocked
Dim TempChar As Char
Dim TextX    As Integer
Dim TextY    As Integer
Dim iPPx     As Integer 'Usado en el Layer de Chars
Dim iPPy     As Integer 'Usado en el Layer de Chars
Dim rSourceRect      As RECT    'Usado en el Layer 1
Dim iGrhIndex        As Integer 'Usado en el Layer 1
Dim PixelOffsetXTemp As Integer 'For centering grhs
Dim PixelOffsetYTemp As Integer 'For centering grhs
Dim nX As Integer
Dim nY As Integer
Dim ColorClan As Long

'Figure out Ends and Starts of screen
' Hardcodeado para speed!
minY = (tiley - (WindowTileHeight \ 2)) - TileBufferSize
maxY = (tiley + (WindowTileHeight \ 2)) + TileBufferSize
minX = (tilex - (WindowTileWidth \ 2)) - TileBufferSize
maxX = (tilex + (WindowTileWidth \ 2)) + TileBufferSize


'Draw floor layer
ScreenY = 8
For Y = (minY + 8) To maxY - 8
    ScreenX = 8
    For X = minX + 8 To maxX - 8
        If X <= 100 And X > 0 And Y <= 100 And Y >= 0 Then
            'Layer 1 **********************************
            With MapData(X, Y).Graphic(1)
                If (.Started = 1) Then
                    If (.SpeedCounter > 0) Then
                        .SpeedCounter = .SpeedCounter - 1
                        If (.SpeedCounter = 0) Then
                            .SpeedCounter = GrhData(.GrhIndex).Speed
                            .FrameCounter = .FrameCounter + (1 / (8 / Velocidad))
                            If (.FrameCounter > GrhData(.GrhIndex).NumFrames) Then _
                                .FrameCounter = 1
                        End If
                    End If
                End If
    
                'Figure out what frame to draw (always 1 if not animated)
                iGrhIndex = GrhData(.GrhIndex).Frames(.FrameCounter)
            End With
    
            rSourceRect.Left = GrhData(iGrhIndex).sX
            rSourceRect.Top = GrhData(iGrhIndex).sY
            rSourceRect.Right = rSourceRect.Left + GrhData(iGrhIndex).pixelWidth
            rSourceRect.Bottom = rSourceRect.Top + GrhData(iGrhIndex).pixelHeight
    
    
    
            'El width fue hardcodeado para speed!
            Call BackBufferSurface.BltFast( _
                    ((32 * ScreenX) - 32) + PixelOffsetX, _
                    ((32 * ScreenY) - 32) + PixelOffsetY, _
                    SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), _
                    rSourceRect, _
                    DDBLTFAST_WAIT)
            '******************************************
            'Layer 2 **********************************
            'MapData(X, Y).Blocked = False
            'If MapData(X, Y).Blocked Then Call DDrawTransGrhtoSurface(BackBufferSurface, charlist(UserCharIndex).Body.Walk(1), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 1, 1)
            
            If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        MapData(X, Y).Graphic(2), _
                        ((32 * ScreenX) - 32) + PixelOffsetX, _
                        ((32 * ScreenY) - 32) + PixelOffsetY, _
                        1, _
                        1)
            End If
            '******************************************
        End If
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
Next Y


'busco que nombre dibujar
Call ConvertCPtoTP(frmMain.MainViewShp.Left, frmMain.MainViewShp.Top, frmMain.MouseX, frmMain.MouseY, nX, nY)


'Draw Transparent Layers  (Layer 2, 3)
ScreenY = 8
For Y = minY + 8 To maxY - 1
    ScreenX = 5
    For X = minX + 5 To maxX - 5
        If X <= 100 And X > 0 And Y <= 100 And Y >= 0 Then
            
            iPPx = 32 * ScreenX - 32 + PixelOffsetX
            iPPy = 32 * ScreenY - 32 + PixelOffsetY
    
            'Object Layer **********************************
            If MapData(X, Y).ObjGrh.GrhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface( _
                            BackBufferSurface, _
                            MapData(X, Y).ObjGrh, _
                            iPPx, iPPy, 1, 1)
            End If
            '***********************************************
            'Char layer ************************************
            If MapData(X, Y).charindex <> 0 Then
                TempChar = charlist(MapData(X, Y).charindex)
                PixelOffsetXTemp = PixelOffsetX
                PixelOffsetYTemp = PixelOffsetY
    
                Moved = 0
                
                '[MaTeO 11]
               'If needed, move left and right
                If TempChar.scrollDirectionX <> 0 Then
                    TempChar.Body.Walk(TempChar.Heading).Started = 1
                    TempChar.Alas.Walk(TempChar.Heading).Started = 1
                    TempChar.Arma.WeaponWalk(TempChar.Heading).Started = 1
                    TempChar.Escudo.ShieldWalk(TempChar.Heading).Started = 1
                    PixelOffsetXTemp = PixelOffsetXTemp + TempChar.MoveOffset.X
                    TempChar.MoveOffset.X = TempChar.MoveOffset.X - (Velocidad * Sgn(TempChar.MoveOffset.X))
                    'tempChar.MoveOffset.X = tempChar.MoveOffset.X - (2 * Sgn(tempChar.MoveOffset.X))
                    Moved = 1
                   
                    'Check if we already got there
                    If (Sgn(TempChar.scrollDirectionX) = 1 And TempChar.MoveOffset.X >= 0) Or _
                            (Sgn(TempChar.scrollDirectionX) = -1 And TempChar.MoveOffset.X <= 0) Then
                        TempChar.MoveOffset.X = 0
                        TempChar.scrollDirectionX = 0
                    End If
                End If
     
                'If needed, move up and down
                If TempChar.scrollDirectionY <> 0 Then
                    TempChar.Body.Walk(TempChar.Heading).Started = 1
                    TempChar.Alas.Walk(TempChar.Heading).Started = 1
                    TempChar.Arma.WeaponWalk(TempChar.Heading).Started = 1
                    TempChar.Escudo.ShieldWalk(TempChar.Heading).Started = 1
                    PixelOffsetYTemp = PixelOffsetYTemp + TempChar.MoveOffset.Y
                    TempChar.MoveOffset.Y = TempChar.MoveOffset.Y - (Velocidad * Sgn(TempChar.MoveOffset.Y))
                    Moved = 1
                    If (Sgn(TempChar.scrollDirectionY) = 1 And TempChar.MoveOffset.Y >= 0) Or _
                            (Sgn(TempChar.scrollDirectionY) = -1 And TempChar.MoveOffset.Y <= 0) Then
                        TempChar.MoveOffset.Y = 0
                        TempChar.scrollDirectionY = 0
                    End If
                End If
                '[/MaTeO 11]
                'If done moving stop animation
                If Moved = 0 And TempChar.Moving = 1 Then
                    TempChar.Moving = 0
                    TempChar.Body.Walk(TempChar.Heading).FrameCounter = 1
                    TempChar.Body.Walk(TempChar.Heading).Started = 0
                    TempChar.Alas.Walk(TempChar.Heading).FrameCounter = 1
                    TempChar.Alas.Walk(TempChar.Heading).Started = 0
                    TempChar.Arma.WeaponWalk(TempChar.Heading).FrameCounter = 1
                    TempChar.Arma.WeaponWalk(TempChar.Heading).Started = 0
                    TempChar.Escudo.ShieldWalk(TempChar.Heading).FrameCounter = 1
                    TempChar.Escudo.ShieldWalk(TempChar.Heading).Started = 0
                End If
                
                '[ANIM ATAK]
                If TempChar.Arma.WeaponAttack > 0 Then
                    TempChar.Arma.WeaponAttack = TempChar.Arma.WeaponAttack - (1 / (8 / Velocidad))
                    If TempChar.Arma.WeaponAttack <= 0 Then
                        TempChar.Arma.WeaponWalk(TempChar.Heading).Started = 0
                            TempChar.Arma.WeaponWalk(TempChar.Heading).FrameCounter = 1
                    End If
                End If
                '[/ANIM ATAK]
                 '[ANIM ESCUDO]
                If TempChar.Escudo.ShieldAttack > 0 Then
                    TempChar.Escudo.ShieldAttack = TempChar.Escudo.ShieldAttack - (1 / (8 / Velocidad))
                    If TempChar.Escudo.ShieldAttack <= 0 Then
                        TempChar.Escudo.ShieldWalk(TempChar.Heading).Started = 0
                        TempChar.Escudo.ShieldWalk(TempChar.Heading).FrameCounter = 1
                    End If
                End If
                'Dibuja solamente players
                iPPx = ((32 * ScreenX) - 32) + PixelOffsetXTemp
                iPPy = ((32 * ScreenY) - 32) + PixelOffsetYTemp
                If TempChar.Head.Head(TempChar.Heading).GrhIndex <> 0 Then
                 
                 'If Not charlist(MapData(X, Y).charindex).invisible Or MismoClan(MapData(X, Y).charindex) = True Or UserCharIndex = MapData(X, Y).charindex Then
                 'ver invis con el gm
                 If Not charlist(MapData(X, Y).charindex).invisible Or charlist(MapData(X, Y).charindex).invisible Or MismoClan(MapData(X, Y).charindex) = True Or UserCharIndex = MapData(X, Y).charindex Then
    #If SeguridadAlkon Then
                        If Not MI(CualMI).IsInvisible(MapData(X, Y).charindex) Then
    #End If
                            '[MaTeO 9]
                            If TempChar.Heading = E_Heading.SOUTH Then
                                If TempChar.Alas.Walk(TempChar.Heading).GrhIndex <> 0 Then
                                    Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Alas.Walk(TempChar.Heading), iPPx + TempChar.Body.HeadOffset.X, iPPy + TempChar.Body.HeadOffset.Y + 25, 1, 1)
                                End If
                            End If
                            '[/MaTeO 9]
                            '[CUERPO]'
                                Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Body.Walk(TempChar.Heading), _
                                        (((32 * ScreenX) - 32) + PixelOffsetXTemp), _
                                        (((32 * ScreenY) - 32) + PixelOffsetYTemp), _
                                        1, 1)
                            '[CABEZA]'
                                Call DDrawTransGrhtoSurface( _
                                        BackBufferSurface, _
                                        TempChar.Head.Head(TempChar.Heading), _
                                        iPPx + TempChar.Body.HeadOffset.X, _
                                        iPPy + TempChar.Body.HeadOffset.Y, _
                                        1, 0)
                            '[Casco]'
                                If TempChar.Casco.Head(TempChar.Heading).GrhIndex <> 0 Then
                                    Call DDrawTransGrhtoSurface( _
                                            BackBufferSurface, _
                                            TempChar.Casco.Head(TempChar.Heading), _
                                            iPPx + TempChar.Body.HeadOffset.X, _
                                            iPPy + TempChar.Body.HeadOffset.Y, _
                                            1, 0)
                                End If
                                '[MaTeO 9]
                                If TempChar.Heading <> E_Heading.SOUTH Then
                                    If TempChar.Alas.Walk(TempChar.Heading).GrhIndex <> 0 Then
                                        Call DDrawTransGrhtoSurface( _
                                                BackBufferSurface, _
                                                TempChar.Alas.Walk(TempChar.Heading), _
                                                iPPx + TempChar.Body.HeadOffset.X, _
                                                iPPy + TempChar.Body.HeadOffset.Y + IIf(TempChar.Heading = E_Heading.NORTH, 25, 30), _
                                                1, 1) 'El primer 25, es cuando esta mirando para arriba, el siguiente 20 es cuando esta mirando para izquierda o derecha ¿Ta?, anda cambiando el "20"
                                    End If
                                End If
                                '[/MaTeO 9]
                            '[ARMA]'
                               Dim xx As Integer
                      If TempChar.Arma.WeaponWalk(TempChar.Heading).GrhIndex <> 0 Then
                         If TempChar.Body.HeadOffset.Y = -69 Then
                                xx = 31
                                ElseIf TempChar.Body.HeadOffset.Y = -94 Then
                                xx = 59
                                 ' mascotas
                                 ElseIf TempChar.Body.HeadOffset.Y = -78 Then
                                xx = 42
                                 ElseIf TempChar.Body.HeadOffset.Y = -75 Then
                                xx = 37
                                 ElseIf TempChar.Body.HeadOffset.Y = -55 Then
                                xx = 21
                                 ElseIf TempChar.Body.HeadOffset.Y = -83 Then
                                xx = 45
                                ElseIf TempChar.Body.HeadOffset.Y = -65 Then
                                xx = 27
                                ElseIf TempChar.Body.HeadOffset.Y = -60 Then
                                xx = 22
                                ElseIf TempChar.Body.HeadOffset.Y = -95 Then
                                xx = 60
                                ElseIf TempChar.Body.HeadOffset.Y = -48 Then
                                xx = 14
                                ElseIf TempChar.Body.HeadOffset.Y = -68 Then
                                xx = 30
                                ElseIf TempChar.Body.HeadOffset.Y = -120 Then
                                xx = 85
                                ' mascotas
                                ElseIf TempChar.Body.HeadOffset.Y = -72 Then
                                xx = 34
                                ElseIf TempChar.Body.HeadOffset.Y = -52 Then
                            xx = 18
                            ElseIf TempChar.Body.HeadOffset.Y = -80 Then
                            xx = 44
                            ElseIf TempChar.Body.HeadOffset.Y = -88 Then
                            xx = 52
                            ElseIf TempChar.Body.HeadOffset.Y = -90 Then
                            xx = 54
                            ElseIf TempChar.Body.HeadOffset.Y = -38 Then
                            xx = 4
                            ElseIf TempChar.Body.HeadOffset.Y = -50 Then
                            xx = 16
                            ElseIf TempChar.Body.HeadOffset.Y = -68 Then
                            xx = 30
                                Else
                                xx = 0
                               End If
                         Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Arma.WeaponWalk(TempChar.Heading), iPPx, iPPy - xx, 1, 1)
                      End If
                      '[Escudo]'
                    If TempChar.Escudo.ShieldWalk(TempChar.Heading).GrhIndex <> 0 Then
                        If TempChar.Body.HeadOffset.Y = -69 Then
                            xx = 31
                        ElseIf TempChar.Body.HeadOffset.Y = -94 Then
                            xx = 59
                        ElseIf TempChar.Body.HeadOffset.Y = -78 Then
                            xx = 40
                             ElseIf TempChar.Body.HeadOffset.Y = -75 Then
                                xx = 37
                                 ElseIf TempChar.Body.HeadOffset.Y = -55 Then
                                xx = 21
                                 ElseIf TempChar.Body.HeadOffset.Y = -83 Then
                                xx = 45
                                ElseIf TempChar.Body.HeadOffset.Y = -65 Then
                                xx = 27
                                ElseIf TempChar.Body.HeadOffset.Y = -60 Then
                                xx = 22
                                ElseIf TempChar.Body.HeadOffset.Y = -95 Then
                                xx = 60
                                ElseIf TempChar.Body.HeadOffset.Y = -48 Then
                                xx = 14
                                   ElseIf TempChar.Body.HeadOffset.Y = -120 Then
                                xx = 85
                                ElseIf TempChar.Body.HeadOffset.Y = -68 Then
                                xx = 30
                        ElseIf TempChar.Body.HeadOffset.Y = -72 Then
                            xx = 34
                         ElseIf TempChar.Body.HeadOffset.Y = -52 Then
                            xx = 18
                             ElseIf TempChar.Body.HeadOffset.Y = -80 Then
                            xx = 44
                            ElseIf TempChar.Body.HeadOffset.Y = -88 Then
                            xx = 52
                            ElseIf TempChar.Body.HeadOffset.Y = -90 Then
                            xx = 54
                            ElseIf TempChar.Body.HeadOffset.Y = -38 Then
                            xx = 4
                            ElseIf TempChar.Body.HeadOffset.Y = -50 Then
                            xx = 16
                            ElseIf TempChar.Body.HeadOffset.Y = -68 Then
                            xx = 30
                        Else
                            xx = 0
                        End If
                        If TempChar.EscudoEqu Then Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Escudo.ShieldWalk(TempChar.Heading), iPPx, iPPy - xx, 1, 1)
      
                                End If
                            '[Escudo]'
                                If TempChar.Escudo.ShieldWalk(TempChar.Heading).GrhIndex <> 0 Then
                                 If TempChar.Body.HeadOffset.Y = -78 Then
                            xx = 40
                             ElseIf TempChar.Body.HeadOffset.Y = -75 Then
                                xx = 37
                                 ElseIf TempChar.Body.HeadOffset.Y = -55 Then
                                xx = 21
                                 ElseIf TempChar.Body.HeadOffset.Y = -83 Then
                                xx = 45
                                ElseIf TempChar.Body.HeadOffset.Y = -65 Then
                                xx = 27
                                ElseIf TempChar.Body.HeadOffset.Y = -60 Then
                                xx = 22
                                ElseIf TempChar.Body.HeadOffset.Y = -95 Then
                                xx = 60
                                ElseIf TempChar.Body.HeadOffset.Y = -48 Then
                                xx = 14
                                   ElseIf TempChar.Body.HeadOffset.Y = -120 Then
                                xx = 85
                                ElseIf TempChar.Body.HeadOffset.Y = -68 Then
                                xx = 30
                                End If
                                    Call DDrawTransGrhtoSurface( _
                                            BackBufferSurface, _
                                            TempChar.Escudo.ShieldWalk(TempChar.Heading), _
                                            iPPx, iPPy - xx, 1, 1)
                                End If
                        
                        
                                 If Nombres Then
                                    'ya estoy dibujando SOLO si esta visible
                                    'If TempChar.invisible = False And Not MI(CualMI).IsInvisible(MapData(X, Y).CharIndex) Then
                                        If TempChar.Nombre <> "" Then
                                            'TempChar.nombre = "MaTeO <clan> [titulo]"
                                            Dim tName As String
                                            Dim tClan As String
                                            Dim tTitulo As String
                                            
                                            Dim CName As Long
                                            Dim cClan As Long
                                            Dim cTitulo As Long
                                                        CName = 0
                                            cClan = 0
                                            cTitulo = 0
                                            
                                            tName = vbNullString
                                            tClan = vbNullString
                                            tTitulo = vbNullString
                                        
                                            If InStr(1, TempChar.Nombre, "[") Then tTitulo = Mid(TempChar.Nombre, InStr(1, TempChar.Nombre, "["), InStr(1, TempChar.Nombre, "]") - InStr(1, TempChar.Nombre, "[") + 1)
                                            If InStr(1, TempChar.Nombre, "<") Then tClan = Mid(TempChar.Nombre, InStr(1, TempChar.Nombre, "<"), InStr(1, TempChar.Nombre, ">") - InStr(1, TempChar.Nombre, "<") + 1)
                                            
                                            tName = Left$(TempChar.Nombre, Len(TempChar.Nombre) - Len(tTitulo) - Len(tClan) - IIf(Right$(TempChar.Nombre, 1) = " ", 1, 0))
                                            Select Case TempChar.priv
                                                Case 0 'Ciudadano o Criminal
                                                    If TempChar.Criminal Then
                                                        cClan = RGB(255, 0, 0)
                                                    Else
                                                        cClan = RGB(0, 128, 255)
                                                    End If
                                                    
                                                    If charlist(MapData(X, Y).charindex).invisible Then
                                                        CName = RGB(255, 255, 0)
                                                    ElseIf TempChar.Criminal Then
                                                        CName = RGB(ColoresPJ(50).r, ColoresPJ(50).G, ColoresPJ(50).b)
                                                    Else
                                                        CName = RGB(ColoresPJ(49).r, ColoresPJ(49).G, ColoresPJ(49).b)
                                                    End If
                                                    
                                                Case 25 'Admin
                                                    cClan = RGB(255, 128, 64)
                                                    CName = cClan
                                                Case Else 'Los demas
                                                    CName = RGB(ColoresPJ(TempChar.priv).r, ColoresPJ(TempChar.priv).G, ColoresPJ(TempChar.priv).b)
                                                    cClan = RGB(255, 128, 64)
                                            End Select
                                            
                                            cTitulo = CName
                                            Call Dialogos.DrawText(iPPx - (frmMain.TextWidth(tName) / 2) + 16, iPPy + 30, tName, CName)
                                            If Len(tClan) Then Call Dialogos.DrawText(iPPx - (frmMain.TextWidth(tClan) / 2) + 16, iPPy + 40, tClan, cClan)
                                            If Len(tTitulo) Then Call Dialogos.DrawText(iPPx - (frmMain.TextWidth(tTitulo) / 2) + 16, iPPy + IIf(Len(tClan), 50, 40), tTitulo, cTitulo)
                                        End If
                                 End If
    #If SeguridadAlkon Then
                        Else
                            Do While True
                                Call MsgBox("WOAAAAA CHEATER!!! Ahora te deben estar matando de lo lindo ;)" & vbNewLine & "Aprieta OK para salir", vbCritical + vbOKOnly, ":D")
                                Call MsgBox("no, mejor no salimos")
                            Loop
                        End If  'end if not mi.isi
    #End If
                    End If  'end if ~in
    
                    If Dialogos.CantidadDialogos > 0 Then
                        Call Dialogos.Update_Dialog_Pos( _
                                (iPPx + TempChar.Body.HeadOffset.X), _
                                (iPPy + TempChar.Body.HeadOffset.Y), _
                                MapData(X, Y).charindex)
                    End If
                    
                    
                Else '<-> If TempChar.Head.Head(TempChar.Heading).GrhIndex <> 0 Then
                    If Dialogos.CantidadDialogos > 0 Then
                        Call Dialogos.Update_Dialog_Pos( _
                                (iPPx + TempChar.Body.HeadOffset.X), _
                                (iPPy + TempChar.Body.HeadOffset.Y), _
                                MapData(X, Y).charindex)
                    End If
    
                    Call DDrawTransGrhtoSurface( _
                            BackBufferSurface, _
                            TempChar.Body.Walk(TempChar.Heading), _
                            iPPx, iPPy, 1, 1)
                End If '<-> If TempChar.Head.Head(TempChar.Heading).GrhIndex <> 0 Then
    
    
                'Refresh charlist
                charlist(MapData(X, Y).charindex) = TempChar
    
                'BlitFX (TM)
                If charlist(MapData(X, Y).charindex).Fx <> 0 Then
    #If (ConAlfaB = 1) Then
                    Call DDrawTransGrhtoSurfaceAlpha( _
                            BackBufferSurface, _
                            FxData(TempChar.Fx).Fx, _
                            iPPx + FxData(TempChar.Fx).OffsetX, _
                            iPPy + FxData(TempChar.Fx).OffsetY, _
                            1, 1, MapData(X, Y).charindex)
    #Else
                    Call DDrawTransGrhtoSurface( _
                            BackBufferSurface, _
                            FxData(TempChar.Fx).Fx, _
                            iPPx + FxData(TempChar.Fx).OffsetX, _
                            iPPy + FxData(TempChar.Fx).OffsetY, _
                            1, 1, MapData(X, Y).charindex)
    #End If
                End If
            End If '<-> If MapData(X, Y).CharIndex <> 0 Then
            '*************************************************
            'Layer 3 *****************************************
            If MapData(X, Y).Graphic(3).GrhIndex <> 0 Then
                'Draw
                Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        MapData(X, Y).Graphic(3), _
                        ((32 * ScreenX) - 32) + PixelOffsetX, _
                        ((32 * ScreenY) - 32) + PixelOffsetY, _
                        1, 1)
            End If
            '************************************************
        End If
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
    If Y >= 100 Or Y < 1 Then Exit For
Next Y

If Not bTecho Then
    'Draw blocked tiles and grid
    ScreenY = 5
    For Y = minY + 5 To maxY - 1
        ScreenX = 5
        For X = minX + 5 To maxX
            'Check to see if in bounds
            If X < 101 And X > 0 And Y < 101 And Y > 0 Then
                If MapData(X, Y).Graphic(4).GrhIndex <> 0 Then
                    'Draw
                    Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        MapData(X, Y).Graphic(4), _
                        ((32 * ScreenX) - 32) + PixelOffsetX, _
                        ((32 * ScreenY) - 32) + PixelOffsetY, _
                        1, 1)
                End If
            End If
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
End If

If bLluvia(UserMap) = 1 Then
    If bRain Then
                'Figure out what frame to draw
                If llTick < DirectX.TickCount - 50 Then
                    iFrameIndex = iFrameIndex + 1
                    If iFrameIndex > 7 Then iFrameIndex = 0
                    llTick = DirectX.TickCount
                End If
    
                For Y = 0 To 4
                    For X = 0 To 4
                        Call BackBufferSurface.BltFast(LTLluvia(Y), LTLluvia(X), SurfaceDB.Surface(5556), RLluvia(iFrameIndex), DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT)
                    Next X
                Next Y
    End If
End If




Dim PP As RECT

PP.Left = 0
PP.Top = 0
PP.Right = WindowTileWidth * TilePixelWidth
PP.Bottom = WindowTileHeight * TilePixelHeight

End Sub
  





Public Function RenderSounds()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 4/22/2006
'Actualiza todos los sonidos del mapa.
'**************************************************************
    If bLluvia(UserMap) = 1 And Sound Then
        If bRain Then
            If bTecho Then
                If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                    If RainBufferIndex Then _
                        Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = Audio.PlayWave("lluviain.wav", LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviain
                End If
            Else
                If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                    If RainBufferIndex Then _
                        Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = Audio.PlayWave("lluviaout.wav", LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviaout
                End If
            End If
        End If
    End If
    
    DoFogataFx
End Function


Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer) As Boolean

If GrhIndex > 0 Then
        
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
        And charlist(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
        And charlist(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
        And charlist(UserCharIndex).Pos.Y <= Y
        
End If
End Function

Function PixelPos(ByVal X As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************
    PixelPos = (TilePixelWidth * X) - TilePixelWidth
End Function

Sub LoadGraphics()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero - complete rewrite
'Last Modify Date: 11/03/2006
'Initializes the SurfaceDB and sets up the rain rects
'**************************************************************
    'New surface manager :D
    Call SurfaceDB.Initialize(DirectDraw, ClientSetup.bUseVideo, App.Path & "\..\Recursos\", ClientSetup.byMemory)
          
    'Set up te rain rects
    RLluvia(0).Top = 0:      RLluvia(1).Top = 0:      RLluvia(2).Top = 0:      RLluvia(3).Top = 0
    RLluvia(0).Left = 0:     RLluvia(1).Left = 128:   RLluvia(2).Left = 256:   RLluvia(3).Left = 384
    RLluvia(0).Right = 128:  RLluvia(1).Right = 256:  RLluvia(2).Right = 384:  RLluvia(3).Right = 512
    RLluvia(0).Bottom = 128: RLluvia(1).Bottom = 128: RLluvia(2).Bottom = 128: RLluvia(3).Bottom = 128

    RLluvia(4).Top = 128:    RLluvia(5).Top = 128:    RLluvia(6).Top = 128:    RLluvia(7).Top = 128
    RLluvia(4).Left = 0:     RLluvia(5).Left = 128:   RLluvia(6).Left = 256:   RLluvia(7).Left = 384
    RLluvia(4).Right = 128:  RLluvia(5).Right = 256:  RLluvia(6).Right = 384:  RLluvia(7).Right = 512
    RLluvia(4).Bottom = 256: RLluvia(5).Bottom = 256: RLluvia(6).Bottom = 256: RLluvia(7).Bottom = 256
    
    'We are done!
    AddtoRichTextBox frmCargando.Status, "Hecho.", , , , 1, , False
End Sub

'[END]'
Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer) As Boolean
'*****************************************************************
'InitEngine
'*****************************************************************
Dim SurfaceDesc As DDSURFACEDESC2
Dim ddck As DDCOLORKEY


'Set intial user position
UserPos.X = MinXBorder
UserPos.Y = MinYBorder

'Fill startup variables

DisplayFormhWnd = setDisplayFormhWnd
MainViewTop = setMainViewTop
MainViewLeft = setMainViewLeft
TilePixelWidth = setTilePixelWidth
TilePixelHeight = setTilePixelHeight
WindowTileHeight = setWindowTileHeight
WindowTileWidth = setWindowTileWidth
TileBufferSize = setTileBufferSize

MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

MainViewWidth = (TilePixelWidth * WindowTileWidth)
MainViewHeight = (TilePixelHeight * WindowTileHeight)


ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock





DirectDraw.SetCooperativeLevel DisplayFormhWnd, DDSCL_NORMAL

'Primary Surface
' Fill the surface description structure
With SurfaceDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With



Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)

Set PrimaryClipper = DirectDraw.CreateClipper(0)
PrimaryClipper.SetHWnd frmMain.hwnd
PrimarySurface.SetClipper PrimaryClipper

Set SecundaryClipper = DirectDraw.CreateClipper(0)

With BackBufferRect
    .Left = 0
    .Top = 0
    .Right = TilePixelWidth * (WindowTileWidth + 2 * TileBufferSize)
    .Bottom = TilePixelHeight * (WindowTileHeight + 2 * TileBufferSize)
End With

With SurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    If ClientSetup.bUseVideo Then
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Else
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End If
    .lHeight = BackBufferRect.Bottom
    .lWidth = BackBufferRect.Right
End With

Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)

ddck.low = 0
ddck.high = 0
BackBufferSurface.SetColorKey DDCKEY_SRCBLT, ddck



Call LoadGrhData
Call CargarCuerpos
Call CargarCabezas
Call CargarCascos
Call CargarFxs


LTLluvia(0) = 224
LTLluvia(1) = 352
LTLluvia(2) = 480
LTLluvia(3) = 608
LTLluvia(4) = 736

AddtoRichTextBox frmCargando.Status, "Cargando Gráficos....", 0, 0, 0, , , True
Call LoadGraphics

InitTileEngine = True

End Function

Sub ShowNextFrame()
'***********************************************
'Updates and draws next frame to screen
'***********************************************

    '[MaTeO 11]
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
'[/MaTeO 11]
    '****** Set main view rectangle ******
    GetWindowRect DisplayFormhWnd, MainViewRect
    
    With MainViewRect
        .Left = .Left + MainViewLeft
        .Top = .Top + MainViewTop
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
    If EngineRun Then
        '****** Move screen Left and Right if needed ******
        If AddtoUserPos.X <> 0 Then
          OffsetCounterX = (OffsetCounterX - (Velocidad * Sgn(AddtoUserPos.X)))
            If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                OffsetCounterX = 0
              AddtoUserPos.X = 0
               UserMoving = 0
            End If
        '****** Move screen Up and Down if needed ******
        ElseIf AddtoUserPos.Y <> 0 Then
            OffsetCounterY = OffsetCounterY - (Velocidad * Sgn(AddtoUserPos.Y))
            If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                OffsetCounterY = 0
               AddtoUserPos.Y = 0
               UserMoving = 0
           End If
        End If
        
                

        '****** Update screen ******
        Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
        
        If IsSeguro = True Then
        Call Dialogos.DrawText(260, 275, "Seguro Activado", vbYellow)
        End If
        
        If IsSeguro = False Then
        Call Dialogos.DrawText(260, 275, "Seguro Desactivado", vbWhite)
        
        End If
        
        If IsSeguroC = True Then
        Call Dialogos.DrawText(260, 290, "Seguro Clan Activado", vbYellow)
        End If
        
        If IsSeguroC = False Then
        Call Dialogos.DrawText(260, 290, "Seguro Clan Desactivado", vbWhite)
        End If
        
        If CartelInvisibilidad Then Call Dialogos.DrawText(260, 305, "Invisibilidad: " & CartelInvisibilidad, vbCyan)
        Call Dialogos.MostrarTexto
        Call DibujarCartel
        
       
      
        Call DialogosClanes.Draw(Dialogos)
        Call Dialogos.DrawText(718, 260, "Hora: " & Time, vbWhite)
        Call Dialogos.DrawText(740, 275, "Fps: " & FramesPerSec, vbWhite)
        Call Dialogos.DrawText(680, 290, "Hay " & NumUsers & " Usuarios Online.", vbCyan)
        Call Dialogos.DrawText(260, 260, namemap & " (" & UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y & ")", vbWhite)

        If TiempoAsedio <> 0 And (UserMap = 114 Or UserMap = 115) Then Call Dialogos.DrawText(260, 655, "Faltan " & TiempoAsedio & " minutos para que finalize el Asedio.", vbCyan)
        Call DrawBackBufferSurface
        
        FramesPerSecCounter = FramesPerSecCounter + 1
    End If
End Sub

Sub CrearGrh(GrhIndex As Integer, index As Integer)
ReDim Preserve Grh(1 To index) As Grh
Grh(index).FrameCounter = 1
Grh(index).GrhIndex = GrhIndex
Grh(index).SpeedCounter = GrhData(GrhIndex).Speed
Grh(index).Started = 1
End Sub

Sub CargarAnimsExtra()
Call CrearGrh(6580, 1) 'Anim Invent
Call CrearGrh(534, 2) 'Animacion de teleport
Dim DDm As DDSURFACEDESC2
DDm.lHeight = 101
DDm.lWidth = 101
DDm.ddsCaps.lCaps = DDSCAPS_SYSTEMMEMORY
DDm.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
Set SupMiniMap = DirectDraw.CreateSurface(DDm)
Set SupBMiniMap = DirectDraw.CreateSurface(DDm)
End Sub

Function ControlVelocidad(ByVal lastTime As Long) As Boolean
ControlVelocidad = (GetTickCount - lastTime > 20)
End Function


#If ConAlfaB Then

Public Sub EfectoNoche(ByRef Surface As DirectDrawSurface7)
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
    Dim rRect As RECT
    
    Surface.GetSurfaceDesc ddsdDest
    
    With rRect
        .Left = 0
        .Top = 0
        .Right = ddsdDest.lWidth
        .Bottom = ddsdDest.lHeight
    End With
    
    If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
        Modo = 0
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
        Modo = 1
    Else
        Modo = 2
    End If
    
    Dim DstLock As Boolean
    DstLock = False
    
    On Local Error GoTo HayErrorAlpha
    
    Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
    DstLock = True
    
    Surface.GetLockedArray dArray()
    Call BltEfectoNoche(ByVal VarPtr(dArray(0, 0)), _
        ddsdDest.lWidth, ddsdDest.lHeight, ddsdDest.lPitch, _
        Modo)
    
HayErrorAlpha:
    If DstLock = True Then
        Surface.Unlock rRect
        DstLock = False
    End If
End Sub

#End If
Private Function MismoClan(ByVal Userindex As Integer) As Boolean
On Error Resume Next
MismoClan = False
If InStr(charlist(Userindex).Nombre, "<") > 0 And InStr(charlist(Userindex).Nombre, ">") > 0 Then
If UserClan = Mid(charlist(Userindex).Nombre, InStr(charlist(Userindex).Nombre, "<")) Then
MismoClan = True
End If
End If
End Function

Public Sub DibujarMiniMapa()
   
Dim map_x As Long, map_y As Long

    For map_y = 1 To 100
        For map_x = 1 To 100
            If MapData(map_x, map_y).Graphic(1).GrhIndex > 0 Then
                SetPixel frmMain.MiniMap.hdc, map_x, map_y, GrhData(MapData(map_x, map_y).Graphic(1).GrhIndex).MiniMap_color
            End If
        Next map_x
    Next map_y
   
    SetPixel frmMain.MiniMap.hdc, UserPos.X, UserPos.Y, RGB(255, 0, 0)
    SetPixel frmMain.MiniMap.hdc, UserPos.X + 1, UserPos.Y, RGB(255, 0, 0)
    SetPixel frmMain.MiniMap.hdc, UserPos.X - 1, UserPos.Y, RGB(255, 0, 0)
    SetPixel frmMain.MiniMap.hdc, UserPos.X, UserPos.Y - 1, RGB(255, 0, 0)
    SetPixel frmMain.MiniMap.hdc, UserPos.X, UserPos.Y + 1, RGB(255, 0, 0)

    frmMain.MiniMap.Refresh

End Sub
Public Sub GenerarMiniMapa()
Dim X As Integer
Dim Y As Integer
Dim I As Integer
Dim dr As RECT
Dim SR As RECT
SR.Left = 0
SR.Top = 0
SR.Bottom = 100
SR.Right = 100
SupBMiniMap.BltColorFill SR, vbBlack
For X = MinYBorder To MaxXBorder
For Y = MinYBorder To MaxYBorder
If MapData(X, Y).Graphic(1).GrhIndex > 0 Then
With MapData(X, Y).Graphic(1)
I = GrhData(.GrhIndex).Frames(1)
End With
SR.Left = GrhData(I).sX
SR.Top = GrhData(I).sY
SR.Bottom = GrhData(I).pixelWidth
SR.Right = GrhData(I).pixelHeight
dr.Left = X
dr.Top = Y
dr.Bottom = Y + 2
dr.Right = X + 2
SupBMiniMap.Blt dr, SurfaceDB.Surface(GrhData(I).FileNum), SR, DDBLT_DONOTWAIT
'SupMiniMap.BltFast x, y, SurfaceDB.GetBMP(GrhData(i).FileNum), Sr, DDBLTFAST_DESTCOLORKEY
End If
Next
Next
End Sub

'[MaTeO 11]
Public Sub LimitarFPS(ByRef index As Integer)
Select Case index
    Case 0 '18fps
        VelocidadLimiter = 56
        Velocidad = 8
    Case 1 '36fps
        VelocidadLimiter = 28
        Velocidad = 4
    Case 2 '72fps
        VelocidadLimiter = 14
        Velocidad = 2
    Case 3 '144fps
        VelocidadLimiter = 8
        Velocidad = 1
End Select
IndexSet = index
ClientSetup.bFPS = index
End Sub
'[/MaTeO 11]
