Attribute VB_Name = "Extra"

Option Explicit

Public Function EsNewbie(ByVal userindex As Integer) As Boolean
EsNewbie = UserList(userindex).Stats.ELV <= LimiteNewbie
End Function



Public Sub DoTileEvents(ByVal userindex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)

On Error GoTo errhandler

If MapData(mapainvo, mapainvoX1, mapainvoY1).userindex > 0 And MapData(mapainvo, mapainvoX2, mapainvoY2).userindex > 0 And MapData(mapainvo, mapainvoX3, mapainvoY3).userindex > 0 And MapData(mapainvo, mapainvoX4, mapainvoY4).userindex > 0 And MapInfo(mapainvo).criatinv = 0 Then
Call SendData(SendTarget.toall, 0, 0, "||Se ha invocado una criatura en la sala de invocaciones." & FONTTYPE_TALK)
Call SendData(SendTarget.toall, 0, 0, "TW107")
MapInfo(mapainvo).criatinv = 1
Dim criatura As Integer
Dim criatura2 As Integer
Dim criatura3 As Integer
Dim criatura4 As Integer
Dim criatura5 As Integer
Dim invoca As Integer
criatura = 919
criatura2 = 920
criatura3 = 921
criatura4 = 922
criatura5 = 923
invoca = RandomNumber(criatura, criatura5)
Call SpawnNpc(invoca, UserList(MapData(mapainvo, mapainvoX3, mapainvoY3).userindex).pos, True, False)
End If



Dim nPos As WorldPos
Dim FxFlag As Boolean

If InMapBounds(Map, x, y) Then
    
    If MapData(Map, x, y).OBJInfo.ObjIndex > 0 Then
        FxFlag = ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).OBJType = eOBJType.otTELEPORT
    End If
    
    If MapData(Map, x, y).TileExit.Map > 0 Then
    
    
'CHOTS | Solo Guerres y Kzas
    If MapData(Map, x, y).TileExit.Map = 69 Then
    If UCase(UserList(userindex).Clase) = "MAGO" Or UCase(UserList(userindex).Clase) = "BARDO" Or UCase(UserList(userindex).Clase) = "ASESINO" Or UCase(UserList(userindex).Clase) = "CLERIGO" Or UCase(UserList(userindex).Clase) = "PALADIN" Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Este mapa es exclusivo para Guerreros y Cazadores." & FONTTYPE_INFO)
    Call WarpUserChar(userindex, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.y + 1)
    Exit Sub
    End If
    End If
    
    If MapData(Map, x, y).TileExit.Map = 85 Then
    If Not UCase(UserList(userindex).Stats.ELV) = 55 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Necesitas ser lvl 55 para poder ingresar a la sala de invocaciones!." & FONTTYPE_INFO)
    Call WarpUserChar(userindex, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.y + 1)
    Exit Sub
    End If
    End If
'CHOTS | Solo Guerres y Kzas
    
        '¿Es mapa de newbies?
        If UCase$(MapInfo(MapData(Map, x, y).TileExit.Map).Restringir) = "SI" Then
            '¿El usuario es un newbie?
            If EsNewbie(userindex) Then
                If LegalPos(MapData(Map, x, y).TileExit.Map, MapData(Map, x, y).TileExit.x, MapData(Map, x, y).TileExit.y, PuedeAtravesarAgua(userindex)) Then
                    If FxFlag Then '¿FX?
                        Call WarpUserChar(userindex, MapData(Map, x, y).TileExit.Map, MapData(Map, x, y).TileExit.x, MapData(Map, x, y).TileExit.y, True)
                    Else
                        Call WarpUserChar(userindex, MapData(Map, x, y).TileExit.Map, MapData(Map, x, y).TileExit.x, MapData(Map, x, y).TileExit.y)
                    End If
                Else
                    Call ClosestLegalPos(MapData(Map, x, y).TileExit, nPos)
                    If nPos.x <> 0 And nPos.y <> 0 Then
                        If FxFlag Then
                            Call WarpUserChar(userindex, nPos.Map, nPos.x, nPos.y, True)
                        Else
                            Call WarpUserChar(userindex, nPos.Map, nPos.x, nPos.y)
                        End If
                    End If
                End If
            Else 'No es newbie
                Call SendData(SendTarget.toindex, userindex, 0, "||Mapa exclusivo para newbies." & FONTTYPE_INFO)
                Dim veces As Byte
                veces = 0
                Call ClosestStablePos(UserList(userindex).pos, nPos)

                If nPos.x <> 0 And nPos.y <> 0 Then
                        Call WarpUserChar(userindex, nPos.Map, nPos.x, nPos.y)
                End If
            End If
        Else 'No es un mapa de newbies
            If LegalPos(MapData(Map, x, y).TileExit.Map, MapData(Map, x, y).TileExit.x, MapData(Map, x, y).TileExit.y, PuedeAtravesarAgua(userindex)) Then
                If FxFlag Then
                    Call WarpUserChar(userindex, MapData(Map, x, y).TileExit.Map, MapData(Map, x, y).TileExit.x, MapData(Map, x, y).TileExit.y, True)
                Else
                    Call WarpUserChar(userindex, MapData(Map, x, y).TileExit.Map, MapData(Map, x, y).TileExit.x, MapData(Map, x, y).TileExit.y)
                End If
            Else
                Call ClosestLegalPos(MapData(Map, x, y).TileExit, nPos)
                If nPos.x <> 0 And nPos.y <> 0 Then
                    If FxFlag Then
                        Call WarpUserChar(userindex, nPos.Map, nPos.x, nPos.y, True)
                    Else
                        Call WarpUserChar(userindex, nPos.Map, nPos.x, nPos.y)
                    End If
                End If
            End If
        End If
    End If
    
End If

Exit Sub

errhandler:
    Call LogError("Error en DotileEvents")

End Sub

Function InRangoVision(ByVal userindex As Integer, x As Integer, y As Integer) As Boolean

If x > UserList(userindex).pos.x - MinXBorder And x < UserList(userindex).pos.x + MinXBorder Then
    If y > UserList(userindex).pos.y - MinYBorder And y < UserList(userindex).pos.y + MinYBorder Then
        InRangoVision = True
        Exit Function
    End If
End If
InRangoVision = False

End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, x As Integer, y As Integer) As Boolean

If x > Npclist(NpcIndex).pos.x - MinXBorder And x < Npclist(NpcIndex).pos.x + MinXBorder Then
    If y > Npclist(NpcIndex).pos.y - MinYBorder And y < Npclist(NpcIndex).pos.y + MinYBorder Then
        InRangoVisionNPC = True
        Exit Function
    End If
End If
InRangoVisionNPC = False

End Function


Function InMapBounds(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean

If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
    InMapBounds = False
Else
    InMapBounds = True
End If

End Function

Sub ClosestLegalPos(pos As WorldPos, ByRef nPos As WorldPos)
'*****************************************************************
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = pos.Map

Do While Not LegalPos(pos.Map, nPos.x, nPos.y)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = pos.y - LoopC To pos.y + LoopC
        For tX = pos.x - LoopC To pos.x + LoopC
            
            If LegalPos(nPos.Map, tX, tY) Then
                nPos.x = tX
                nPos.y = tY
                '¿Hay objeto?
                
                tX = pos.x + LoopC
                tY = pos.y + LoopC
  
            End If
        
        Next tX
    Next tY
    
    LoopC = LoopC + 1
    
Loop

If Notfound = True Then
    nPos.x = 0
    nPos.y = 0
End If

End Sub

Sub ClosestStablePos(pos As WorldPos, ByRef nPos As WorldPos)
'*****************************************************************
'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = pos.Map

Do While Not LegalPos(pos.Map, nPos.x, nPos.y)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = pos.y - LoopC To pos.y + LoopC
        For tX = pos.x - LoopC To pos.x + LoopC
            
            If LegalPos(nPos.Map, tX, tY) And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                nPos.x = tX
                nPos.y = tY
                '¿Hay objeto?
                
                tX = pos.x + LoopC
                tY = pos.y + LoopC
  
            End If
        
        Next tX
    Next tY
    
    LoopC = LoopC + 1
    
Loop

If Notfound = True Then
    nPos.x = 0
    nPos.y = 0
End If

End Sub

Function NameIndex(ByRef name As String) As Integer

Dim userindex As Integer
'¿Nombre valido?
If name = "" Then
    NameIndex = 0
    Exit Function
End If

name = UCase$(Replace(name, "+", " "))

userindex = 1
Do Until UCase$(UserList(userindex).name) = name
    
    userindex = userindex + 1
    
    If userindex > MaxUsers Then
        NameIndex = 0
        Exit Function
    End If
    
Loop
 
NameIndex = userindex
 
End Function



Function IP_Index(ByVal inIP As String) As Integer
 
Dim userindex As Integer
'¿Nombre valido?
If inIP = "" Then
    IP_Index = 0
    Exit Function
End If
  
userindex = 1
Do Until UserList(userindex).ip = inIP
    
    userindex = userindex + 1
    
    If userindex > MaxUsers Then
        IP_Index = 0
        Exit Function
    End If
    
Loop
 
IP_Index = userindex

Exit Function

End Function


Function CheckForSameIP(ByVal userindex As Integer, ByVal UserIP As String) As Boolean
Dim LoopC As Integer
For LoopC = 1 To MaxUsers
    If UserList(LoopC).flags.UserLogged = True Then
        If UserList(LoopC).ip = UserIP And userindex <> LoopC Then
            CheckForSameIP = True
            Exit Function
        End If
    End If
Next LoopC
CheckForSameIP = False
End Function

Function CheckForSameName(ByVal userindex As Integer, ByVal name As String) As Boolean
'Controlo que no existan usuarios con el mismo nombre
Dim LoopC As Long
For LoopC = 1 To MaxUsers
    If UserList(LoopC).flags.UserLogged Then
        
        'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
        'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
        'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
        'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ITEMS
        'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
        
        If UCase$(UserList(LoopC).name) = UCase$(name) Then
            CheckForSameName = True
            Exit Function
        End If
    End If
Next LoopC
CheckForSameName = False
End Function

Sub HeadtoPos(ByVal Head As eHeading, ByRef pos As WorldPos)
'*****************************************************************
'Toma una posicion y se mueve hacia donde esta perfilado
'*****************************************************************
Dim x As Integer
Dim y As Integer
Dim tempVar As Single
Dim nX As Integer
Dim nY As Integer

x = pos.x
y = pos.y

If Head = eHeading.NORTH Then
    nX = x
    nY = y - 1
End If

If Head = eHeading.SOUTH Then
    nX = x
    nY = y + 1
End If

If Head = eHeading.EAST Then
    nX = x + 1
    nY = y
End If

If Head = eHeading.WEST Then
    nX = x - 1
    nY = y
End If

'Devuelve valores
pos.x = nX
pos.y = nY

End Sub

Function LegalPos(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, Optional ByVal PuedeAgua As Boolean = False) As Boolean

'¿Es un mapa valido?
If (Map <= 0 Or Map > NumMaps) Or _
   (x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder) Then
            LegalPos = False
Else
  
If Not PuedeAgua Then
LegalPos = (MapData(Map, x, y).Blocked <> 1) And _
(MapData(Map, x, y).userindex = 0) And _
(MapData(Map, x, y).NpcIndex = 0) And _
(Not HayAgua(Map, x, y))
Else
LegalPos = (MapData(Map, x, y).Blocked <> 1) And _
(MapData(Map, x, y).userindex = 0) And _
(MapData(Map, x, y).NpcIndex = 0) And _
(HayAgua(Map, x, y))
End If
   
End If

End Function

Function LegalPosNPC(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal AguaValida As Byte) As Boolean

If (Map <= 0 Or Map > NumMaps) Or _
   (x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder) Then
    LegalPosNPC = False
Else

 If AguaValida = 0 Then
   LegalPosNPC = (MapData(Map, x, y).Blocked <> 1) And _
     (MapData(Map, x, y).userindex = 0) And _
     (MapData(Map, x, y).NpcIndex = 0) And _
     (MapData(Map, x, y).trigger <> eTrigger.POSINVALIDA) _
     And Not HayAgua(Map, x, y)
 Else
   LegalPosNPC = (MapData(Map, x, y).Blocked <> 1) And _
     (MapData(Map, x, y).userindex = 0) And _
     (MapData(Map, x, y).NpcIndex = 0) And _
     (MapData(Map, x, y).trigger <> eTrigger.POSINVALIDA)
 End If
 
End If


End Function

Sub SendHelp(ByVal Index As Integer)
Dim NumHelpLines As Integer
Dim LoopC As Integer

NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

For LoopC = 1 To NumHelpLines
    Call SendData(SendTarget.toindex, Index, 0, "||" & GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC) & FONTTYPE_INFO)
Next LoopC

End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal userindex As Integer)
    If Npclist(NpcIndex).NroExpresiones > 0 Then
        Dim randomi
        randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & Npclist(NpcIndex).Expresiones(randomi) & "°" & Npclist(NpcIndex).char.CharIndex & FONTTYPE_INFO)
    End If
End Sub
Sub LookatTile_AutoAim(ByVal userindex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)

Dim myX As Integer, myY As Integer
Dim Area As Integer

Call LookatTile(userindex, Map, x, y)
If UserList(userindex).flags.TargetUser <> 0 Or UserList(userindex).flags.TargetNPC <> 0 Then Exit Sub

For Area = 1 To 3
    For myX = (x - Area) To (x + Area)
    For myY = (y - Area) To (y + Area)
        Call LookatTile(userindex, Map, myX, myY)
        If (UserList(userindex).flags.TargetUser <> 0 Or UserList(userindex).flags.TargetNPC <> 0) And UserList(userindex).flags.TargetUser <> userindex Then Exit Sub
    
    Next myY
    Next myX
Next Area
Call LookatTile(userindex, Map, x, y)
End Sub

Sub LookatTile(ByVal userindex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)
On Error GoTo errhandler
'Responde al click del usuario sobre el mapa
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Stat As String
Dim OBJType As Integer
Dim Colorvesa As String

'¿Posicion valida?
If InMapBounds(Map, x, y) Then
    UserList(userindex).flags.TargetMap = Map
    UserList(userindex).flags.TargetX = x
    UserList(userindex).flags.TargetY = y
    '¿Es un obj?
    If MapData(Map, x, y).OBJInfo.ObjIndex > 0 Then
        'Informa el nombre
        UserList(userindex).flags.TargetObjMap = Map
        UserList(userindex).flags.TargetObjX = x
        UserList(userindex).flags.TargetObjY = y
        FoundSomething = 1
    ElseIf MapData(Map, x + 1, y).OBJInfo.ObjIndex > 0 Then
        'Informa el nombre
        If ObjData(MapData(Map, x + 1, y).OBJInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            UserList(userindex).flags.TargetObjMap = Map
            UserList(userindex).flags.TargetObjX = x + 1
            UserList(userindex).flags.TargetObjY = y
            FoundSomething = 1
        End If
    ElseIf MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            'Informa el nombre
            UserList(userindex).flags.TargetObjMap = Map
            UserList(userindex).flags.TargetObjX = x + 1
            UserList(userindex).flags.TargetObjY = y + 1
            FoundSomething = 1
        End If
    ElseIf MapData(Map, x, y + 1).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(Map, x, y + 1).OBJInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            'Informa el nombre
            UserList(userindex).flags.TargetObjMap = Map
            UserList(userindex).flags.TargetObjX = x
            UserList(userindex).flags.TargetObjY = y + 1
            FoundSomething = 1
        End If
    End If
    
    If FoundSomething = 1 Then
        UserList(userindex).flags.TargetObj = MapData(Map, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.ObjIndex
        If MostrarCantidad(UserList(userindex).flags.TargetObj) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||" & ObjData(UserList(userindex).flags.TargetObj).name & " - " & MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.Amount & "" & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||" & ObjData(UserList(userindex).flags.TargetObj).name & FONTTYPE_INFO)
        End If
    
    End If
    '¿Es un personaje?
    If y + 1 <= YMaxMapSize Then
        If MapData(Map, x, y + 1).userindex > 0 Then
            TempCharIndex = MapData(Map, x, y + 1).userindex
            If UserList(TempCharIndex).showName Then    ' Es GM y pidió que se oculte su nombre??
                FoundChar = 1
            End If
        End If
        If MapData(Map, x, y + 1).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, x, y + 1).NpcIndex
            FoundChar = 2
        End If
    End If
    '¿Es un personaje?
    If FoundChar = 0 Then
        If MapData(Map, x, y).userindex > 0 Then
            TempCharIndex = MapData(Map, x, y).userindex
            If UserList(TempCharIndex).showName Then    ' Es GM y pidió que se oculte su nombre??
                FoundChar = 1
            End If
        End If
        If MapData(Map, x, y).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, x, y).NpcIndex
            FoundChar = 2
        End If
    End If
    
    
    'Reaccion al personaje
    If FoundChar = 1 Then '  ¿Encontro un Usuario?
            
       If UserList(TempCharIndex).flags.AdminInvisible = 0 Or UserList(userindex).flags.Privilegios = PlayerType.Dios Then
            
            If UserList(TempCharIndex).DescRM = "" Then
                If EsNewbie(TempCharIndex) Then
                    Stat = " <NEWBIE>"
                End If
                
                If UserList(TempCharIndex).Faccion.ArmadaReal = 1 Then
                    Stat = Stat & " <Armada Real> " & "<" & TituloReal(TempCharIndex) & ">"
                ElseIf UserList(TempCharIndex).Faccion.FuerzasCaos = 1 Then
                    Stat = Stat & " <Fuerzas del Caos> " & "<" & TituloCaos(TempCharIndex) & ">"
                End If
                
                If UserList(TempCharIndex).GuildIndex > 0 Then
                    Stat = Stat & " <" & Guilds(UserList(TempCharIndex).GuildIndex).GuildName & ">"
                End If
                
                If Len(UserList(TempCharIndex).Desc) > 1 Then
                    Stat = "Ves a " & UserList(TempCharIndex).name & Stat & " " & UserList(TempCharIndex).Desc
                Else
                    Stat = "Ves a " & UserList(TempCharIndex).name & Stat
                End If
                
                If UserList(TempCharIndex).flags.PertAlCons > 0 Then
                If UserList(TempCharIndex).Stats.UsuariosMatados < 100 Then
                    Stat = Stat & " <Rey Imperial> -" & " <UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Aprendiz)> " & "<" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & FONTTYPE_CONSEJOVesA
                End If
                If UserList(TempCharIndex).Stats.UsuariosMatados < 200 Then
                    Stat = Stat & " <Rey Imperial> -" & " <UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Soldado)> " & "<" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & FONTTYPE_CONSEJOVesA
                End If
                  If UserList(TempCharIndex).Stats.UsuariosMatados < 300 Then
                    Stat = Stat & " <Rey Imperial> -" & " <UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Teniente)> " & "<" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & FONTTYPE_CONSEJOVesA
                End If
                  If UserList(TempCharIndex).Stats.UsuariosMatados < 500 Then
                    Stat = Stat & " <Rey Imperial> -" & " <UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Capitan)> " & "<" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & FONTTYPE_CONSEJOVesA
                End If
                  If UserList(TempCharIndex).Stats.UsuariosMatados < 1000 Then
                    Stat = Stat & " <Rey Imperial> -" & " <UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Comandante)> " & "<" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & FONTTYPE_CONSEJOVesA
                End If
                 If UserList(TempCharIndex).Stats.UsuariosMatados >= 1000 Then
                    Stat = Stat & " <Rey Imperial> -" & " <UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Heroe RevivalAo)> " & "<" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & FONTTYPE_CONSEJOVesA
                End If
                ElseIf UserList(TempCharIndex).flags.PertAlConsCaos > 0 Then
                If UserList(TempCharIndex).Stats.UsuariosMatados < 100 Then
                    Stat = Stat & " <Rey Del Caos> -" & " <UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Aprendiz)>" & " <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & FONTTYPE_WETAS
                End If
                If UserList(TempCharIndex).Stats.UsuariosMatados < 200 Then
                    Stat = Stat & " <Rey Del Caos> -" & " <UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Asesino)>" & " <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & FONTTYPE_WETAS
                End If
                If UserList(TempCharIndex).Stats.UsuariosMatados < 300 Then
                    Stat = Stat & " <Rey Del Caos> -" & " <UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Mutilador)>" & " <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & FONTTYPE_WETAS
                End If
                If UserList(TempCharIndex).Stats.UsuariosMatados < 500 Then
                    Stat = Stat & " <Rey Del Caos> -" & " <UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Descuartizador)>" & " <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & FONTTYPE_WETAS
                End If
                If UserList(TempCharIndex).Stats.UsuariosMatados < 1000 Then
                    Stat = Stat & " <Rey Del Caos> -" & " <UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Destripador)>" & " <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & FONTTYPE_WETAS
                End If
                If UserList(TempCharIndex).Stats.UsuariosMatados >= 1000 Then
                    Stat = Stat & " <Rey Del Caos> -" & " <UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Dark RevivalAo)>" & " <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & FONTTYPE_WETAS
                End If
                Else
                   If UserList(TempCharIndex).flags.Privilegios > 0 Then
                   
                     If UserList(TempCharIndex).flags.Privilegios = 3 Then
                            Stat = Stat & " - <GM> - <Dios> - <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & FONTTYPE_CONSEJOCAOSVesA
                ElseIf UserList(TempCharIndex).flags.Privilegios = 2 Then
                            Stat = Stat & " - <GM> - <SemiDios> - <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & FONTTYPE_CONSEJOCAOSVesA
                ElseIf UserList(TempCharIndex).flags.Privilegios = 1 Then
                            Stat = Stat & " - <GM> - <Consejero> - <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & FONTTYPE_CONSEJOCAOSVesA
         End If

ElseIf Criminal(TempCharIndex) Then
If UserList(TempCharIndex).Stats.UsuariosMatados < 100 Then
Stat = Stat & " <CRIMINAL> " & "<UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Aprendiz)>" & " <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & "~255~0~0~1~0"
End If
If UserList(TempCharIndex).Stats.UsuariosMatados < 200 Then
Stat = Stat & " <CRIMINAL> " & "<UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Asesino)>" & " <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & "~255~0~0~1~0"
End If
If UserList(TempCharIndex).Stats.UsuariosMatados < 300 Then
Stat = Stat & " <CRIMINAL> " & "<UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Mutilador)>" & " <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & "~255~0~0~1~0"
End If
If UserList(TempCharIndex).Stats.UsuariosMatados < 500 Then
Stat = Stat & " <CRIMINAL> " & "<UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Descuartizador)>" & " <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & "~255~0~0~1~0"
End If
If UserList(TempCharIndex).Stats.UsuariosMatados < 1000 Then
Stat = Stat & " <CRIMINAL> " & "<UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Destripador)>" & " <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & "~255~0~0~1~0"
End If
If UserList(TempCharIndex).Stats.UsuariosMatados >= 1000 Then
Stat = Stat & " <CRIMINAL> " & "<UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Dark Revival)>" & " <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & "~255~0~0~1~0"
End If
Else
If UserList(TempCharIndex).Stats.UsuariosMatados < 100 Then
Stat = Stat & " <CIUDADANO> " & "<UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Aprendiz)>" & " <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & "~0~0~200~1~0"
End If
If UserList(TempCharIndex).Stats.UsuariosMatados < 200 Then
Stat = Stat & " <CIUDADANO> " & "<UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Soldado)>" & " <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & "~0~0~200~1~0"
End If
If UserList(TempCharIndex).Stats.UsuariosMatados < 300 Then
Stat = Stat & " <CIUDADANO> " & "<UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Teniente)>" & " <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & "~0~0~200~1~0"
End If
If UserList(TempCharIndex).Stats.UsuariosMatados < 500 Then
Stat = Stat & " <CIUDADANO> " & "<UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Capitan)>" & " <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & "~0~0~200~1~0"
End If
If UserList(TempCharIndex).Stats.UsuariosMatados < 1000 Then
Stat = Stat & " <CIUDADANO> " & "<UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Comandante)>" & " <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & "~0~0~200~1~0"
End If
If UserList(TempCharIndex).Stats.UsuariosMatados >= 1000 Then
Stat = Stat & " <CIUDADANO> " & "<UserDies: " & UserList(TempCharIndex).Stats.UsuariosMatados & " (Heroe RevivalAo)>" & " <" & UserList(TempCharIndex).Clase & " " & UserList(TempCharIndex).Raza & ">" & "~0~0~200~1~0"
End If
End If
            End If

            Else
                Stat = UserList(TempCharIndex).DescRM & " " & FONTTYPE_INFOBOLD
            End If
            
            

            
            If Len(Stat) > 0 Then _
                Call SendData(SendTarget.toindex, userindex, 0, "||" & Stat)

            FoundSomething = 1
            UserList(userindex).flags.TargetUser = TempCharIndex
            UserList(userindex).flags.TargetNPC = 0
            UserList(userindex).flags.TargetNpcTipo = eNPCType.Comun
       End If

    End If
    If FoundChar = 2 Then '¿Encontro un NPC?
            Dim estatus As String
            
            If UserList(userindex).flags.Privilegios >= PlayerType.SemiDios Then
                estatus = "(" & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & ")"
            Else
                If UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) >= 0 And UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) <= 10 Then
                    estatus = "(Dudoso) "
                ElseIf UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) > 10 And UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) <= 20 Then
                    If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP / 2) Then
                        estatus = "(Herido) "
                    Else
                        estatus = "(Sano) "
                    End If
                ElseIf UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) > 20 And UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) <= 30 Then
                    If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.5) Then
                        estatus = "(Malherido) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.75) Then
                        estatus = "(Herido) "
                    Else
                        estatus = "(Sano) "
                    End If
                ElseIf UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) > 30 And UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) <= 40 Then
                    If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.25) Then
                        estatus = "(Muy malherido) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.5) Then
                        estatus = "(Herido) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.75) Then
                        estatus = "(Levemente herido) "
                    Else
                        estatus = "(Sano) "
                    End If
                ElseIf UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) > 40 And UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) < 60 Then
                    If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.05) Then
                        estatus = "(Agonizando) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.1) Then
                        estatus = "(Casi muerto) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.25) Then
                        estatus = "(Muy Malherido) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.5) Then
                        estatus = "(Herido) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.75) Then
                        estatus = "(Levemente herido) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP) Then
                        estatus = "(Sano) "
                    Else
                        estatus = "(Intacto) "
                    End If
                ElseIf UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) >= 60 Then
                    estatus = "(" & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & ") "
                Else
                    estatus = "!error!"
                End If
            End If
            
            If Len(Npclist(TempCharIndex).Desc) > 1 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & Npclist(TempCharIndex).Desc & "°" & Npclist(TempCharIndex).char.CharIndex & FONTTYPE_INFO)
  
            Else
                If Npclist(TempCharIndex).MaestroUser > 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "|| " & estatus & Npclist(TempCharIndex).name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).name & FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "|| " & estatus & Npclist(TempCharIndex).name & "." & FONTTYPE_INFO)
                End If
                
            End If
            FoundSomething = 1
            UserList(userindex).flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
            UserList(userindex).flags.TargetNPC = TempCharIndex
            UserList(userindex).flags.TargetUser = 0
            UserList(userindex).flags.TargetObj = 0
        
    End If
    
    If FoundChar = 0 Then
        UserList(userindex).flags.TargetNPC = 0
        UserList(userindex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(userindex).flags.TargetUser = 0
    End If
    
    '*** NO ENCOTRO NADA ***
    If FoundSomething = 0 Then
        UserList(userindex).flags.TargetNPC = 0
        UserList(userindex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(userindex).flags.TargetUser = 0
        UserList(userindex).flags.TargetObj = 0
        UserList(userindex).flags.TargetObjMap = 0
        UserList(userindex).flags.TargetObjX = 0
    End If

Else
    If FoundSomething = 0 Then
        UserList(userindex).flags.TargetNPC = 0
        UserList(userindex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(userindex).flags.TargetUser = 0
        UserList(userindex).flags.TargetObj = 0
        UserList(userindex).flags.TargetObjMap = 0
        UserList(userindex).flags.TargetObjX = 0
        UserList(userindex).flags.TargetObjY = 0
    End If
End If

errhandler: Debug.Print "Error en LookAtTILE" & Err.Number & Err.Description
End Sub

Function FindDirection(pos As WorldPos, Target As WorldPos) As eHeading
'*****************************************************************
'Devuelve la direccion en la cual el target se encuentra
'desde pos, 0 si la direc es igual
'*****************************************************************
Dim x As Integer
Dim y As Integer

x = pos.x - Target.x
y = pos.y - Target.y

'NE
If Sgn(x) = -1 And Sgn(y) = 1 Then
    FindDirection = eHeading.NORTH
    Exit Function
End If

'NW
If Sgn(x) = 1 And Sgn(y) = 1 Then
    FindDirection = eHeading.WEST
    Exit Function
End If

'SW
If Sgn(x) = 1 And Sgn(y) = -1 Then
    FindDirection = eHeading.WEST
    Exit Function
End If

'SE
If Sgn(x) = -1 And Sgn(y) = -1 Then
    FindDirection = eHeading.SOUTH
    Exit Function
End If

'Sur
If Sgn(x) = 0 And Sgn(y) = -1 Then
    FindDirection = eHeading.SOUTH
    Exit Function
End If

'norte
If Sgn(x) = 0 And Sgn(y) = 1 Then
    FindDirection = eHeading.NORTH
    Exit Function
End If

'oeste
If Sgn(x) = 1 And Sgn(y) = 0 Then
    FindDirection = eHeading.WEST
    Exit Function
End If

'este
If Sgn(x) = -1 And Sgn(y) = 0 Then
    FindDirection = eHeading.EAST
    Exit Function
End If

'misma
If Sgn(x) = 0 And Sgn(y) = 0 Then
    FindDirection = 0
    Exit Function
End If

End Function

'[Barrin 30-11-03]
Public Function ItemNoEsDeMapa(ByVal Index As Integer) As Boolean

ItemNoEsDeMapa = ObjData(Index).OBJType <> eOBJType.otPuertas And _
            ObjData(Index).OBJType <> eOBJType.otFOROS And _
            ObjData(Index).OBJType <> eOBJType.otCARTELES And _
            ObjData(Index).OBJType <> eOBJType.otArboles And _
            ObjData(Index).OBJType <> eOBJType.otYacimiento And _
            ObjData(Index).OBJType <> eOBJType.otTELEPORT
End Function
'[/Barrin 30-11-03]

Public Function MostrarCantidad(ByVal Index As Integer) As Boolean
MostrarCantidad = ObjData(Index).OBJType <> eOBJType.otPuertas And _
            ObjData(Index).OBJType <> eOBJType.otFOROS And _
            ObjData(Index).OBJType <> eOBJType.otCARTELES And _
            ObjData(Index).OBJType <> eOBJType.otArboles And _
            ObjData(Index).OBJType <> eOBJType.otYacimiento And _
            ObjData(Index).OBJType <> eOBJType.otTELEPORT
End Function

Public Function EsObjetoFijo(ByVal OBJType As eOBJType) As Boolean

EsObjetoFijo = OBJType = eOBJType.otFOROS Or _
               OBJType = eOBJType.otCARTELES Or _
               OBJType = eOBJType.otArboles Or _
               OBJType = eOBJType.otYacimiento

End Function
