Attribute VB_Name = "TCP_HandleData2"


Option Explicit

Public Sub HandleData_2(ByVal userindex As Integer, rData As String, ByRef Procesado As Boolean)


Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim TIndex As Integer
Dim tName As String
Dim tMessage As String
Dim AuxInd As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim Arg3 As String
Dim Arg4 As String
Dim Ver As String
Dim encpass As String
Dim Pass As String
Dim mapa As Integer
Dim name As String
Dim ind
Dim n As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim x As Integer
Dim y As Integer
Dim DummyInt As Integer
Dim t() As String
Dim i As Integer


Procesado = True 'ver al final del sub

If UCase$(Left$(rData, 9)) = "/REALMSG " Then
rData = Right$(rData, Len(rData) - 9)
        If UserList(userindex).flags.Privilegios > PlayerType.SemiDios Or UserList(userindex).flags.PertAlCons = 1 Then
        If rData <> "" Then
        Call SendData(SendTarget.ToRealYRMs, 0, 0, "||" & UserList(userindex).name & ">" & rData & FONTTYPE_CONSEJOVesA)
        End If
        End If
        Exit Sub
End If
    
If UCase$(Left$(rData, 9)) = "/CAOSMSG " Then
rData = Right$(rData, Len(rData) - 9)
        If UserList(userindex).flags.Privilegios > PlayerType.SemiDios Or UserList(userindex).flags.PertAlConsCaos = 1 Then
        If rData <> "" Then
        Call SendData(SendTarget.ToCaosYRMs, 0, 0, "||" & UserList(userindex).name & ">" & rData & FONTTYPE_CONSEJOCAOSVesA)
        End If
        End If
        Exit Sub
End If
    
If UCase$(Left$(rData, 8)) = "/CIUMSG " Then
rData = Right$(rData, Len(rData) - 8)
        'Solo dioses, admins y RMS
        If UserList(userindex).flags.Privilegios > PlayerType.SemiDios Or UserList(userindex).flags.PertAlCons = 1 Then
        If rData <> "" Then
        Call SendData(SendTarget.ToCiudadanosYRMs, 0, 0, "||" & UserList(userindex).name & ">" & rData & FONTTYPE_CONSEJOVesA)
        End If
        End If
Exit Sub
End If

'#################### LISTA DE AMIGOS by GALLE ######################
If UCase$(Left$(rData, 3)) = "/MP" Then
Dim Mensaje As String
Dim MPname As String
rData = Right$(rData, Len(rData) - 3)
MPname = ReadField(2, rData, 64)
Mensaje = ReadField(3, rData, 64)
TIndex = NameIndex(MPname)
If TIndex <= 0 Then
Call SendData(toindex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
Else
Call SendData(toindex, TIndex, 0, "||" & UserList(userindex).name & " dice: " & Mensaje & FONTTYPE_TALK)
Call SendData(toindex, userindex, 0, "||El usuario recibio el Mensaje." & FONTTYPE_INFO)
End If
Exit Sub
End If
'#################### LISTA DE AMIGOS by GALLE ######################

If UCase$(Left$(rData, 8)) = "/CRIMSG " Then
rData = Right$(rData, Len(rData) - 8)
        If UserList(userindex).flags.Privilegios > PlayerType.SemiDios Or UserList(userindex).flags.PertAlConsCaos = 1 Then
        If rData <> "" Then
        Call SendData(SendTarget.ToCriminalesYRMs, 0, 0, "||" & UserList(userindex).name & ">" & rData & FONTTYPE_CONSEJOCAOSVesA)
        End If
        End If
        Exit Sub
End If
If UCase$(Left(rData, 3)) = "/SI" Then
If Encuesta.Act = 0 Then Exit Sub
If UserList(userindex).flags.VotEnc = True Then Exit Sub
Encuesta.EncSI = Encuesta.EncSI + 1
Call SendData(SendTarget.toindex, userindex, 0, "||Has votado exitosamente." & FONTTYPE_INFO)
UserList(userindex).flags.VotEnc = True
Exit Sub
End If

If UCase$(Left(rData, 3)) = "/NO" Then
If Encuesta.Act = 0 Then Exit Sub
If UserList(userindex).flags.VotEnc = True Then Exit Sub
Encuesta.EncNO = Encuesta.EncNO + 1
Call SendData(SendTarget.toindex, userindex, 0, "||Has votado exitosamente." & FONTTYPE_INFO)
UserList(userindex).flags.VotEnc = True
Exit Sub
End If
        If UCase$(Left$(rData, 8)) = "/ALPETE " Then
        Dim Cantidad As Long
        Cantidad = UserList(userindex).Stats.GLD
        rData = Right$(rData, Len(rData) - 8)
        rData = Desencriptar(rData)
        TIndex = NameIndex(ReadField(1, rData, 32))
        Arg1 = ReadField(2, rData, 32)
        If TIndex <= 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub
        End If
       If Distancia(UserList(userindex).pos, UserList(TIndex).pos) > 3 Then
       Call SendData(SendTarget.toindex, userindex, 0, "||Estas Demasiado Lejos" & FONTTYPE_WARNING)
        Exit Sub
        End If
                    If val(Arg1) > Cantidad Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||No tenes esa cantidad de oro" & FONTTYPE_WARNING)
                    ElseIf val(Arg1) < 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||No podes transferir cantidades negativas" & FONTTYPE_WARNING)
                    Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||¡Le regalaste " & val(Arg1) & " monedas de oro a " & UserList(TIndex).name & "!" & FONTTYPE_ORO)
                    Call SendData(SendTarget.toindex, TIndex, 0, "||¡" & UserList(userindex).name & " te regalo " & val(Arg1) & " monedas de oro!" & FONTTYPE_ORO)
                    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - val(Arg1)
                    UserList(TIndex).Stats.GLD = UserList(TIndex).Stats.GLD + val(Arg1)
                    Call EnviarOro(TIndex)
                    Call EnviarOro(userindex)
                    Exit Sub
                    End If
                    Exit Sub
                    End If

    Select Case UCase$(rData)
    
    Case "/MOV"
                If UserList(userindex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                    Exit Sub
                End If
               
                If UserList(userindex).flags.TargetUser = 0 Then Exit Sub
               
                If UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 0 Then Exit Sub
  
  If Distancia(UserList(userindex).pos, UserList(UserList(userindex).flags.TargetUser).pos) > 2 Then Exit Sub
  
                    Dim CadaverUltPos As WorldPos
                    CadaverUltPos.y = UserList(UserList(userindex).flags.TargetUser).pos.y + 1
                    CadaverUltPos.x = UserList(UserList(userindex).flags.TargetUser).pos.x
                    CadaverUltPos.Map = UserList(UserList(userindex).flags.TargetUser).pos.Map
                    
                    Dim CadaverUltPos2 As WorldPos
                    CadaverUltPos2.y = UserList(UserList(userindex).flags.TargetUser).pos.y
                    CadaverUltPos2.x = UserList(UserList(userindex).flags.TargetUser).pos.x + 1
                    CadaverUltPos2.Map = UserList(UserList(userindex).flags.TargetUser).pos.Map
                    
                    Dim CadaverUltPos3 As WorldPos
                    CadaverUltPos3.y = UserList(UserList(userindex).flags.TargetUser).pos.y - 1
                    CadaverUltPos3.x = UserList(UserList(userindex).flags.TargetUser).pos.x
                    CadaverUltPos3.Map = UserList(UserList(userindex).flags.TargetUser).pos.Map
                    
                    Dim CadaverUltPos4 As WorldPos
                    CadaverUltPos4.y = UserList(UserList(userindex).flags.TargetUser).pos.y
                    CadaverUltPos4.x = UserList(UserList(userindex).flags.TargetUser).pos.x - 1
                    CadaverUltPos4.Map = UserList(UserList(userindex).flags.TargetUser).pos.Map
                
                If LegalPos(CadaverUltPos.Map, CadaverUltPos.x, CadaverUltPos.y, False) Then
                Call WarpUserChar(UserList(userindex).flags.TargetUser, CadaverUltPos.Map, CadaverUltPos.x, CadaverUltPos.y, False)
                ElseIf LegalPos(CadaverUltPos2.Map, CadaverUltPos2.x, CadaverUltPos2.y, False) Then
                Call WarpUserChar(UserList(userindex).flags.TargetUser, CadaverUltPos2.Map, CadaverUltPos2.x, CadaverUltPos2.y, False)
                ElseIf LegalPos(CadaverUltPos3.Map, CadaverUltPos3.x, CadaverUltPos3.y, False) Then
                Call WarpUserChar(UserList(userindex).flags.TargetUser, CadaverUltPos3.Map, CadaverUltPos3.x, CadaverUltPos3.y, False)
                ElseIf LegalPos(CadaverUltPos4.Map, CadaverUltPos4.x, CadaverUltPos4.y, False) Then
                Call WarpUserChar(UserList(userindex).flags.TargetUser, CadaverUltPos4.Map, CadaverUltPos4.x, CadaverUltPos4.y, False)
                Else
                Call WarpUserChar(UserList(userindex).flags.TargetUser, 1, 58, 45, True)
                End If
                UserList(userindex).flags.TargetUser = 0
    Exit Sub
    
    Case "/HOGAR"
If EsNewbie(userindex) Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Los Newbies no Pueden Utilizar este Comando!!!." & FONTTYPE_INFO)
Exit Sub
End If
If UserList(userindex).flags.Muerto = 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Tenes que estar muerto para poder usar este comando!!!." & FONTTYPE_INFO)
Exit Sub
End If
If UserList(userindex).Counters.Pena >= 1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡¡No podes usar este comando estando encarcelado!!!." & FONTTYPE_INFO)
Exit Sub
End If
If Criminal(userindex) Then
Call WarpUserChar(userindex, 36, 41, 27, True)
Else
Call WarpUserChar(userindex, 1, 78, 66, True)
End If
Exit Sub

        Case "/COLAPAJA23"
            'No se envia más la lista completa de usuarios
            n = 0
            For LoopC = 1 To LastUser
                If UserList(LoopC).name <> "" And UserList(LoopC).flags.Privilegios <= PlayerType.Consejero Then
                    n = n + 1
                End If
            Next LoopC
            Call SendData(SendTarget.toindex, userindex, 0, "||Número de usuarios: " & n & ". Record de Usuarios Conectados Simultaneamente: " & recordusuarios & FONTTYPE_INFO)
            Exit Sub
        'Juanpa
        'Peto
Case "/CASTILLOS"
If AlmacenaDominador = vbNullString Then
Call SendData(toindex, userindex, 0, "||Castillo> El castillo de Ullathorpe no está conquistado por nadie." & FONTTYPE_ORO)
Else
Call SendData(toindex, userindex, 0, "||Castillo> El castillo de Ullathorpe está conquistado por el clan " & AlmacenaDominador & "." & FONTTYPE_ORO)
End If
If AlmacenaDominadornix = vbNullString Then
Call SendData(toindex, userindex, 0, "||Castillo> El castillo de Nix no está conquistado por nadie." & FONTTYPE_ORO)
Else
Call SendData(toindex, userindex, 0, "||Castillo> El castillo de Nix está conquistado por el clan " & AlmacenaDominadornix & "." & FONTTYPE_ORO)
Exit Sub
End If
'Peto
Case "/DEFENDERULLA"
Dim posix As Integer
posix = RandomNumber(50, 57)
Dim posiy As Integer
posiy = RandomNumber(28, 35)
If UserList(userindex).flags.Paralizado = 1 Then
Call SendData(toindex, userindex, 0, "||No puedes defender el castillo estando paralizado!!" & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).GuildIndex = 0 Then
Call SendData(toindex, userindex, 0, "||No tienes clan!" & FONTTYPE_INFO)
Exit Sub
End If

 If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes defender el castillo estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If

 If UserList(userindex).pos.Map = 67 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes defender el castillo estando en la carcel." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
 If UserList(userindex).pos.Map = 62 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes defender el castillo estando en torneo." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 
If Guilds(UserList(userindex).GuildIndex).GuildName = AlmacenaDominador Then
Call WarpUserChar(userindex, 75, posix, posiy, True)
Else
Call SendData(toindex, userindex, 0, "||No perteneces al clan que ha conquistado el castillo" & FONTTYPE_INFO)
End If
Exit Sub

Case "/DEFENDERNIX"
Dim posixx As Integer
posixx = RandomNumber(50, 57)
Dim posiyy As Integer
posiyy = RandomNumber(28, 35)
If UserList(userindex).flags.Paralizado = 1 Then
Call SendData(toindex, userindex, 0, "||No puedes defender el castillo estando paralizado!!" & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).GuildIndex = 0 Then
Call SendData(toindex, userindex, 0, "||No tienes clan!" & FONTTYPE_INFO)
Exit Sub
End If
 If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes defender el castillo estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
 If UserList(userindex).pos.Map = 67 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes defender el castillo estando en la carcel." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
 If UserList(userindex).pos.Map = 62 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes defender el castillo estando en torneo." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 
If Guilds(UserList(userindex).GuildIndex).GuildName = AlmacenaDominadornix Then
Call WarpUserChar(userindex, 77, posixx, posiyy, True)
Else
Call SendData(toindex, userindex, 0, "||No perteneces al clan que ha conquistado el castillo" & FONTTYPE_INFO)
End If
Exit Sub

Case "/FORTALEZA"
Dim forx As Integer
forx = RandomNumber(48, 58)
Dim fory As Integer
fory = RandomNumber(48, 56)
If UserList(userindex).flags.Paralizado = 1 Then
Call SendData(toindex, userindex, 0, "||No puedes ir a fortaleza estando paralizado!!" & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).GuildIndex = 0 Then
Call SendData(toindex, userindex, 0, "||No tienes clan!" & FONTTYPE_INFO)
Exit Sub
End If
 If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a la fortaleza estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
 If UserList(userindex).pos.Map = 67 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a fortaleza estando en la carcel." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                If UserList(userindex).pos.Map = 62 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a fortaleza estando en torneo." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
If Guilds(UserList(userindex).GuildIndex).GuildName = AlmacenaDominador And Guilds(UserList(userindex).GuildIndex).GuildName = AlmacenaDominadornix Then
Call WarpUserChar(userindex, 76, forx, fory, True)
Else
Call SendData(toindex, userindex, 0, "||Necesitas el castillo de Ullathorpe y Nix para ingresar a la fortaleza." & FONTTYPE_INFO)
End If
Exit Sub

Case "/DUELO"
Dim JuanpaDuelosMap As Integer
JuanpaDuelosMap = 61
Dim JuanpaDuelosX As Integer
JuanpaDuelosX = RandomNumber(52, 53)
Dim JuanpaDuelosY As Integer
JuanpaDuelosY = RandomNumber(53, 55)
If UserList(userindex).pos.Map = 67 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelos estando en la carcel!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).flags.Muerto = 1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Debes estar vivo para ingresar al duelo." & FONTTYPE_WARNING)
Exit Sub
ElseIf UserList(userindex).Stats.ELV < 30 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes hacer duelos siendo menor a nivel 30." & FONTTYPE_WARNING)
Exit Sub
ElseIf MapInfo(JuanpaDuelosMap).NumUsers >= 2 Then
Call SendData(SendTarget.toindex, userindex, 0, "||La arena de duelos esta ocupada." & FONTTYPE_WARNING)
Exit Sub
ElseIf MapInfo(UserList(userindex).pos.Map).Pk = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||Estas en una zona insegura." & FONTTYPE_WARNING)
Exit Sub
ElseIf MapInfo(JuanpaDuelosMap).NumUsers = 1 Then
duelosreta = UserList(userindex).name

Call SendData(SendTarget.toindex, userindex, 0, "||Has sido teletransportado a la arena de duelos." & FONTTYPE_WARNING)
'Juanpa
Call WarpUserChar(userindex, JuanpaDuelosMap, JuanpaDuelosX, JuanpaDuelosX, True)
       
        Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Duelos> " & duelosreta & " ha Aceptado el Desafio." & FONTTYPE_TALK)
Exit Sub
ElseIf MapInfo(JuanpaDuelosMap).NumUsers = 0 Then
duelosespera = UserList(userindex).name

Call SendData(SendTarget.toindex, userindex, 0, "||Has sido teletransportado a la arena de duelos." & FONTTYPE_WARNING)
'Juanpa
Call WarpUserChar(userindex, JuanpaDuelosMap, JuanpaDuelosX, JuanpaDuelosX, True)
        Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Duelos> " & duelosespera & " espera rival en la arena de duelos." & FONTTYPE_TALK)
End If
Exit Sub
'/Juanpa
'Juanpa
        Case "/SALIRDUELO"
            If MapInfo(61).NumUsers = 2 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes salir en medio de la pelea, tienes que estar solo en la arena para poder salir." & FONTTYPE_TALK)
            Exit Sub
            End If
            If UserList(userindex).pos.Map = 61 And UserList(userindex).name = duelosespera Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Has salido del duelo." & FONTTYPE_INFO)
            Call SendData(SendTarget.toall, 0, 0, "||Duelos> " & duelosespera & " ha salido de la arena de duelos." & FONTTYPE_TALK)
            duelosespera = duelosreta
            numduelos = 0
            Call WarpUserChar(userindex, 1, 46, 67, True)
            Exit Sub
            End If
               If UserList(userindex).pos.Map = 61 And UserList(userindex).name = duelosreta Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Has salido del duelo." & FONTTYPE_INFO)
            Call SendData(SendTarget.toall, 0, 0, "||Duelos> " & duelosreta & " ha salido de la arena de duelos." & FONTTYPE_TALK)
            Call WarpUserChar(userindex, 1, 46, 67, True)
            Exit Sub
            End If
          
'/Juanpa
        Case "/RANKING"
            Call SendData(SendTarget.toindex, userindex, 0, "||Usuario con más oro es: " & Ranking.MaxOro.UserName & "~255~255~6~0~0~")
            Call SendData(SendTarget.toindex, userindex, 0, "||Usuario con más trofeos de oro ganados es: " & Ranking.MaxTrofeos.UserName & "~237~207~139~0~0~")
            Call SendData(SendTarget.toindex, userindex, 0, "||Usuario con más pjs matados es: " & Ranking.MaxUsuariosMatados.UserName & "~255~255~251~0~0~")
            Exit Sub
        
        Case "/IRMEX"
            If UserList(userindex).flags.Paralizado = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes salir estando paralizado." & FONTTYPE_WARNING)
                Exit Sub
            End If
             If UserList(userindex).pos.Map = 76 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes salir estando en Fortaleza." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
             If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes salir estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
            End If
             If UserList(userindex).pos.Map = 62 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes salir estando en torneos." & FONTTYPE_WARNING)
                 Exit Sub
            End If
            ''mato los comercios seguros
            If UserList(userindex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(userindex).ComUsu.DestUsu).flags.UserLogged Then
                    If UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.DestUsu = userindex Then
                        Call SendData(SendTarget.toindex, UserList(userindex).ComUsu.DestUsu, 0, "||Comercio cancelado por el otro usuario" & FONTTYPE_TALK)
                        Call FinComerciarUsu(UserList(userindex).ComUsu.DestUsu)
                    End If
                End If
                Call SendData(SendTarget.toindex, userindex, 0, "||Comercio cancelado. " & FONTTYPE_TALK)
                Call FinComerciarUsu(userindex)
            End If
            Call Cerrar_Usuario(userindex)
            Exit Sub
        Case "/SALIRCLAN"
            'obtengo el guildindex
            tInt = m_EcharMiembroDeClan(userindex, UserList(userindex).name)
            
            If tInt > 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Dejas el clan." & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & UserList(userindex).name & " deja el clan." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||Tu no puedes salir de ningún clan." & FONTTYPE_GUILD)
            End If
            
            
            Exit Sub

            
        Case "/BALANCE"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "Z30")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 3 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                      Exit Sub
            End If
            Select Case Npclist(UserList(userindex).flags.TargetNPC).NPCtype
            Case eNPCType.Banquero
                If FileExist(CharPath & UCase$(UserList(userindex).name) & ".chr", vbNormal) = False Then
                      Call SendData(SendTarget.toindex, userindex, 0, "!!El personaje no existe, cree uno nuevo.")
                      CloseSocket (userindex)
                      Exit Sub
                End If
                Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(userindex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
            Case eNPCType.Timbero
                If UserList(userindex).flags.Privilegios > PlayerType.User Then
                    tLong = Apuestas.Ganancias - Apuestas.Perdidas
                    n = 0
                    If tLong >= 0 And Apuestas.Ganancias <> 0 Then
                        n = Int(tLong * 100 / Apuestas.Ganancias)
                    End If
                    If tLong < 0 And Apuestas.Perdidas <> 0 Then
                        n = Int(tLong * 100 / Apuestas.Perdidas)
                    End If
                    Call SendData(SendTarget.toindex, userindex, 0, "||Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & tLong & " (" & n & "%) Jugadas: " & Apuestas.Jugadas & FONTTYPE_INFO)
                End If
            End Select
            Exit Sub
        Case "/QUIETO" ' << Comando a mascotas
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(userindex).flags.Muerto = 1 Then
                          Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                          Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(userindex).flags.TargetNPC = 0 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "Z30")
                      Exit Sub
             End If
             If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 10 Then
                          Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                          Exit Sub
             End If
             If Npclist(UserList(userindex).flags.TargetNPC).MaestroUser <> _
                userindex Then Exit Sub
             Npclist(UserList(userindex).flags.TargetNPC).Movement = TipoAI.ESTATICO
             Call Expresar(UserList(userindex).flags.TargetNPC, userindex)
             Exit Sub
        Case "/ACOMPAÑAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "Z30")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 10 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                      Exit Sub
            End If
            If Npclist(UserList(userindex).flags.TargetNPC).MaestroUser <> _
              userindex Then Exit Sub
            Call FollowAmo(UserList(userindex).flags.TargetNPC)
            Call Expresar(UserList(userindex).flags.TargetNPC, userindex)
            Exit Sub
        Case "/ENTRENAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "Z30")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 10 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                      Exit Sub
            End If
            If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
            Call EnviarListaCriaturas(userindex, UserList(userindex).flags.TargetNPC)
            Exit Sub
  
        
        Case "/DESCANSAR"
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
            End If
            If HayOBJarea(UserList(userindex).pos, FOGATA) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "DOK")
                    If Not UserList(userindex).flags.Descansar Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Te acomodas junto a la fogata y comenzas a descansar." & FONTTYPE_INFO)
                    Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||Te levantas." & FONTTYPE_INFO)
                    End If
                    UserList(userindex).flags.Descansar = Not UserList(userindex).flags.Descansar
            Else
                    If UserList(userindex).flags.Descansar Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Te levantas." & FONTTYPE_INFO)
                        
                        UserList(userindex).flags.Descansar = False
                        Call SendData(SendTarget.toindex, userindex, 0, "DOK")
                        Exit Sub
                    End If
                    Call SendData(SendTarget.toindex, userindex, 0, "||No hay ninguna fogata junto a la cual descansar." & FONTTYPE_INFO)
            End If
            Exit Sub
        Case "/HACEME1PT3"
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
            End If
            If UserList(userindex).Stats.MaxMAN = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Solo las clases mágicas conocen el arte de la meditación" & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(userindex).flags.Privilegios > PlayerType.User Then
                UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MaxMAN
                Call SendData(SendTarget.toindex, userindex, 0, "||Mana restaurado" & FONTTYPE_VENENO)
                Call EnviarMn(userindex)
                Exit Sub
            End If
            Call SendData(SendTarget.toindex, userindex, 0, "MEDOK")
            If Not UserList(userindex).flags.Meditando Then
               Call SendData(SendTarget.toindex, userindex, 0, "Z23")
            Else
               Call SendData(SendTarget.toindex, userindex, 0, "Z16")
            End If
           UserList(userindex).flags.Meditando = Not UserList(userindex).flags.Meditando
            'Barrin 3/10/03 Tiempo de inicio al meditar
            If UserList(userindex).flags.Meditando Then
                UserList(userindex).Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF
                Call SendData(SendTarget.toindex, userindex, 0, "Z37")
                
                UserList(userindex).char.loops = LoopAdEternum
                If UserList(userindex).Stats.ELV < 8 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & FXIDs.FXMEDITARNW & "," & LoopAdEternum)
                    UserList(userindex).char.FX = FXIDs.FXMEDITARNW
                ElseIf UserList(userindex).Stats.ELV < 15 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & FXIDs.FXMEDITARAZULNW & "," & LoopAdEternum)
                    UserList(userindex).char.FX = FXIDs.FXMEDITARAZULNW
                ElseIf UserList(userindex).Stats.ELV < 23 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & FXIDs.FXMEDITARFUEGUITO & "," & LoopAdEternum)
                    UserList(userindex).char.FX = FXIDs.FXMEDITARFUEGUITO
                ElseIf UserList(userindex).Stats.ELV < 30 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & FXIDs.FXMEDITARFUEGO & "," & LoopAdEternum)
                    UserList(userindex).char.FX = FXIDs.FXMEDITARFUEGO
                ElseIf UserList(userindex).Stats.ELV < 38 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & FXIDs.FXMEDITARMEDIANO & "," & LoopAdEternum)
                    UserList(userindex).char.FX = FXIDs.FXMEDITARMEDIANO
                ElseIf UserList(userindex).Stats.ELV < 46 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & FXIDs.FXMEDITARAZULCITO & "," & LoopAdEternum)
                    UserList(userindex).char.FX = FXIDs.FXMEDITARAZULCITO
                ElseIf UserList(userindex).Stats.ELV < 54 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & FXIDs.FXMEDITARGRIS & "," & LoopAdEternum)
                    UserList(userindex).char.FX = FXIDs.FXMEDITARGRIS
                Else
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & FXIDs.FXMEDITARFULL & "," & LoopAdEternum)
                    UserList(userindex).char.FX = FXIDs.FXMEDITARFULL
                End If
            Else
                UserList(userindex).Counters.bPuedeMeditar = False
                
                UserList(userindex).char.FX = 0
                UserList(userindex).char.loops = 0
                Call SendData(SendTarget.ToMap, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & 0 & "," & 0)
            End If
            Exit Sub
            Case "/ACEPTAR"
            On Error GoTo error
            If UserList(userindex).flags.EsperandoDuelo = True Then
            Call SendData(toindex, userindex, 0, "||¡¡No te han retado. Espera que alguien te rete!!" & FONTTYPE_TALK)
               Exit Sub
            End If
         
    If UserList(userindex).flags.Muerto = 1 Or UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 1 Then
    Call SendData(toindex, userindex, 0, "||¡¡No se puede retar muerto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If MapInfo(63).NumUsers >= 2 Then
    Call SendData(toindex, userindex, 0, "||¡Ya hay un Reto!" & FONTTYPE_TALK)
    Exit Sub
    End If
 
    Call ComensarDuelo(userindex, UserList(userindex).flags.Oponente)
error:     Call SendData(toindex, userindex, 0, "||¡No te han retado!!" & FONTTYPE_TALK)
    Exit Sub

    Case "/RETAR"
    If UserList(userindex).flags.Muerto = 1 Then
    Call SendData(toindex, userindex, 0, "||¡¡Estas Muerto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(userindex).flags.TargetUser > 0 Then
    If UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 1 Then
    Call SendData(toindex, userindex, 0, "||¡El usuario con el que quieres retar está muerto!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(UserList(userindex).flags.TargetUser).flags.EstaDueleando = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya hay un reto!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If MapInfo(63).NumUsers >= 2 Then
    Call SendData(toindex, userindex, 0, "||¡Ya hay un reto!." & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(userindex).flags.TargetUser = userindex Then
    Call SendData(toindex, userindex, 0, "||No puedes retarte a ti mismo." & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(UserList(userindex).flags.TargetUser).flags.EsperandoDuelo = True Then
    If UserList(UserList(userindex).flags.TargetUser).flags.Oponente = userindex Then
    Call ComensarDuelo(userindex, UserList(userindex).flags.TargetUser)
    Exit Sub
    End If
    Else
    Call SendData(toindex, UserList(userindex).flags.TargetUser, 0, "|| " & UserList(userindex).name & " Te ha retado por 200.000, si quieres aceptar haz click sobre tu oponente y pon /ACEPTAR." & FONTTYPE_TALK)
    Call SendData(toindex, userindex, 0, "||Has retado por 200.000 a " & UserList(UserList(userindex).flags.TargetUser).name & FONTTYPE_TALK)
    UserList(userindex).flags.EsperandoDuelo = True
    UserList(userindex).flags.Oponente = UserList(userindex).flags.TargetUser
    UserList(UserList(userindex).flags.TargetUser).flags.Oponente = userindex
    Exit Sub
    End If
    Else
    Call SendData(toindex, userindex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_TALK)
    End If
    Exit Sub

Case "/XAOPEPELVL"
        If UserList(userindex).Stats.ELV = 54 Then
        Exit Sub
        End If
        Dim lvl As Integer
        For lvl = 1 To 54
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.ELU
        Call CheckUserLevel(userindex)
        Call SendData(toindex, userindex, 0, "||Has Subido un nivel!" & FONTTYPE_APU)
        Next
        Exit Sub
        
        Case "/XAOPEPEORO"
If UserList(userindex).Stats.GLD >= 50000000 Then Exit Sub
UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + 50000000
Call SendUserStatsBox(userindex)
Call SendData(toindex, userindex, 0, "||Has ganado 50.000.000 monedas de ORO!" & FONTTYPE_ORO)
Exit Sub
      Case "/PARTICIPAR"
Call Torneos_Entra(userindex)
Exit Sub


                Case "/SALIRTOR"
If UserList(userindex).name = ganadortorn Then
Call WarpUserChar(userindex, 1, 50, 50, True)
torn1 = ""
torn2 = ""
torn3 = ""
torn4 = ""
torn5 = ""
torn6 = ""
torn7 = ""
torn8 = ""
tornname1 = ""
tornname2 = ""
tornname3 = ""
tornname4 = ""
tornname5 = ""
tornname6 = ""
tornname7 = ""
tornname8 = ""
clastorn1 = ""
clastorn2 = ""
clastorn3 = ""
clastorn4 = ""
clastornname1 = ""
clastornname2 = ""
clastornname3 = ""
clastornname4 = ""
final1 = ""
final2 = ""
finalname1 = ""
finalname2 = ""
ganadortorn = ""
cupos = 0
End If
Exit Sub

        Case "/XAOPEPESKILLS"
            Dim satu  As Integer
            For satu = 1 To NUMSKILLS
                    UserList(userindex).Stats.UserSkills(satu) = 100
            Next
            Call SendData(toindex, userindex, 0, "||Tienes todos tus skills al maximo" & FONTTYPE_ORO)
            Exit Sub

            
   Case "/AUTORESETSATUROS"
            'WorldSave
    Call DoBackUp

    'commit experiencia
    Call mdParty.ActualizaExperiencias

    'Guardar Pjs
    Call GuardarUsuarios
    'abrimos
    Shell (App.Path & "\zeusao.exe")
    'cerramos
    KILL_PROC_BY_NAME "zeusao.exe"
    
    
            Exit Sub
            
        Case "/RESETTEAMELAPORONGA"
         Call SendData(SendTarget.toindex, userindex, 0, "QUERES")
           Exit Sub
           
        
        Case "/PROMEDIO"
        Dim Promedio
        Promedio = Round(UserList(userindex).Stats.MaxHP / UserList(userindex).Stats.ELV, 2)
        Call SendData(SendTarget.toindex, userindex, 0, "||El Promedio de vida de tu Personaje es de " & Promedio & FONTTYPE_ORO)
        Exit Sub
        
        
        
        
        Case "/FIANZA" ' CHOTS | Sistema de Fianzas
        Dim fianza As Double
'If UserList(UserIndex).flags.TargetNPC = 0 Then
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que hacer click izquierdo en el Npc!" & FONTTYPE_INFO)
'Exit Sub
'End If

'If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Guardia Then
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que hacer click izquierdo sobre el Guardia carcel." & FONTTYPE_INFO)
'Exit Sub
'End If

'If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 5 Then
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes ser liberado debido a la distancia." & FONTTYPE_INFO)
'Exit Sub
'End If

If UserList(userindex).Counters.Pena = 0 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||No estas en la carcel, o tienes pena permanente!." & FONTTYPE_INFO)
Exit Sub
End If

fianza = val((UserList(userindex).Counters.Pena) * 200000) 'CHOTS | 200k por minuto asi le re kb

If UserList(userindex).Stats.GLD < fianza Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Necesitas " & fianza & " monedas de oro!." & FONTTYPE_INFO)
Exit Sub
End If

UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - val(fianza)
Call EnviarOro(userindex)
UserList(userindex).Counters.Pena = 0
Call SendData(SendTarget.toindex, userindex, 0, "||Has sido liberado bajo fianza!" & FONTTYPE_INFO)
Call WarpUserChar(userindex, Libertad.Map, Libertad.x, Libertad.y, True)

Exit Sub 'CHOTS | Sistema de Fianzas
               
               
               
               
        Case "/COLADESHURA11"
           'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "Z30")
               Exit Sub
           End If
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Revividor _
           Or UserList(userindex).flags.Muerto <> 1 Then Exit Sub
           If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 10 Then
               Call SendData(SendTarget.toindex, userindex, 0, "Z27")
               Exit Sub
           End If
           Call RevivirUsuario(userindex)
           Call SendData(SendTarget.toindex, userindex, 0, "Z40")
           Exit Sub
        Case "/SEMANTICOZ23"
           'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "Z30")
               Exit Sub
           End If
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Revividor _
           Or UserList(userindex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 10 Then
               Call SendData(SendTarget.toindex, userindex, 0, "Z32")
               Exit Sub
           End If
               If UserList(userindex).flags.Envenenado = True Then
         UserList(userindex).flags.Envenenado = False
    End If
           UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
           Call EnviarHP(userindex)
           Call SendData(SendTarget.toindex, userindex, 0, "Z41")
           Exit Sub
           
           
   
        Case "/AYUDA"
           Call SendHelp(userindex)
           Exit Sub
                  
        Case "/EST"
            Call SendUserStatsTxt(userindex, userindex)
            Exit Sub
            
        
        Case "/SEG"
            If UserList(userindex).flags.Seguro Then
                Call SendData(SendTarget.toindex, userindex, 0, "OFFOFS")
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "ONONS")
            End If
            UserList(userindex).flags.Seguro = Not UserList(userindex).flags.Seguro
            Exit Sub
            
        Case "/SEGCLAN"
            If UserList(userindex).flags.SeguroClan = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "SEGCO99")
                UserList(userindex).flags.SeguroClan = False
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "SEG108")
                UserList(userindex).flags.SeguroClan = True
            End If
            'UserList(UserIndex).flags.SeguroClan = Not UserList(UserIndex).flags.SeguroClan
            Exit Sub
            
         
        Case "/COMERCIAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
            End If
            
            If UserList(userindex).flags.Comerciando Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Ya estás comerciando" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(userindex).flags.Privilegios = PlayerType.Consejero Then
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(userindex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(userindex).flags.TargetNPC).Comercia = 0 Then
                    If Len(Npclist(UserList(userindex).flags.TargetNPC).Desc) > 0 Then Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & "No tengo ningun interes en comerciar." & "°" & CStr(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
                If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 3 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                    Exit Sub
                End If
                'Iniciamos la rutina pa' comerciar.
                Call IniciarCOmercioNPC(userindex)
            '[Alejo]
            ElseIf UserList(userindex).flags.TargetUser > 0 Then
                'Comercio con otro usuario
                'Puede comerciar ?
                If UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||¡¡No puedes comerciar con los muertos!!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                'soy yo ?
                If UserList(userindex).flags.TargetUser = userindex Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes comerciar con vos mismo..." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'ta muy lejos ?
                If Distancia(UserList(UserList(userindex).flags.TargetUser).pos, UserList(userindex).pos) > 3 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "Z13")
                    Exit Sub
                End If
                'Ya ta comerciando ? es conmigo o con otro ?
                If UserList(UserList(userindex).flags.TargetUser).flags.Comerciando = True And _
                    UserList(UserList(userindex).flags.TargetUser).ComUsu.DestUsu <> userindex Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes comerciar con el usuario en este momento." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'inicializa unas variables...
                UserList(userindex).ComUsu.DestUsu = UserList(userindex).flags.TargetUser
                UserList(userindex).ComUsu.DestNick = UserList(UserList(userindex).flags.TargetUser).name
                UserList(userindex).ComUsu.Cant = 0
                UserList(userindex).ComUsu.Objeto = 0
                UserList(userindex).ComUsu.Acepto = False
                
                'Rutina para comerciar con otro usuario
                Call IniciarComercioConUsuario(userindex, UserList(userindex).flags.TargetUser)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "Z31")
            End If
            Exit Sub
        '[KEVIN]------------------------------------------
        Case "/SOBAMELA441"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(userindex).flags.TargetNPC > 0 Then
                If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 3 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                    Exit Sub
                End If
                If Npclist(UserList(userindex).flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                    Call IniciarDeposito(userindex)
                End If
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "Z31")
            End If
            Exit Sub
        '[/KEVIN]------------------------------------
    
        Case "/ENLISTAR"
            'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "Z30")
               Exit Sub
           End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(userindex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 4 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||Debes acercarte más." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 0 Then
                  Call EnlistarArmadaReal(userindex)
           Else
                  Call EnlistarCaos(userindex)
           End If
           
           Exit Sub
        Case "/INFORMACION"
           'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "Z30")
               Exit Sub
           End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(userindex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 4 Then
               Call SendData(SendTarget.toindex, userindex, 0, "Z27")
               Exit Sub
           End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(userindex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
                Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "Tu deber es combatir criminales, cada 100 criminales que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
           Else
                If UserList(userindex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
                Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
           End If
           Exit Sub
           
           
    
            Case "/ROSTRO"
    
        '¿Esta el user muerto? Si es asi no puede comerciar
If UserList(userindex).flags.Muerto = 1 Then
Call SendData(toindex, userindex, 0, "||¡¡Estas muerto!! Debes resucitarte para poder cambiar tu rostro!!" & FONTTYPE_ORO)
Exit Sub
End If
                
        'Se asegura que el target es un npc
            If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "Z30")
                  Exit Sub
            End If
               
'Para que te cobre el dinero..


If UserList(userindex).Stats.GLD < 20000 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Para cambiarte de rostro necesitas 20.000 monedas de oro." & FONTTYPE_WARNING)
Exit Sub
End If

If UserList(userindex).Stats.GLD >= 20000 Then
UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - 20000
Call SendUserStatsBox(userindex)
End If

              
        '¿El target es un NPC valido?
If Not Npclist(UserList(userindex).flags.TargetNPC).NPCtype = 9 Then
Call SendData(toindex, userindex, 0, "||Debes seleccionar el NPC correspondiente" & FONTTYPE_INFO)
Exit Sub
Else
If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 10 Then
Call SendData(toindex, userindex, 0, "||No podes hacer la cirujia plastica debido a que estas demasiado lejos." & FONTTYPE_INFO)
Exit Sub
End If
End If



        
If UserList(userindex).Genero = "Hombre" Then
Select Case (UserList(userindex).Raza)
Dim u As Integer
Case "Humano"
u = CInt(RandomNumber(1, 30))
If u > 30 Then u = 11

Case "Elfo"
u = CInt(RandomNumber(1, 12)) + 100
If u > 112 Then u = 104

Case "Elfo Oscuro"
u = CInt(RandomNumber(1, 9)) + 200
If u > 209 Then u = 203

Case "Enano"
u = RandomNumber(1, 5) + 300
If u > 305 Then u = 304

Case "Gnomo"
u = RandomNumber(1, 6) + 400
If u > 406 Then u = 404
Case Else
u = 1
End Select
End If
'mujer
If UserList(userindex).Genero = "Mujer" Then
Select Case (UserList(userindex).Raza)
Case "Humano"
u = CInt(RandomNumber(1, 7)) + 69
If u > 76 Then u = 74

Case "Elfo"
u = CInt(RandomNumber(1, 7)) + 166
If u > 177 Then u = 172

Case "Elfo Oscuro"
u = CInt(RandomNumber(1, 11)) + 269
If u > 280 Then u = 265

Case "Gnomo"
u = RandomNumber(1, 5) + 469
If u > 474 Then u = 472

Case "Enano"
u = RandomNumber(1, 3) + 369
If u > 372 Then u = 372
Case Else
u = 1

End Select
End If
UserList(userindex).char.Head = u
UserList(userindex).OrigChar.Head = u
Call SendData(toindex, userindex, 0, "||" & "Espero que te guste tu nuevo rostro!!" & FONTTYPE_APU)
Call ChangeUserChar(ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, val(u), UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim)

Exit Sub
           
        Case "/RECOMPENSA"
           'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "Z30")
               Exit Sub
           End If
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(userindex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 4 Then
               Call SendData(SendTarget.toindex, userindex, 0, "Z32")
               Exit Sub
           End If
           If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(userindex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
                Call RecompensaArmadaReal(userindex)
           Else
                If UserList(userindex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
                Call RecompensaCaos(userindex)
           End If
           Exit Sub
           
        Case "/MOTD"
            Call SendMOTD(userindex)
            Exit Sub
            
        Case "/UPTIME"
            tLong = Int(((GetTickCount() And &H7FFFFFFF) - tInicioServer) / 1000)
            tStr = (tLong Mod 60) & " segundos."
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 60) & " minutos, " & tStr
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 24) & " horas, " & tStr
            tLong = Int(tLong / 24)
            tStr = (tLong) & " dias, " & tStr
            Call SendData(SendTarget.toindex, userindex, 0, "||Uptime: " & tStr & FONTTYPE_INFO)
            
            tLong = IntervaloAutoReiniciar
            tStr = (tLong Mod 60) & " segundos."
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 60) & " minutos, " & tStr
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 24) & " horas, " & tStr
            tLong = Int(tLong / 24)
            tStr = (tLong) & " dias, " & tStr
            Call SendData(SendTarget.toindex, userindex, 0, "||Próximo mantenimiento automático: " & tStr & FONTTYPE_INFO)
            Exit Sub
        
        Case "/SALIRPARTY"
            Call mdParty.SalirDeParty(userindex)
            Exit Sub
        
        Case "/CREARPARTY"
            If Not mdParty.PuedeCrearParty(userindex) Then Exit Sub
            Call mdParty.CrearParty(userindex)
            Exit Sub
        Case "/PARTY"
            Call mdParty.SolicitarIngresoAParty(userindex)
            Exit Sub
    End Select
    
  
    If UCase$(Left$(rData, 6)) = "/CMSG " Then
        'clanesnuevo
        rData = Right$(rData, Len(rData) - 6)
        If UserList(userindex).GuildIndex > 0 Then
            Call SendData(SendTarget.ToDiosesYclan, UserList(userindex).GuildIndex, 0, "|+" & UserList(userindex).name & "> " & rData & FONTTYPE_GUILDMSG)
            Call SendData(SendTarget.ToClanArea, userindex, UserList(userindex).pos.Map, "||" & vbYellow & "°< " & rData & " >°" & CStr(UserList(userindex).char.CharIndex))
        End If
        
        Exit Sub
    End If
    
    If UCase$(Left$(rData, 6)) = "/PMSG " Then
        If Len(rData) > 6 Then
            Call mdParty.BroadCastParty(userindex, mid$(rData, 7))
            Call SendData(SendTarget.ToPartyArea, userindex, UserList(userindex).pos.Map, "||" & vbYellow & "°< " & mid$(rData, 7) & " >°" & CStr(UserList(userindex).char.CharIndex))
        End If
        Exit Sub
    End If
    
    If UCase$(Left$(rData, 11)) = "/CENTINELA " Then
        'Evitamos overflow y underflow
        If val(Right$(rData, Len(rData) - 11)) > &H7FFF Or val(Right$(rData, Len(rData) - 11)) < &H8000 Then Exit Sub
        
        tInt = val(Right$(rData, Len(rData) - 11))
        Call CentinelaCheckClave(userindex, tInt)
        Exit Sub
    End If
    
    If UCase$(rData) = "/COLAPINCHADA32" Then
        tStr = modGuilds.m_ListaDeMiembrosOnline(userindex, UserList(userindex).GuildIndex)
        If UserList(userindex).GuildIndex <> 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Compañeros de tu clan conectados: " & tStr & FONTTYPE_GUILDMSG)
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||No pertences a ningún clan." & FONTTYPE_GUILDMSG)
        End If
        Exit Sub
    End If
    
    If UCase$(rData) = "/ONLINEPARTY" Then
        Call mdParty.OnlineParty(userindex)
        Exit Sub
    End If
    
    '[yb]
    If UCase$(Left$(rData, 6)) = "/BMSG " Then
        rData = Right$(rData, Len(rData) - 6)
        If UserList(userindex).flags.PertAlCons = 1 Then
            Call SendData(SendTarget.ToConsejo, userindex, 0, "|| (Consejero) " & UserList(userindex).name & "> " & rData & FONTTYPE_CONSEJO)
        End If
        If UserList(userindex).flags.PertAlConsCaos = 1 Then
            Call SendData(SendTarget.ToConsejoCaos, userindex, 0, "|| (Consejero) " & UserList(userindex).name & "> " & rData & FONTTYPE_CONSEJOCAOS)
        End If
        Exit Sub
    End If
    '[/yb]
    
    If UCase$(Left$(rData, 5)) = "/ROL " Then
        rData = Right$(rData, Len(rData) - 5)
        Call SendData(SendTarget.toindex, 0, 0, "|| " & "Su solicitud ha sido enviada" & FONTTYPE_INFO)
        Call SendData(SendTarget.ToRolesMasters, 0, 0, "|| " & LCase$(UserList(userindex).name) & " PREGUNTA ROL: " & rData & FONTTYPE_GUILDMSG)
        Exit Sub
    End If
    
        'Mensaje del servidor a GMs - Lo ubico aqui para que no se confunda con /GM [Gonzalo]
    If UCase$(Left$(rData, 3)) = "/G " And UserList(userindex).flags.Privilegios > PlayerType.User Then
        rData = Right$(rData, Len(rData) - 3)
        Call LogGM(UserList(userindex).name, "Mensaje a Gms:" & rData, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
        If rData <> "" Then
            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(userindex).name & "> " & rData & "~255~255~255~0~1")
        End If
        Exit Sub
    End If
    
    Select Case UCase$(Left$(rData, 7))
    ' vaya mierda de codigo, solamente sumonea JAJA
        Case "/TORNEO"
            If Hay_Torneo = False Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No hay ningún torneo disponible." & FONTTYPE_INFO)
                Exit Sub
            End If
             If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 
             If UserList(userindex).pos.Map = 67 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en la carcel." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
            If UserList(userindex).Stats.ELV < Torneo_Nivel_Minimo Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Tu nivel es: " & UserList(userindex).Stats.ELV & ".El requerido es: " & Torneo_Nivel_Minimo & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(userindex).Stats.ELV > Torneo_Nivel_Maximo Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Tu nivel es: " & UserList(userindex).Stats.ELV & ".El máximo es: " & Torneo_Nivel_Maximo & FONTTYPE_INFO)
                Exit Sub
            End If
            If Torneo_Inscriptos >= Torneo_Cantidad Then
                Call SendData(SendTarget.toindex, userindex, 0, "||El cupo ya ha sido alcanzado." & FONTTYPE_INFO)
                Exit Sub
            End If
            For i = 1 To 8
                If UCase$(UserList(userindex).Clase) = UCase$(Torneo_Clases_Validas(i)) And Torneo_Clases_Validas2(i) = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Tu clase no es válida en este torneo." & FONTTYPE_INFO)
                Exit Sub
                End If
            Next
            
            Dim NuevaPos As WorldPos
            
            
            'Old, si entras no salis =P
            If Not Torneo.Existe(UserList(userindex).name) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Estás en la lista de espera del torneo. Estás en el puesto nº " & Torneo.Longitud + 1 & FONTTYPE_INFO)
                Call Torneo.Push(rData, UserList(userindex).name)
                
                Call SendData(SendTarget.ToAdmins, 0, 0, "||/TORNEO [" & UserList(userindex).name & "]" & FONTTYPE_INFOBOLD)
                Torneo_Inscriptos = Torneo_Inscriptos + 1
                If Torneo_Inscriptos = Torneo_Cantidad Then
                Call SendData(SendTarget.toall, 0, 0, "||Cupo alcanzado." & FONTTYPE_CELESTE_NEGRITA)
                End If
                If Torneo_SumAuto = 1 Then
                    Dim FuturePos As WorldPos
                    FuturePos.Map = Torneo_Map
                    FuturePos.x = Torneo_X: FuturePos.y = Torneo_Y
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then Call WarpUserChar(userindex, NuevaPos.Map, NuevaPos.x, NuevaPos.y, True)
                End If
            Else
'                Call Torneo.Quitar(UserList(Userindex).Name)
                Call SendData(SendTarget.toindex, userindex, 0, "||Ya estás en la lista de espera del torneo." & FONTTYPE_INFO)
'                Torneo_Inscriptos = Torneo_Inscriptos - 1
'                If Torneo_SumAuto = 1 Then
'                    Call WarpUserChar(Userindex, 1, 50, 50, True)
'                End If
            End If
            Exit Sub
    End Select
    
    
    Select Case UCase$(Left$(rData, 7))
    ' QUE CODIGO MAS WAPO - SATUROS
        Case "/AUTORN"
        If autorneo = False Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No hay ningun torneo automatico! o ya alcanzo su limite." & FONTTYPE_INFO)
        ' como mierda le abia puesto a la variable? autorneo? es q miro el reality y me olvido skjajkasjksakj
        Exit Sub
        End If
         If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 
             If UserList(userindex).pos.Map = 67 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en la carcel." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 
                 If UserList(userindex).pos.Map = 62 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas en el torneo." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
    If UserList(userindex).Stats.ELV < 50 Then
     Call SendData(SendTarget.toindex, userindex, 0, "||Tu nivel minimo para ingresar debe ser de 50! Sube de nivel primero y no me jodas, que con menos de lvl 50 vas a ir a dar pena!." & FONTTYPE_INFO)
     Exit Sub
     End If
     If cupos = 8 Then
      Call SendData(SendTarget.toindex, userindex, 0, "||El limite del torneo ya ha sido alcanzado, espera el próximo torneo!." & FONTTYPE_WARNING)
     Exit Sub
     End If
     If cupos = 0 Then
     cupos = cupos + 1
     tornname1 = UserList(userindex).name
     torn1 = userindex
     Call SendData(SendTarget.toindex, userindex, 0, "||Has ingresado al torneo en el puesto Nº1." & FONTTYPE_WARNING)
     Call WarpUserChar(userindex, 62, 38, 66, True)
     Exit Sub
     End If
     ' saludos pa los del foro eh xD
     If cupos = 1 Then
     cupos = cupos + 1
     tornname2 = UserList(userindex).name
     torn2 = userindex
     Call SendData(SendTarget.toindex, userindex, 0, "||Has ingresado al torneo en el puesto Nº2." & FONTTYPE_WARNING)
     Call WarpUserChar(userindex, 62, 38, 69, True)
     Exit Sub
     End If
     
      If cupos = 2 Then
      cupos = cupos + 1
      tornname3 = UserList(userindex).name
      torn3 = userindex
     Call SendData(SendTarget.toindex, userindex, 0, "||Has ingresado al torneo en el puesto Nº3." & FONTTYPE_WARNING)
     Call WarpUserChar(userindex, 62, 42, 66, True)
     Exit Sub
     End If
     
      If cupos = 3 Then
      cupos = cupos + 1
      tornname4 = UserList(userindex).name
      torn4 = userindex
     Call SendData(SendTarget.toindex, userindex, 0, "||Has ingresado al torneo en el puesto Nº4." & FONTTYPE_WARNING)
     Call WarpUserChar(userindex, 62, 42, 69, True)
     Exit Sub
     End If
     
      If cupos = 4 Then
      cupos = cupos + 1
      tornname5 = UserList(userindex).name
      torn5 = userindex
     Call SendData(SendTarget.toindex, userindex, 0, "||Has ingresado al torneo en el puesto Nº5." & FONTTYPE_WARNING)
     Call WarpUserChar(userindex, 62, 46, 66, True)
     Exit Sub
     End If
     
      If cupos = 5 Then
      cupos = cupos + 1
      tornname6 = UserList(userindex).name
      torn6 = userindex
     Call SendData(SendTarget.toindex, userindex, 0, "||Has ingresado al torneo en el puesto Nº6." & FONTTYPE_WARNING)
     Call WarpUserChar(userindex, 62, 46, 69, True)
     Exit Sub
     End If
     
     If cupos = 6 Then
     cupos = cupos + 1
     tornname7 = UserList(userindex).name
     torn7 = userindex
     Call SendData(SendTarget.toindex, userindex, 0, "||Has ingresado al torneo en el puesto Nº7." & FONTTYPE_WARNING)
     Call WarpUserChar(userindex, 62, 50, 66, True)
     Exit Sub
     End If
     
     If cupos = 7 Then
     cupos = cupos + 1
     tornname8 = UserList(userindex).name
     torn8 = userindex
     Call SendData(SendTarget.toindex, userindex, 0, "||Has ingresado al torneo en el puesto Nº8." & FONTTYPE_WARNING)
     Call WarpUserChar(userindex, 62, 50, 69, True)
     autorneo = False
     Call SendData(SendTarget.toall, 0, 0, "||Torneo> Cupo del torneo Alcanzado" & FONTTYPE_WARNING)
     Call SendData(SendTarget.toall, 0, 0, "||Torneo> Comienza el torneo. Primera batalla: " & tornname1 & " VS " & tornname2 & FONTTYPE_GUILD)
     Call WarpUserChar(torn1, 62, 41, 46, True)
     Call WarpUserChar(torn2, 62, 58, 57, True)
     End If
     
    End Select
    
    
    Select Case UCase$(Left$(rData, 3))
        Case "/GM"
            If Not Ayuda.Existe(UserList(userindex).name) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & FONTTYPE_INFO)
                Call Ayuda.Push(rData, UserList(userindex).name)
            Else
                Call Ayuda.Quitar(UserList(userindex).name)
                Call Ayuda.Push(rData, UserList(userindex).name)
                Call SendData(SendTarget.toindex, userindex, 0, "||Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes." & FONTTYPE_INFO)
            End If
            Exit Sub
    End Select
    
    
    
   
    
    Select Case UCase(Left(rData, 5))
        Case "/_BUG "
            n = FreeFile
            Open App.Path & "\LOGS\BUGs.log" For Append Shared As n
            Print #n,
            Print #n,
            Print #n, "########################################################################"
            Print #n, "########################################################################"
            Print #n, "Usuario:" & UserList(userindex).name & "  Fecha:" & Date & "    Hora:" & Time
            Print #n, "########################################################################"
            Print #n, "BUG:"
            Print #n, Right$(rData, Len(rData) - 5)
            Print #n, "########################################################################"
            Print #n, "########################################################################"
            Print #n,
            Print #n,
            Close #n
            Exit Sub
    
    End Select
    
    Select Case UCase$(Left$(rData, 6))
        Case "/DESC "
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12" & FONTTYPE_INFO)
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 6)
            If Not AsciiValidos(rData) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||La descripcion tiene caracteres invalidos." & FONTTYPE_INFO)
                Exit Sub
            End If
            UserList(userindex).Desc = Trim$(rData)
            Call SendData(SendTarget.toindex, userindex, 0, "||La descripcion a cambiado." & FONTTYPE_INFO)
            Exit Sub
        Case "/VOTO "
                rData = Right$(rData, Len(rData) - 6)
                If Not modGuilds.v_UsuarioVota(userindex, rData, tStr) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||Voto NO contabilizado: " & tStr & FONTTYPE_GUILD)
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||Voto contabilizado." & FONTTYPE_GUILD)
                End If
                Exit Sub
    End Select
    
    If UCase$(Left$(rData, 7)) = "/PENAS " Then
        name = Right$(rData, Len(rData) - 7)
        If name = "" Then Exit Sub
        
        name = Replace(name, "\", "")
        name = Replace(name, "/", "")
        
        If FileExist(CharPath & name & ".chr", vbNormal) Then
            tInt = val(GetVar(CharPath & name & ".chr", "PENAS", "Cant"))
            If tInt = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Sin prontuario.." & FONTTYPE_INFO)
            Else
                While tInt > 0
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & tInt & "- " & GetVar(CharPath & name & ".chr", "PENAS", "P" & tInt) & FONTTYPE_INFO)
                    tInt = tInt - 1
                Wend
            End If
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||Personaje """ & name & """ inexistente." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
    
    
    
    
    
    Select Case UCase$(Left$(rData, 8))
        Case "/PASSWD "
            rData = Right$(rData, Len(rData) - 8)
            If Len(rData) < 6 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||El password debe tener al menos 6 caracteres." & FONTTYPE_INFO)
            Else
                 Call SendData(SendTarget.toindex, userindex, 0, "||El password ha sido cambiado." & FONTTYPE_INFO)
                 UserList(userindex).Password = rData
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 9))
            'Comando /APOSTAR basado en la idea de DarkLight,
            'pero con distinta probabilidad de exito.
        Case "/APOSTAR "
            rData = Right(rData, Len(rData) - 9)
            tLong = CLng(val(rData))
            If tLong > 32000 Then tLong = 32000
            n = tLong
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
            ElseIf UserList(userindex).flags.TargetNPC = 0 Then
                'Se asegura que el target es un npc
                Call SendData(SendTarget.toindex, userindex, 0, "Z30")
            ElseIf Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 10 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z27")
            ElseIf Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "No tengo ningun interes en apostar." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
            ElseIf n < 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "El minimo de apuesta es 1 moneda." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
            ElseIf n > 5000 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "El maximo de apuesta es 5000 monedas." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
            ElseIf UserList(userindex).Stats.GLD < n Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "No tienes esa cantidad." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
            Else
                If RandomNumber(1, 100) <= 47 Then
                    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + n
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "Felicidades! Has ganado " & CStr(n) & " monedas de oro!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                    
                    Apuestas.Perdidas = Apuestas.Perdidas + n
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
                Else
                    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - n
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "Lo siento, has perdido " & CStr(n) & " monedas de oro." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                
                    Apuestas.Ganancias = Apuestas.Ganancias + n
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
                End If
                Apuestas.Jugadas = Apuestas.Jugadas + 1
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
                
                Call EnviarOro(userindex)
            End If
            Exit Sub
    End Select
    
    
    
    Select Case UCase$(Left$(rData, 8))
        Case "/RETIRAR" 'RETIRA ORO EN EL BANCO o te saca de la armada
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(userindex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                      Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "Z30")
                  Exit Sub
             End If
             
             If Npclist(UserList(userindex).flags.TargetNPC).NPCtype = 5 Then
                
                'Se quiere retirar de la armada
                If UserList(userindex).Faccion.ArmadaReal = 1 Then
                    If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 0 Then
                        Call ExpulsarFaccionReal(userindex)
                        Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "º" & "Serás bienvenido a las fuerzas imperiales si deseas regresar." & "º" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                        Debug.Print "||" & vbWhite & "º" & "Serás bienvenido a las fuerzas imperiales si deseas regresar." & "º" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex)
                    Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "º" & "¡¡¡Sal de aquí bufón!!!" & "º" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                    End If
                ElseIf UserList(userindex).Faccion.FuerzasCaos = 1 Then
                    If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 1 Then
                        Call ExpulsarFaccionCaos(userindex)
                        Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "º" & "Ya volverás arrastrandote." & "º" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                    Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "º" & "Sal de aquí maldito criminal" & "º" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                    End If
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "º" & "¡No perteneces a ninguna fuerza!" & "º" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                End If
                Exit Sub
             
             End If
             
             If Len(rData) = 8 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Debes indicar el monto de cuanto quieres retirar" & FONTTYPE_INFO)
                Exit Sub
             End If
             
             rData = Right$(rData, Len(rData) - 9)
             If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Banquero _
             Or UserList(userindex).flags.Muerto = 1 Then Exit Sub
             If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 10 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                  Exit Sub
             End If
             If FileExist(CharPath & UCase$(UserList(userindex).name) & ".chr", vbNormal) = False Then
                  Call SendData(SendTarget.toindex, userindex, 0, "!!El personaje no existe, cree uno nuevo.")
                  CloseSocket (userindex)
                  Exit Sub
             End If
             If val(rData) > 0 And val(rData) <= UserList(userindex).Stats.Banco Then
                  UserList(userindex).Stats.Banco = UserList(userindex).Stats.Banco - val(rData)
                  UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + val(rData)
                  Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(userindex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
             Else
                  Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
             End If
             Call EnviarOro(val(userindex)) 'ak antes habia un senduserstatsbox. lo saque. NicoNZ
             Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 11))
        Case "/DEPOSITAR " 'DEPOSITAR ORO EN EL BANCO
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "Z30")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 10 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                      Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 11)
            If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Banquero _
            Or UserList(userindex).flags.Muerto = 1 Then Exit Sub
            If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 10 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                  Exit Sub
            End If
            If CLng(val(rData)) > 0 And CLng(val(rData)) <= UserList(userindex).Stats.GLD Then
                  UserList(userindex).Stats.Banco = UserList(userindex).Stats.Banco + val(rData)
                  UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - val(rData)
                  Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(userindex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
            Else
                  Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
            End If
            Call EnviarOro(val(userindex))
            Exit Sub
            

        Case "/DENUNCIAR "
        If denuncias = False Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Las denuncias estan desactivadas!" & FONTTYPE_DENUNCIAR)
        Exit Sub
        End If
            If UserList(userindex).flags.YaDenuncio = 2 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Has Alcanzado el Limite de Denuncias Maximo por Log: 2!" & FONTTYPE_DENUNCIAR)
            Exit Sub
            End If
            
            If UserList(userindex).flags.Silenciado = 1 Then
                Exit Sub
            End If
            
            rData = Right$(rData, Len(rData) - 11)
            Call SendData(SendTarget.ToAdmins, 0, 0, "||El PJ " & LCase$(UserList(userindex).name) & " Denuncia: " & rData & FONTTYPE_DENUNCIAR)
            Call SendData(SendTarget.toindex, userindex, 0, "||Tu Denuncia ha sido enviada." & FONTTYPE_DENUNCIAR)
            UserList(userindex).flags.YaDenuncio = UserList(userindex).flags.YaDenuncio + 1
            Exit Sub
            
              
            
            Case "/CERRARCLAN"
If Not UserList(userindex).GuildIndex >= 1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No perteneces a ningún clan." & FONTTYPE_GUILD)
Exit Sub
End If

If UCase$(Guilds(UserList(userindex).GuildIndex).Fundador) <> UCase$(UserList(userindex).name) Then
Call SendData(SendTarget.toindex, userindex, 0, "||No eres líder del clan." & FONTTYPE_GUILD)
Exit Sub
End If

If Guilds(UserList(userindex).GuildIndex).CantidadDeMiembros > 1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Debes hechar a todos los miembros del clan para cerrarlo." & FONTTYPE_GUILD)
Exit Sub
End If

'If UserList(UserIndex).flags.YaCerroClan = 1 Then
'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya has cerrado un clan antes" & FONTTYPE_GUILD)
'Exit Sub
'End If


Call SendData(SendTarget.toall, 0, 0, "||El Clan " & Guilds(UserList(userindex).GuildIndex).GuildName & " cerró." & FONTTYPE_GUILD)

Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Founder", "NADIE")
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "GuildName", Guilds(UserList(userindex).GuildIndex).GuildName & "(CLAN CERRADO)")
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex1", "CLAN CERRADO")
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex2", "CLAN CERRADO")
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex3", "CLAN CERRADO")
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex4", "CLAN CERRADO")
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Leader", "NADIE")


Call Guilds(UserList(userindex).GuildIndex).DesConectarMiembro(userindex)
Call Guilds(UserList(userindex).GuildIndex).ExpulsarMiembro(UserList(userindex).name)
UserList(userindex).GuildIndex = 0
'UserList(UserIndex).flags.YaCerroClan = 1
Call WarpUserChar(userindex, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.y)
Exit Sub

            
            
        Case "/FUNDARCLAN"
            rData = Right$(rData, Len(rData) - 11)
            If Trim$(rData) = vbNullString Then
                Call SendData(SendTarget.toindex, userindex, 0, "|| Para fundar un clan debes especificar la alineación del mismo." & FONTTYPE_GUILD)
                Call SendData(SendTarget.toindex, userindex, 0, "|| Atención, que la misma no podrá cambiar luego, te aconsejamos leer las reglas sobre clanes antes de fundar." & FONTTYPE_GUILD)
                Exit Sub
            Else
                Select Case UCase$(Trim(rData))
                    Case "ARMADA"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_ARMADA
                    Case "MAL"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_LEGION
                    Case "NEUTRO"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_NEUTRO
                    Case "LEGAL"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_CIUDA
                    Case "CRIMINAL"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_CRIMINAL
                    Case Else
                        Call SendData(SendTarget.toindex, userindex, 0, "|| Alineación inválida." & FONTTYPE_GUILD)
                        Exit Sub
                End Select
            End If

            If modGuilds.PuedeFundarUnClan(userindex, UserList(userindex).FundandoGuildAlineacion, tStr) Then
                Call SendData(SendTarget.toindex, userindex, 0, "SHOWFUN")
            Else
                UserList(userindex).FundandoGuildAlineacion = 0
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            End If
           
            Exit Sub
    
    End Select
  
    
     

    
    

    Select Case UCase$(Left$(rData, 12))
        Case "/ECHARPARTY "
            rData = Right$(rData, Len(rData) - 12)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.ExpulsarDeParty(userindex, tInt)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
        Case "/PARTYLIDER "
            rData = Right$(rData, Len(rData) - 12)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.TransformarEnLider(userindex, tInt)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
    
    End Select

    Select Case UCase$(Left$(rData, 13))
        Case "/ACCEPTPARTY "
            rData = Right$(rData, Len(rData) - 13)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.AprobarIngresoAParty(userindex, tInt)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
    
    End Select
    

    Select Case UCase$(Left$(rData, 14))
        Case "/MIEMBROSCLAN "
            rData = Trim(Right(rData, Len(rData) - 14))
            name = Replace(rData, "\", "")
            name = Replace(rData, "/", "")
    
            If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
                Call SendData(SendTarget.toindex, userindex, 0, "|| No existe el clan: " & rData & FONTTYPE_INFO)
                Exit Sub
            End If
            
            tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))
            
            For i = 1 To tInt
                tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)
                'tstr es la victima
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & "<" & rData & ">." & FONTTYPE_INFO)
            Next i
        
            Exit Sub
    End Select
    
    Procesado = False
    
           
End Sub
