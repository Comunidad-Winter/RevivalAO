Attribute VB_Name = "NPCs"

Option Explicit

Sub QuitarMascota(ByVal userindex As Integer, ByVal NpcIndex As Integer)

Dim i As Integer
UserList(userindex).NroMacotas = UserList(userindex).NroMacotas - 1
For i = 1 To MAXMASCOTAS
  If UserList(userindex).MascotasIndex(i) = NpcIndex Then
     UserList(userindex).MascotasIndex(i) = 0
     UserList(userindex).MascotasType(i) = 0
     Exit For
  End If
Next i

End Sub

Sub QuitarMascotaNpc(ByVal Maestro As Integer, ByVal Mascota As Integer)
    Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1
End Sub

Sub MuereNpc(ByVal NpcIndex As Integer, ByVal userindex As Integer)
On Error GoTo errhandler
Dim Npc1 As Integer
Dim Npc1Pos As WorldPos
Npc1 = 906
Npc1Pos.Map = 75
Npc1Pos.x = 48
Npc1Pos.y = 56
Dim Npc2 As Integer
Dim Npc2Pos As WorldPos
Npc2 = 910
Npc2Pos.Map = 77
Npc2Pos.x = 48
Npc2Pos.y = 56

Dim Npc3 As Integer
Dim Npc3Pos As WorldPos
Npc3 = 940
Npc3Pos.Map = 66
Npc3Pos.x = 77
Npc3Pos.y = 23

Dim Npc4 As Integer
Dim Npc4Pos As WorldPos
Npc4 = 941
Npc4Pos.Map = 66
Npc4Pos.x = 77
Npc4Pos.y = 77

Dim Npc5 As Integer
Dim Npc5Pos As WorldPos
Npc5 = 616 ' lemuria
Npc5Pos.Map = 106
Npc5Pos.x = 48
Npc5Pos.y = 56

Dim Npc6 As Integer
Dim Npc6Pos As WorldPos
Npc6 = 617 ' tale
Npc6Pos.Map = 107
Npc6Pos.x = 48
Npc6Pos.y = 56

   Dim MiNPC As npc
   MiNPC = Npclist(NpcIndex)
      '[MaTeO 13]
    If NpcIndex = MazIndex And userindex <> 0 Then
        If RandomNumber(1, 100) <= 60 Then
            Call SendData(SendTarget.toall, 0, 0, "||¡" & UserList(userindex).name & " ha conseguido el Teletransportador! ~255~255~255~1~0" & FONTTYPE_GUERRA)
            Call EraseObj(SendTarget.ToMap, 0, 0, 10000, Npclist(NpcIndex).pos.Map, Npclist(NpcIndex).pos.x, Npclist(NpcIndex).pos.y)
            Dim Objeto As Obj
        
            Objeto.Amount = 1
            Objeto.ObjIndex = 4295
    
            Call MakeObj(SendTarget.ToMap, 0, Npclist(NpcIndex).pos.Map, Objeto, Npclist(NpcIndex).pos.Map, Npclist(NpcIndex).pos.x, Npclist(NpcIndex).pos.y)
        End If
    End If
   '[/MaTeO 13]
   If MiNPC.Numero = 940 Then
        Call SendData(toall, 0, 0, "||Ganan los Angeles la Guerra RevivalAo!" & FONTTYPE_GUERRA)
        Call SendData(toall, 0, 0, "||PREMIO ANGELES: 1.000.000.!" & FONTTYPE_GUERRA)
        Call SendData(SendTarget.toall, userindex, UserList(userindex).pos.Map, "TW44")
    
       Call RespGuerrasAngeles
        Call Ban_Angeles
        End If

  
   If MiNPC.Numero = 941 Then
        Call SendData(toall, 0, 0, "||Ganan los Demonios la Guerra RevivalAo!" & FONTTYPE_GUERRA)
        Call SendData(toall, 0, 0, "||PREMIO DEMONIOS: 1.000.000.!" & FONTTYPE_GUERRA)
        Call SendData(SendTarget.toall, userindex, UserList(userindex).pos.Map, "TW44")
        Call RespGuerrasDemonio
        Call Ban_Demonios
        End If

      If MiNPC.Numero = 906 Then
    AlmacenaDominador = Guilds(UserList(userindex).GuildIndex).GuildName
    HoraUlla = Now
        Call SendData(toall, 0, 0, "||El castillo de Ullathorpe fue conquistado por el clan " & AlmacenaDominador & "." & FONTTYPE_GUILD)
        Call SendData(SendTarget.toall, userindex, UserList(userindex).pos.Map, "TW44")
        Call SpawnNpc(val(Npc1), Npc1Pos, True, False)
        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Ulla", AlmacenaDominador)
        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraUlla", HoraUlla)
        If AlmacenaDominador = AlmacenaDominadornix And AlmacenaDominador = Lemuria And AlmacenaDominador = Tale Then
        Fortaleza = AlmacenaDominador
        HoraForta = Now
        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Fortaleza", AlmacenaDominador)
        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraForta", HoraForta)
        End If
    End If
    
      If MiNPC.Numero = 910 Then
    AlmacenaDominadornix = Guilds(UserList(userindex).GuildIndex).GuildName
    HoraNix = Now
        Call SendData(toall, 0, 0, "||El castillo de Nix fue conquistado por el clan " & AlmacenaDominadornix & "." & FONTTYPE_GUILD)
        Call SendData(SendTarget.toall, userindex, UserList(userindex).pos.Map, "TW44")
        Call SpawnNpc(val(Npc2), Npc2Pos, True, False)
        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Nix", AlmacenaDominadornix)
        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraNix", HoraNix)
          If AlmacenaDominador = AlmacenaDominadornix And AlmacenaDominador = Lemuria And AlmacenaDominador = Tale Then
        Fortaleza = AlmacenaDominadornix
        HoraForta = Now
        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Fortaleza", AlmacenaDominadornix)
        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraForta", HoraForta)
        End If
    End If
    
       If MiNPC.Numero = 616 Then
    Lemuria = Guilds(UserList(userindex).GuildIndex).GuildName
    HoraLemuria = Now
        Call SendData(toall, 0, 0, "||El castillo de Asgard fue conquistado por el clan " & Lemuria & "." & FONTTYPE_GUILD)
        Call SendData(SendTarget.toall, userindex, UserList(userindex).pos.Map, "TW44")
        Call SpawnNpc(val(Npc5), Npc5Pos, True, False)
        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Lemuria", Lemuria)
        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraLemuria", HoraLemuria)
        If AlmacenaDominador = AlmacenaDominadornix And AlmacenaDominador = Lemuria And AlmacenaDominador = Tale Then
        Fortaleza = AlmacenaDominador
        HoraForta = Now
        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Fortaleza", AlmacenaDominador)
        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraForta", HoraForta)
        End If
    End If
    
      If MiNPC.Numero = 617 Then
    Tale = Guilds(UserList(userindex).GuildIndex).GuildName
    HoraTale = Now
        Call SendData(toall, 0, 0, "||El castillo de Tale fue conquistado por el clan " & Tale & "." & FONTTYPE_GUILD)
        Call SendData(SendTarget.toall, userindex, UserList(userindex).pos.Map, "TW44")
        Call SpawnNpc(val(Npc6), Npc6Pos, True, False)
        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Tale", Tale)
        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraTale", HoraTale)
        If AlmacenaDominador = AlmacenaDominadornix And AlmacenaDominador = Lemuria And AlmacenaDominador = Tale Then
        Fortaleza = AlmacenaDominador
        HoraForta = Now
        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "Fortaleza", AlmacenaDominador)
        Call WriteVar(App.Path & "\Dat\Castillos.dat", "CASTILLOS", "HoraForta", HoraForta)
        End If
    End If
    
    
 
   If MiNPC.pos.Map = mapainvo Then MapInfo(mapainvo).criatinv = 0
   'Quitamos el npc
   Call QuitarNPC(NpcIndex)
   
   
    
   If userindex > 0 Then ' Lo mato un usuario?
        If MiNPC.flags.Snd3 > 0 Then Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & MiNPC.flags.Snd3)
        UserList(userindex).flags.TargetNPC = 0
        UserList(userindex).flags.TargetNpcTipo = eNPCType.Comun
        
        'El user que lo mato tiene mascotas?
        If UserList(userindex).NroMacotas > 0 Then
            Dim t As Integer
            For t = 1 To MAXMASCOTAS
                  If UserList(userindex).MascotasIndex(t) > 0 Then
                      If Npclist(UserList(userindex).MascotasIndex(t)).TargetNPC = NpcIndex Then
                              Call FollowAmo(UserList(userindex).MascotasIndex(t))
                      End If
                  End If
            Next t
        End If
        
        '[KEVIN]
        If MiNPC.flags.ExpCount > 0 Then
                If Multexp <> 0 Then MiNPC.flags.ExpCount = MiNPC.flags.ExpCount * Multexp

            If UserList(userindex).PartyIndex > 0 Then
                Call mdParty.ObtenerExito(userindex, MiNPC.flags.ExpCount, MiNPC.pos.Map, MiNPC.pos.x, MiNPC.pos.y)
            Else
                UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + MiNPC.flags.ExpCount
                If UserList(userindex).Stats.Exp > MAXEXP Then _
                    UserList(userindex).Stats.Exp = MAXEXP
                Call SendData(SendTarget.toindex, userindex, 0, "||Has ganado " & MiNPC.flags.ExpCount & " puntos de experiencia." & FONTTYPE_FIGHT)
            Call EnviarExp(userindex)
            End If
            MiNPC.flags.ExpCount = 0
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||No has ganado experiencia al matar la criatura." & FONTTYPE_FIGHT)
        End If
        
        '[/KEVIN]
        Call SendData(SendTarget.toindex, userindex, 0, "||Has matado a la criatura!" & FONTTYPE_FIGHT)
        If UserList(userindex).Stats.NPCsMuertos < 32000 Then _
            UserList(userindex).Stats.NPCsMuertos = UserList(userindex).Stats.NPCsMuertos + 1
        
        If MiNPC.Stats.Alineacion = 0 Then
            If MiNPC.Numero = Guardias Then
                UserList(userindex).Reputacion.NobleRep = 0
                UserList(userindex).Reputacion.PlebeRep = 0
                UserList(userindex).Reputacion.AsesinoRep = UserList(userindex).Reputacion.AsesinoRep + 500
                If UserList(userindex).Reputacion.AsesinoRep > MAXREP Then _
                    UserList(userindex).Reputacion.AsesinoRep = MAXREP
            End If
            If MiNPC.MaestroUser = 0 Then
                UserList(userindex).Reputacion.AsesinoRep = UserList(userindex).Reputacion.AsesinoRep + vlASESINO
                If UserList(userindex).Reputacion.AsesinoRep > MAXREP Then _
                    UserList(userindex).Reputacion.AsesinoRep = MAXREP
            End If
        ElseIf MiNPC.Stats.Alineacion = 1 Then
            UserList(userindex).Reputacion.PlebeRep = UserList(userindex).Reputacion.PlebeRep + vlCAZADOR
            If UserList(userindex).Reputacion.PlebeRep > MAXREP Then _
                UserList(userindex).Reputacion.PlebeRep = MAXREP
        ElseIf MiNPC.Stats.Alineacion = 2 Then
            UserList(userindex).Reputacion.NobleRep = UserList(userindex).Reputacion.NobleRep + vlASESINO / 2
            If UserList(userindex).Reputacion.NobleRep > MAXREP Then _
                UserList(userindex).Reputacion.NobleRep = MAXREP
        ElseIf MiNPC.Stats.Alineacion = 4 Then
            UserList(userindex).Reputacion.PlebeRep = UserList(userindex).Reputacion.PlebeRep + vlCAZADOR
            If UserList(userindex).Reputacion.PlebeRep > MAXREP Then _
                UserList(userindex).Reputacion.PlebeRep = MAXREP
        End If
        If Not Criminal(userindex) And UserList(userindex).Faccion.FuerzasCaos = 1 Then Call ExpulsarFaccionCaos(userindex)
        
        Call CheckUserLevel(userindex)
   End If ' Userindex > 0

   
   If MiNPC.MaestroUser = 0 Then
        'Tiramos el oro
        Call NPCTirarOro(MiNPC, userindex)
        Call EnviarOro(userindex)
        'Tiramos el inventario
        Call NPC_TIRAR_ITEMS(MiNPC)
   End If
   
   'ReSpawn o no
   Call RespawnNPC(MiNPC)
   
Exit Sub

errhandler:
    Call LogError("Error en MuereNpc")
    
End Sub

Sub ResetNpcFlags(ByVal NpcIndex As Integer)
    'Clear the npc's flags
    
    With Npclist(NpcIndex).flags
        .AfectaParalisis = 0
        .AguaValida = 0
        .AttackedBy = ""
        .Attacking = 0
        .BackUp = 0
        .Bendicion = 0
        .Domable = 0
        .Envenenado = 0
        .Faccion = 0
        .Follow = False
        .LanzaSpells = 0
        .GolpeExacto = 0
        .Invisible = 0
        .Maldicion = 0
        .OldHostil = 0
        .OldMovement = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Respawn = 0
        .RespawnOrigPos = 0
        .Snd1 = 0
        .Snd2 = 0
        .Snd3 = 0
        .TierraInvalida = 0
        .UseAINow = False
        .AtacaAPJ = 0
        .AtacaANPC = 0
        .AIAlineacion = e_Alineacion.ninguna
        .AIPersonalidad = e_Personalidad.ninguna
    End With
End Sub

Sub ResetNpcCounters(ByVal NpcIndex As Integer)

Npclist(NpcIndex).Contadores.Paralisis = 0
Npclist(NpcIndex).Contadores.TiempoExistencia = 0

End Sub

Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)

Npclist(NpcIndex).char.Body = 0
Npclist(NpcIndex).char.CascoAnim = 0
Npclist(NpcIndex).char.CharIndex = 0
Npclist(NpcIndex).char.FX = 0
Npclist(NpcIndex).char.Head = 0
Npclist(NpcIndex).char.Heading = 0
Npclist(NpcIndex).char.loops = 0
Npclist(NpcIndex).char.ShieldAnim = 0
Npclist(NpcIndex).char.WeaponAnim = 0


End Sub


Sub ResetNpcCriatures(ByVal NpcIndex As Integer)


Dim j As Integer
For j = 1 To Npclist(NpcIndex).NroCriaturas
    Npclist(NpcIndex).Criaturas(j).NpcIndex = 0
    Npclist(NpcIndex).Criaturas(j).NpcName = ""
Next j

Npclist(NpcIndex).NroCriaturas = 0

End Sub

Sub ResetExpresiones(ByVal NpcIndex As Integer)

Dim j As Integer
For j = 1 To Npclist(NpcIndex).NroExpresiones: Npclist(NpcIndex).Expresiones(j) = "": Next j

Npclist(NpcIndex).NroExpresiones = 0

End Sub


Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)

    Npclist(NpcIndex).Attackable = 0
    Npclist(NpcIndex).CanAttack = 0
    Npclist(NpcIndex).Comercia = 0
    Npclist(NpcIndex).GiveEXP = 0
    Npclist(NpcIndex).GiveGLD = 0
    Npclist(NpcIndex).Hostile = 0
    Npclist(NpcIndex).Inflacion = 0
    Npclist(NpcIndex).InvReSpawn = 0
    Npclist(NpcIndex).level = 0
    
    If Npclist(NpcIndex).MaestroUser > 0 Then Call QuitarMascota(Npclist(NpcIndex).MaestroUser, NpcIndex)
    If Npclist(NpcIndex).MaestroNpc > 0 Then Call QuitarMascotaNpc(Npclist(NpcIndex).MaestroNpc, NpcIndex)
    
    Npclist(NpcIndex).MaestroUser = 0
    Npclist(NpcIndex).MaestroNpc = 0
    
    Npclist(NpcIndex).Mascotas = 0
    Npclist(NpcIndex).Movement = 0
    Npclist(NpcIndex).name = "NPC SIN INICIAR"
    Npclist(NpcIndex).NPCtype = 0
    Npclist(NpcIndex).Numero = 0
    Npclist(NpcIndex).Orig.Map = 0
    Npclist(NpcIndex).Orig.x = 0
    Npclist(NpcIndex).Orig.y = 0
    Npclist(NpcIndex).PoderAtaque = 0
    Npclist(NpcIndex).PoderEvasion = 0
    Npclist(NpcIndex).pos.Map = 0
    Npclist(NpcIndex).pos.x = 0
    Npclist(NpcIndex).pos.y = 0
    Npclist(NpcIndex).SkillDomar = 0
    Npclist(NpcIndex).Target = 0
    Npclist(NpcIndex).TargetNPC = 0
    Npclist(NpcIndex).TipoItems = 0
    Npclist(NpcIndex).Veneno = 0
    Npclist(NpcIndex).Desc = ""
    
    
    Dim j As Integer
    For j = 1 To Npclist(NpcIndex).NroSpells
        Npclist(NpcIndex).Spells(j) = 0
    Next j
    
    Call ResetNpcCharInfo(NpcIndex)
    Call ResetNpcCriatures(NpcIndex)
    Call ResetExpresiones(NpcIndex)

End Sub

Sub QuitarNPC(ByVal NpcIndex As Integer)

On Error GoTo errhandler

    Npclist(NpcIndex).flags.NPCActive = False
    
    If InMapBounds(Npclist(NpcIndex).pos.Map, Npclist(NpcIndex).pos.x, Npclist(NpcIndex).pos.y) Then
        Call EraseNPCChar(SendTarget.ToMap, 0, Npclist(NpcIndex).pos.Map, NpcIndex)
    End If
    
    'Nos aseguramos de que el inventario sea removido...
    'asi los lobos no volveran a tirar armaduras ;))
    Call ResetNpcInv(NpcIndex)
    Call ResetNpcFlags(NpcIndex)
    Call ResetNpcCounters(NpcIndex)
    
    Call ResetNpcMainInfo(NpcIndex)
    
    If NpcIndex = LastNPC Then
        Do Until Npclist(LastNPC).flags.NPCActive
            LastNPC = LastNPC - 1
            If LastNPC < 1 Then Exit Do
        Loop
    End If
        
      
    If NumNPCs <> 0 Then
        NumNPCs = NumNPCs - 1
    End If

Exit Sub

errhandler:
    Npclist(NpcIndex).flags.NPCActive = False
    Call LogError("Error en QuitarNPC")

End Sub

Function TestSpawnTrigger(pos As WorldPos) As Boolean
    
    If LegalPos(pos.Map, pos.x, pos.y) Then
        TestSpawnTrigger = _
        MapData(pos.Map, pos.x, pos.y).trigger <> 3 And _
        MapData(pos.Map, pos.x, pos.y).trigger <> 2 And _
        MapData(pos.Map, pos.x, pos.y).trigger <> 1
    End If

End Function

Sub CrearNPC(NroNPC As Integer, mapa As Integer, OrigPos As WorldPos)
'Call LogTarea("Sub CrearNPC")
'Crea un NPC del tipo NRONPC

Dim pos As WorldPos
Dim newpos As WorldPos
Dim altpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim Iteraciones As Long


Dim Map As Integer
Dim x As Integer
Dim y As Integer

    nIndex = OpenNPC(NroNPC) 'Conseguimos un indice
    
    If nIndex = 0 Then Exit Sub
    
    'Necesita ser respawned en un lugar especifico
    If InMapBounds(OrigPos.Map, OrigPos.x, OrigPos.y) Then
        
        Map = OrigPos.Map
        x = OrigPos.x
        y = OrigPos.y
        Npclist(nIndex).Orig = OrigPos
        Npclist(nIndex).pos = OrigPos
       
    Else
        
        pos.Map = mapa 'mapa
        altpos.Map = mapa
        
        Do While Not PosicionValida
            pos.x = RandomNumber(1, 100)    'Obtenemos posicion al azar en x
            pos.y = RandomNumber(1, 100)    'Obtenemos posicion al azar en y
            
            Call ClosestLegalPos(pos, newpos)  'Nos devuelve la posicion valida mas cercana
            If newpos.x <> 0 Then altpos.x = newpos.x
            If newpos.y <> 0 Then altpos.y = newpos.y     'posicion alternativa (para evitar el anti respawn)
            
            'Si X e Y son iguales a 0 significa que no se encontro posicion valida
            If LegalPosNPC(newpos.Map, newpos.x, newpos.y, Npclist(nIndex).flags.AguaValida) And _
               Not HayPCarea(newpos) And TestSpawnTrigger(newpos) Then
                'Asignamos las nuevas coordenas solo si son validas
                Npclist(nIndex).pos.Map = newpos.Map
                Npclist(nIndex).pos.x = newpos.x
                Npclist(nIndex).pos.y = newpos.y
                PosicionValida = True
            Else
                newpos.x = 0
                newpos.y = 0
            
            End If
                
            'for debug
            Iteraciones = Iteraciones + 1
            If Iteraciones > MAXSPAWNATTEMPS Then
                If altpos.x <> 0 And altpos.y <> 0 Then
                    Map = altpos.Map
                    x = altpos.x
                    y = altpos.y
                    Npclist(nIndex).pos.Map = Map
                    Npclist(nIndex).pos.x = x
                    Npclist(nIndex).pos.y = y
                    Call MakeNPCChar(SendTarget.ToMap, 0, Map, nIndex, Map, x, y)
                    Exit Sub
                Else
                    altpos.x = 50
                    altpos.y = 50
                    Call ClosestLegalPos(altpos, newpos)
                    If newpos.x <> 0 And newpos.y <> 0 Then
                        Npclist(nIndex).pos.Map = newpos.Map
                        Npclist(nIndex).pos.x = newpos.x
                        Npclist(nIndex).pos.y = newpos.y
                        Call MakeNPCChar(SendTarget.ToMap, 0, newpos.Map, nIndex, newpos.Map, newpos.x, newpos.y)
                        Exit Sub
                    Else
                        Call QuitarNPC(nIndex)
                        Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & mapa & " NroNpc:" & NroNPC)
                        Exit Sub
                    End If
                End If
            End If
        Loop
        
        'asignamos las nuevas coordenas
        Map = newpos.Map
        x = Npclist(nIndex).pos.x
        y = Npclist(nIndex).pos.y
    End If
    
    'Crea el NPC
    Call MakeNPCChar(SendTarget.ToMap, 0, Map, nIndex, Map, x, y)

End Sub

Sub MakeNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, NpcIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)
Dim CharIndex As Integer

    If Npclist(NpcIndex).char.CharIndex = 0 Then
        CharIndex = NextOpenCharIndex
        Npclist(NpcIndex).char.CharIndex = CharIndex
        CharList(CharIndex) = NpcIndex
    End If
    
    MapData(Map, x, y).NpcIndex = NpcIndex
    
    If sndRoute = SendTarget.ToMap Then
        Call ArgegarNpc(NpcIndex)
        Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "BC" & Npclist(NpcIndex).char.Body & "," & Npclist(NpcIndex).char.Head & "," & Npclist(NpcIndex).char.Heading & "," & Npclist(NpcIndex).char.CharIndex & "," & x & "," & y & ",0")
    End If

End Sub

Sub ChangeNPCChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal NpcIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As eHeading)

If NpcIndex > 0 Then
    Npclist(NpcIndex).char.Body = Body
    Npclist(NpcIndex).char.Head = Head
    Npclist(NpcIndex).char.Heading = Heading
    If sndRoute = SendTarget.ToMap Then
        Call SendToNpcArea(NpcIndex, "CP" & Npclist(NpcIndex).char.CharIndex & "," & Body & "," & Head & "," & Heading)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "CP" & Npclist(NpcIndex).char.CharIndex & "," & Body & "," & Head & "," & Heading)
    End If
End If

End Sub

Sub EraseNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, ByVal NpcIndex As Integer)

If Npclist(NpcIndex).char.CharIndex <> 0 Then CharList(Npclist(NpcIndex).char.CharIndex) = 0

If Npclist(NpcIndex).char.CharIndex = LastChar Then
    Do Until CharList(LastChar) > 0
        LastChar = LastChar - 1
        If LastChar <= 1 Then Exit Do
    Loop
End If

'Quitamos del mapa
MapData(Npclist(NpcIndex).pos.Map, Npclist(NpcIndex).pos.x, Npclist(NpcIndex).pos.y).NpcIndex = 0

Dim code As String
code = str(Npclist(NpcIndex).char.CharIndex)
If sndRoute = SendTarget.ToMap Then
    
       ' Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "BP" & code)
    Call SendToNpcArea(NpcIndex, "BP" & code)
Else
    Call SendData(sndRoute, sndIndex, sndMap, "BP" & code)
End If

'Update la lista npc
Npclist(NpcIndex).char.CharIndex = 0


'update NumChars
NumChars = NumChars - 1


End Sub

Sub MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte)

On Error GoTo errh
    Dim nPos As WorldPos
    nPos = Npclist(NpcIndex).pos
    Call HeadtoPos(nHeading, nPos)
    
    'Es mascota ????
    If Npclist(NpcIndex).MaestroUser > 0 Then
        ' es una posicion legal
        If LegalPos(Npclist(NpcIndex).pos.Map, nPos.x, nPos.y, Npclist(NpcIndex).flags.AguaValida = 1) Then
        
            If Npclist(NpcIndex).flags.AguaValida = 0 And HayAgua(Npclist(NpcIndex).pos.Map, nPos.x, nPos.y) Then Exit Sub
            If Npclist(NpcIndex).flags.TierraInvalida = 1 And Not HayAgua(Npclist(NpcIndex).pos.Map, nPos.x, nPos.y) Then Exit Sub
            
#If SeguridadAlkon Then
            Call SendToNpcArea(NpcIndex, "*" & Encriptacion.MoveNPCCrypt(NpcIndex, nPos.x, nPos.y))
#Else
            Call SendToNpcArea(NpcIndex, "*" & Npclist(NpcIndex).char.CharIndex & "," & nPos.x & "," & nPos.y)
#End If
            
            'Update map and user pos
            MapData(Npclist(NpcIndex).pos.Map, Npclist(NpcIndex).pos.x, Npclist(NpcIndex).pos.y).NpcIndex = 0
            Npclist(NpcIndex).pos = nPos
            Npclist(NpcIndex).char.Heading = nHeading
            MapData(Npclist(NpcIndex).pos.Map, Npclist(NpcIndex).pos.x, Npclist(NpcIndex).pos.y).NpcIndex = NpcIndex
            Call CheckUpdateNeededNpc(NpcIndex, nHeading)
        End If
Else ' No es mascota
        ' Controlamos que la posicion sea legal, los npc que
        ' no son mascotas tienen mas restricciones de movimiento.
        If LegalPosNPC(Npclist(NpcIndex).pos.Map, nPos.x, nPos.y, Npclist(NpcIndex).flags.AguaValida) Then
            
            If Npclist(NpcIndex).flags.AguaValida = 0 And HayAgua(Npclist(NpcIndex).pos.Map, nPos.x, nPos.y) Then Exit Sub
            If Npclist(NpcIndex).flags.TierraInvalida = 1 And Not HayAgua(Npclist(NpcIndex).pos.Map, nPos.x, nPos.y) Then Exit Sub
            
            '[Alejo-18-5]
            'server
#If SeguridadAlkon Then
            Call SendToNpcArea(NpcIndex, "*" & Encriptacion.MoveNPCCrypt(NpcIndex, nPos.x, nPos.y))
#Else
            Call SendToNpcArea(NpcIndex, "*" & Npclist(NpcIndex).char.CharIndex & "," & nPos.x & "," & nPos.y)
#End If
            
            'Update map and user pos
            MapData(Npclist(NpcIndex).pos.Map, Npclist(NpcIndex).pos.x, Npclist(NpcIndex).pos.y).NpcIndex = 0
            Npclist(NpcIndex).pos = nPos
            Npclist(NpcIndex).char.Heading = nHeading
            MapData(Npclist(NpcIndex).pos.Map, Npclist(NpcIndex).pos.x, Npclist(NpcIndex).pos.y).NpcIndex = NpcIndex
            
            Call CheckUpdateNeededNpc(NpcIndex, nHeading)
        
        Else
            If Npclist(NpcIndex).Movement = TipoAI.NpcPathfinding Then
                'Someone has blocked the npc's way, we must to seek a new path!
                Npclist(NpcIndex).PFINFO.PathLenght = 0
            End If
        
        End If
    End If

Exit Sub

errh:
    LogError ("Error en move npc " & NpcIndex)


End Sub

Function NextOpenNPC() As Integer
'Call LogTarea("Sub NextOpenNPC")

On Error GoTo errhandler

Dim LoopC As Integer
  
For LoopC = 1 To MAXNPCS + 1
    If LoopC > MAXNPCS Then Exit For
    If Not Npclist(LoopC).flags.NPCActive Then Exit For
Next LoopC
  
NextOpenNPC = LoopC


Exit Function
errhandler:
    Call LogError("Error en NextOpenNPC")
End Function

Sub NpcEnvenenarUser(ByVal userindex As Integer)

Dim n As Integer
n = RandomNumber(1, 100)
If n < 30 Then
    UserList(userindex).flags.Envenenado = 1
    Call SendData(SendTarget.toindex, userindex, 0, "||¡¡La criatura te ha envenenado!!" & FONTTYPE_FIGHT)
End If

End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, pos As WorldPos, ByVal FX As Boolean, ByVal Respawn As Boolean) As Integer
'Crea un NPC del tipo Npcindex

Dim newpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean


Dim Map As Integer
Dim x As Integer
Dim y As Integer
Dim it As Integer

nIndex = OpenNPC(NpcIndex, Respawn)   'Conseguimos un indice

it = 0

If nIndex > MAXNPCS Then
    SpawnNpc = 0
    Exit Function
End If

Do While Not PosicionValida
        
        Call ClosestLegalPos(pos, newpos)  'Nos devuelve la posicion valida mas cercana
        'Si X e Y son iguales a 0 significa que no se encontro posicion valida
        If Npclist(nIndex).flags.TierraInvalida Then
            If LegalPos(newpos.Map, newpos.x, newpos.y, True) Then _
                PosicionValida = True
        Else
            If LegalPos(newpos.Map, newpos.x, newpos.y, False) Or LegalPos(newpos.Map, newpos.x, newpos.y, Npclist(nIndex).flags.AguaValida) Then _
                PosicionValida = True
        End If
        
        If PosicionValida Then
            'Asignamos las nuevas coordenas solo si son validas
            Npclist(nIndex).pos.Map = newpos.Map
            Npclist(nIndex).pos.x = newpos.x
            Npclist(nIndex).pos.y = newpos.y
        Else
            newpos.x = 0
            newpos.y = 0
        End If
        
        it = it + 1
        
        If it > MAXSPAWNATTEMPS Then
            Call QuitarNPC(nIndex)
            SpawnNpc = 0
            Call LogError("Mas de " & MAXSPAWNATTEMPS & " iteraciones en SpawnNpc Mapa:" & pos.Map & " Index:" & NpcIndex)
            Exit Function
        End If
Loop

'asignamos las nuevas coordenas
Map = newpos.Map
x = Npclist(nIndex).pos.x
y = Npclist(nIndex).pos.y

'Crea el NPC
Call MakeNPCChar(SendTarget.ToMap, 0, Map, nIndex, Map, x, y)

If FX Then
    Call SendData(SendTarget.ToNPCArea, nIndex, Map, "TW" & SND_WARP)
    Call SendData(SendTarget.ToNPCArea, nIndex, Map, "CFX" & Npclist(nIndex).char.CharIndex & "," & FXIDs.FXWARP & "," & 0)
End If

SpawnNpc = nIndex

End Function

Sub RespawnNPC(MiNPC As npc)

If (MiNPC.flags.Respawn = 0) Then Call CrearNPC(MiNPC.Numero, MiNPC.pos.Map, MiNPC.Orig)

End Sub

'Devuelve el nro de enemigos que hay en el Mapa Map
Function NPCHostiles(ByVal Map As Integer) As Integer

Dim NpcIndex As Integer
Dim cont As Integer

'Contador
cont = 0
For NpcIndex = 1 To LastNPC

    '¿esta vivo?
    If Npclist(NpcIndex).flags.NPCActive _
       And Npclist(NpcIndex).pos.Map = Map _
       And Npclist(NpcIndex).Hostile = 1 And _
       Npclist(NpcIndex).Stats.Alineacion = 2 Then
            cont = cont + 1
           
    End If
    
Next NpcIndex

NPCHostiles = cont

End Function

Sub NPCTirarOro(MiNPC As npc, userindex As Integer)

'SI EL NPC TIENE ORO LO TIRAMOS
If MiNPC.GiveGLD > 0 Then
UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + MiNPC.GiveGLD
Call SendData(SendTarget.toindex, userindex, 0, "||Has Ganado " & MiNPC.GiveGLD & " Monedas de Oro." & FONTTYPE_ORO)
End If

End Sub

Function OpenNPC(ByVal NpcNumber As Integer, Optional ByVal Respawn = True) As Integer
On Error GoTo Err
      '###################################################
      '#               ATENCION PELIGRO                  #
      '###################################################
      '
      '    ¡¡¡¡ NO USAR GetVar PARA LEER LOS NPCS !!!!
      '
      'El que ose desafiar esta LEY, se las tendrá que ver
      'con migo. Para leer los NPCS se deberá usar la
      'nueva clase clsLeerInis.
      '
      'Alejo
      '
      '###################################################

      Dim NpcIndex As Integer
      Dim npcfile As String
      Dim Leer As clsIniReader

10    If NpcNumber > 499 Then
              'NpcFile = DatPath & "NPCs-HOSTILES.dat"
20            Set Leer = LeerNPCsHostiles
30    Else
              'NpcFile = DatPath & "NPCs.dat"
40            Set Leer = LeerNPCs
50    End If

60    NpcIndex = NextOpenNPC

70    If NpcIndex > MAXNPCS Then 'Limite de npcs
80        OpenNPC = NpcIndex
90        Exit Function
100   End If

110   Npclist(NpcIndex).Numero = NpcNumber
120   Npclist(NpcIndex).name = Leer.GetValue("NPC" & NpcNumber, "Name")
130   Npclist(NpcIndex).Desc = Leer.GetValue("NPC" & NpcNumber, "Desc")

140   Npclist(NpcIndex).Movement = val(Leer.GetValue("NPC" & NpcNumber, "Movement"))
150   Npclist(NpcIndex).flags.OldMovement = Npclist(NpcIndex).Movement

160   Npclist(NpcIndex).flags.AguaValida = val(Leer.GetValue("NPC" & NpcNumber, "AguaValida"))
170   Npclist(NpcIndex).flags.TierraInvalida = val(Leer.GetValue("NPC" & NpcNumber, "TierraInValida"))
180   Npclist(NpcIndex).flags.Faccion = val(Leer.GetValue("NPC" & NpcNumber, "Faccion"))

190   Npclist(NpcIndex).NPCtype = val(Leer.GetValue("NPC" & NpcNumber, "NpcType"))

200   Npclist(NpcIndex).char.Body = val(Leer.GetValue("NPC" & NpcNumber, "Body"))
210   Npclist(NpcIndex).char.Head = val(Leer.GetValue("NPC" & NpcNumber, "Head"))
220   Npclist(NpcIndex).char.Heading = val(Leer.GetValue("NPC" & NpcNumber, "Heading"))

230   Npclist(NpcIndex).Attackable = val(Leer.GetValue("NPC" & NpcNumber, "Attackable"))
240   Npclist(NpcIndex).Comercia = val(Leer.GetValue("NPC" & NpcNumber, "Comercia"))
250   Npclist(NpcIndex).Hostile = val(Leer.GetValue("NPC" & NpcNumber, "Hostile"))
260   Npclist(NpcIndex).flags.OldHostil = Npclist(NpcIndex).Hostile

270   Npclist(NpcIndex).GiveEXP = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXP"))

      'Npclist(NpcIndex).flags.ExpDada = Npclist(NpcIndex).GiveEXP
280   Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).GiveEXP

290   Npclist(NpcIndex).Veneno = val(Leer.GetValue("NPC" & NpcNumber, "Veneno"))

300   Npclist(NpcIndex).flags.Domable = val(Leer.GetValue("NPC" & NpcNumber, "Domable"))


310   Npclist(NpcIndex).GiveGLD = val(Leer.GetValue("NPC" & NpcNumber, "GiveGLD"))

320   Npclist(NpcIndex).PoderAtaque = val(Leer.GetValue("NPC" & NpcNumber, "PoderAtaque"))
330   Npclist(NpcIndex).PoderEvasion = val(Leer.GetValue("NPC" & NpcNumber, "PoderEvasion"))

340   Npclist(NpcIndex).InvReSpawn = val(Leer.GetValue("NPC" & NpcNumber, "InvReSpawn"))


350   Npclist(NpcIndex).Stats.MaxHP = val(Leer.GetValue("NPC" & NpcNumber, "MaxHP"))
360   Npclist(NpcIndex).Stats.MinHP = val(Leer.GetValue("NPC" & NpcNumber, "MinHP"))
370   Npclist(NpcIndex).Stats.MaxHIT = val(Leer.GetValue("NPC" & NpcNumber, "MaxHIT"))
380   Npclist(NpcIndex).Stats.MinHIT = val(Leer.GetValue("NPC" & NpcNumber, "MinHIT"))
390   Npclist(NpcIndex).Stats.def = val(Leer.GetValue("NPC" & NpcNumber, "DEF"))
400   Npclist(NpcIndex).Stats.Alineacion = val(Leer.GetValue("NPC" & NpcNumber, "Alineacion"))


      Dim LoopC As Integer
      Dim ln As String
410   Npclist(NpcIndex).Invent.NroItems = val(Leer.GetValue("NPC" & NpcNumber, "NROITEMS"))
      '[MaTeO 5]
420   For LoopC = 1 To MAX_INVENTORY_SLOTS_NPC
      '[/MaTeO 5]


430       ln = Leer.GetValue("NPC" & NpcNumber, "Obj" & LoopC)
440       Npclist(NpcIndex).Invent.Object(LoopC).ProbTirar = val(ReadField(3, ln, 45))
450       Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
460       Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
470   Next LoopC

480   Npclist(NpcIndex).flags.LanzaSpells = val(Leer.GetValue("NPC" & NpcNumber, "LanzaSpells"))
490   If Npclist(NpcIndex).flags.LanzaSpells > 0 Then ReDim Npclist(NpcIndex).Spells(1 To Npclist(NpcIndex).flags.LanzaSpells)
500   For LoopC = 1 To Npclist(NpcIndex).flags.LanzaSpells
510       Npclist(NpcIndex).Spells(LoopC) = val(Leer.GetValue("NPC" & NpcNumber, "Sp" & LoopC))
520   Next LoopC


530   If Npclist(NpcIndex).NPCtype = eNPCType.Entrenador Then
540       Npclist(NpcIndex).NroCriaturas = val(Leer.GetValue("NPC" & NpcNumber, "NroCriaturas"))
550       ReDim Npclist(NpcIndex).Criaturas(1 To Npclist(NpcIndex).NroCriaturas) As tCriaturasEntrenador
560       For LoopC = 1 To Npclist(NpcIndex).NroCriaturas
570           Npclist(NpcIndex).Criaturas(LoopC).NpcIndex = Leer.GetValue("NPC" & NpcNumber, "CI" & LoopC)
580           Npclist(NpcIndex).Criaturas(LoopC).NpcName = Leer.GetValue("NPC" & NpcNumber, "CN" & LoopC)
590       Next LoopC
600   End If


610   Npclist(NpcIndex).Inflacion = val(Leer.GetValue("NPC" & NpcNumber, "Inflacion"))

620   Npclist(NpcIndex).flags.NPCActive = True
630   Npclist(NpcIndex).flags.UseAINow = False

640   If Respawn Then
650       Npclist(NpcIndex).flags.Respawn = val(Leer.GetValue("NPC" & NpcNumber, "ReSpawn"))
660   Else
670       Npclist(NpcIndex).flags.Respawn = 1
680   End If

690   Npclist(NpcIndex).flags.BackUp = val(Leer.GetValue("NPC" & NpcNumber, "BackUp"))
700   Npclist(NpcIndex).flags.RespawnOrigPos = val(Leer.GetValue("NPC" & NpcNumber, "OrigPos"))
710   Npclist(NpcIndex).flags.AfectaParalisis = val(Leer.GetValue("NPC" & NpcNumber, "AfectaParalisis"))
720   Npclist(NpcIndex).flags.GolpeExacto = val(Leer.GetValue("NPC" & NpcNumber, "GolpeExacto"))


730   Npclist(NpcIndex).flags.Snd1 = val(Leer.GetValue("NPC" & NpcNumber, "Snd1"))
740   Npclist(NpcIndex).flags.Snd2 = val(Leer.GetValue("NPC" & NpcNumber, "Snd2"))
750   Npclist(NpcIndex).flags.Snd3 = val(Leer.GetValue("NPC" & NpcNumber, "Snd3"))

      '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

      Dim aux As String
760   aux = Leer.GetValue("NPC" & NpcNumber, "NROEXP")
770   If aux = "" Then
780       Npclist(NpcIndex).NroExpresiones = 0
790   Else
800       Npclist(NpcIndex).NroExpresiones = val(aux)
810       ReDim Npclist(NpcIndex).Expresiones(1 To Npclist(NpcIndex).NroExpresiones) As String
820       For LoopC = 1 To Npclist(NpcIndex).NroExpresiones
830           Npclist(NpcIndex).Expresiones(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Exp" & LoopC)
840       Next LoopC
850   End If

      '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

      'Tipo de items con los que comercia
860   Npclist(NpcIndex).TipoItems = val(Leer.GetValue("NPC" & NpcNumber, "TipoItems"))

      'Update contadores de NPCs
870   If NpcIndex > LastNPC Then LastNPC = NpcIndex
880   NumNPCs = NumNPCs + 1


      'Devuelve el nuevo Indice
890   OpenNPC = NpcIndex
Exit Function
Err:
Debug.Print "OpenNPC linea: " & Erl()
End Function


Sub EnviarListaCriaturas(ByVal userindex As Integer, ByVal NpcIndex)
  Dim SD As String
  Dim k As Integer
  SD = SD & Npclist(NpcIndex).NroCriaturas & ","
  For k = 1 To Npclist(NpcIndex).NroCriaturas
        SD = SD & Npclist(NpcIndex).Criaturas(k).NpcName & ","
  Next k
  SD = "LSTCRI" & SD
  Call SendData(SendTarget.toindex, userindex, 0, SD)
End Sub


Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserName As String)

If Npclist(NpcIndex).flags.Follow Then
  Npclist(NpcIndex).flags.AttackedBy = ""
  Npclist(NpcIndex).flags.Follow = False
  Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
  Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
Else
  Npclist(NpcIndex).flags.AttackedBy = UserName
  Npclist(NpcIndex).flags.Follow = True
  Npclist(NpcIndex).Movement = 4 'follow
  Npclist(NpcIndex).Hostile = 0
End If

End Sub

Sub FollowAmo(ByVal NpcIndex As Integer)

  Npclist(NpcIndex).flags.Follow = True
  Npclist(NpcIndex).Movement = TipoAI.SigueAmo 'follow
  Npclist(NpcIndex).Hostile = 0
  Npclist(NpcIndex).Target = 0
  Npclist(NpcIndex).TargetNPC = 0

End Sub

