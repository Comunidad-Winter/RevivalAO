Attribute VB_Name = "modHechizos"

'Pablo Ignacio Márquez

Option Explicit

Public Const HELEMENTAL_FUEGO As Integer = 26
Public Const HELEMENTAL_TIERRA As Integer = 28
Public Const SUPERANILLO As Integer = 700

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal userindex As Integer, ByVal Spell As Integer)

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
If UserList(userindex).flags.Invisible = 1 Or UserList(userindex).flags.Oculto = 1 Then Exit Sub

Npclist(NpcIndex).CanAttack = 0
Dim daño As Integer

If Hechizos(Spell).SubeHP = 1 Then

    daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & Hechizos(Spell).WAV)
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP + daño
    If UserList(userindex).Stats.MinHP > UserList(userindex).Stats.MaxHP Then UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
    
    Call SendData(SendTarget.toIndex, userindex, 0, "||" & Npclist(NpcIndex).name & " te ha quitado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
    Call EnviarHP(val(userindex))

ElseIf Hechizos(Spell).SubeHP = 2 Then
    
    If UserList(userindex).flags.Privilegios = PlayerType.User Then
    
        daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        
        If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then
            daño = daño - RandomNumber(ObjData(UserList(userindex).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(userindex).Invent.CascoEqpObjIndex).DefensaMagicaMax)
        End If
        
        If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
            daño = daño - RandomNumber(ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).DefensaMagicaMax)
        End If
        
        If UserList(userindex).Invent.HerramientaEqpObjIndex > 0 Then
            daño = daño - RandomNumber(ObjData(UserList(userindex).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList(userindex).Invent.HerramientaEqpObjIndex).DefensaMagicaMax)
        End If
        
        If daño < 0 Then daño = 0
        
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & Hechizos(Spell).WAV)
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
    
        UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - daño
        
        Call SendData(SendTarget.toIndex, userindex, 0, "||" & Npclist(NpcIndex).name & " te ha quitado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
        Call EnviarHP(val(userindex))
        
        'Muere
        If UserList(userindex).Stats.MinHP < 1 Then
            UserList(userindex).Stats.MinHP = 0
            If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                RestarCriminalidad (userindex)
            End If
            Call UserDie(userindex)
            '[Barrin 1-12-03]
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call ContarMuerte(userindex, Npclist(NpcIndex).MaestroUser)
                Call ActStats(userindex, Npclist(NpcIndex).MaestroUser)
            End If
            '[/Barrin]
        End If
    
    End If
    
End If

If Hechizos(Spell).Paraliza = 1 Then
     If UserList(userindex).flags.Paralizado = 0 Then
          Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & Hechizos(Spell).WAV)
          Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
          
            If UserList(userindex).Invent.HerramientaEqpObjIndex = SUPERANILLO Then
                Call SendData(SendTarget.toIndex, userindex, 0, "|| Tu anillo rechaza los efectos del hechizo." & FONTTYPE_FIGHT)
                Exit Sub
            End If
          UserList(userindex).flags.Paralizado = 1
          UserList(userindex).Counters.Paralisis = IntervaloParalizado
            Call SendData(SendTarget.toIndex, userindex, 0, "PARADOW")
            Call SendData(SendTarget.toIndex, userindex, 0, "PU" & UserList(userindex).pos.x & "," & UserList(userindex).pos.Y)
     End If
     
     
End If


End Sub


Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal Spell As Integer)
'solo hechizos ofensivos!

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
Npclist(NpcIndex).CanAttack = 0

Dim daño As Integer

If Hechizos(Spell).SubeHP = 2 Then
    
        daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        Call SendData(SendTarget.ToNPCArea, TargetNPC, Npclist(TargetNPC).pos.Map, "TW" & Hechizos(Spell).WAV)
        Call SendData(SendTarget.ToNPCArea, TargetNPC, Npclist(TargetNPC).pos.Map, "CFX" & Npclist(TargetNPC).char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
        
        Npclist(TargetNPC).Stats.MinHP = Npclist(TargetNPC).Stats.MinHP - daño
        
        'Muere
        If Npclist(TargetNPC).Stats.MinHP < 1 Then
            Npclist(TargetNPC).Stats.MinHP = 0
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
            Else
                Call MuereNpc(TargetNPC, 0)
            End If
        End If
    
End If
    
End Sub



Function TieneHechizo(ByVal i As Integer, ByVal userindex As Integer) As Boolean

On Error GoTo errhandler
    
    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(userindex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

Exit Function
errhandler:

End Function

Sub AgregarHechizo(ByVal userindex As Integer, ByVal Slot As Integer)
Dim hIndex As Integer
Dim j As Integer
hIndex = ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex).HechizoIndex

If Not TieneHechizo(hIndex, userindex) Then
    'Buscamos un slot vacio
    For j = 1 To MAXUSERHECHIZOS
        If UserList(userindex).Stats.UserHechizos(j) = 0 Then Exit For
    Next j
        
    If UserList(userindex).Stats.UserHechizos(j) <> 0 Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||No tenes espacio para mas hechizos." & FONTTYPE_INFO)
    Else
        UserList(userindex).Stats.UserHechizos(j) = hIndex
        Call UpdateUserHechizos(False, userindex, CByte(j))
        'Quitamos del inv el item
        Call QuitarUserInvItem(userindex, CByte(Slot), 1)
    End If
Else
    Call SendData(SendTarget.toIndex, userindex, 0, "||Ya tenes ese hechizo." & FONTTYPE_INFO)
End If

End Sub
            
Sub DecirPalabrasMagicas(ByVal s As String, ByVal userindex As Integer)
On Error Resume Next

    Dim ind As String
    ind = UserList(userindex).char.CharIndex
    If Criminal(userindex) Then
     Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbRed & "°" & s & "°" & ind)
      Exit Sub
     End If
    If Not Criminal(userindex) Then
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbCyan & "°" & s & "°" & ind)
     Exit Sub
    End If
   
End Sub

Function PuedeLanzar(ByVal userindex As Integer, ByVal HechizoIndex As Integer) As Boolean

If UserList(userindex).flags.Muerto = 0 Then
    Dim wp2 As WorldPos
    wp2.Map = UserList(userindex).flags.TargetMap
    wp2.x = UserList(userindex).flags.TargetX
    wp2.Y = UserList(userindex).flags.TargetY
    
    If Hechizos(HechizoIndex).NeedStaff > 0 Then
        If UCase$(UserList(userindex).Clase) = "MAGO" Then
            If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                    Call SendData(SendTarget.toIndex, userindex, 0, "||Tu Báculo no es lo suficientemente poderoso para que puedas lanzar el conjuro." & FONTTYPE_INFO)
                    PuedeLanzar = False
                    Exit Function
                End If
            Else
                Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes lanzar este conjuro sin la ayuda de un báculo." & FONTTYPE_INFO)
                PuedeLanzar = False
                Exit Function
            End If
        End If
    End If
          
             If Hechizos(HechizoIndex).NeedStaff > 0 Then
        If UCase$(UserList(userindex).Clase) = "CLERIGO" Then
            If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                    Call SendData(SendTarget.toIndex, userindex, 0, "||Tu espada no es lo suficientemente fuerte para lanzar este hechizo." & FONTTYPE_INFO)
                    PuedeLanzar = False
                    Exit Function
                End If
            Else
                Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes lanzar este conjuro sin la ayuda de una Espada Argentum." & FONTTYPE_INFO)
                PuedeLanzar = False
                Exit Function
            End If
        End If
    End If
    If UserList(userindex).Stats.MinMAN >= Hechizos(HechizoIndex).ManaRequerido Then
        If UserList(userindex).Stats.UserSkills(eSkill.Magia) >= Hechizos(HechizoIndex).MinSkill Then
            If UserList(userindex).Stats.MinSta >= Hechizos(HechizoIndex).StaRequerido Then
                PuedeLanzar = True
            Else
                Call SendData(SendTarget.toIndex, userindex, 0, "Z1")
                PuedeLanzar = False
            End If
                
        Else
            Call SendData(SendTarget.toIndex, userindex, 0, "Z2")
            PuedeLanzar = False
        End If
    Else
            Call SendData(SendTarget.toIndex, userindex, 0, "Z3")
            PuedeLanzar = False
    End If
Else
   Call SendData(SendTarget.toIndex, userindex, 0, "Z4")
   PuedeLanzar = False
End If

End Function

Sub HechizoTerrenoEstado(ByVal userindex As Integer, ByRef b As Boolean)
Dim PosCasteadaX As Integer
Dim PosCasteadaY As Integer
Dim PosCasteadaM As Integer
Dim h As Integer
Dim TempX As Integer
Dim TempY As Integer


    PosCasteadaX = UserList(userindex).flags.TargetX
    PosCasteadaY = UserList(userindex).flags.TargetY
    PosCasteadaM = UserList(userindex).flags.TargetMap
    
    h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
    
    If Hechizos(h).RemueveInvisibilidadParcial = 1 Then
        b = True
        For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
            For TempY = PosCasteadaY - 8 To PosCasteadaY + 8
                If InMapBounds(PosCasteadaM, TempX, TempY) Then
                    If MapData(PosCasteadaM, TempX, TempY).userindex > 0 Then
                        'hay un user
                        If UserList(MapData(PosCasteadaM, TempX, TempY).userindex).flags.Invisible = 1 Or UserList(MapData(PosCasteadaM, TempX, TempY).userindex).flags.Oculto = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).userindex).flags.Privilegios = PlayerType.User Then
                            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(MapData(PosCasteadaM, TempX, TempY).userindex).char.CharIndex & "," & Hechizos(h).FXgrh & "," & Hechizos(h).loops)
                        End If
                    End If
                End If
            Next TempY
        Next TempX
    
        Call InfoHechizo(userindex)
    End If

End Sub

Sub HechizoInvocacion(ByVal userindex As Integer, ByRef b As Boolean)

If UserList(userindex).NroMacotas >= MAXMASCOTAS Then Exit Sub

'No permitimos se invoquen criaturas en zonas seguras
If MapInfo(UserList(userindex).pos.Map).Pk = False Or MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).trigger = eTrigger.ZONASEGURA Then
    Call SendData(SendTarget.toIndex, userindex, 0, "Z5")
    Exit Sub
End If
If UserList(userindex).pos.Map = 66 Then
 Call SendData(SendTarget.toIndex, userindex, 0, "||Aqui no puedes invocar mascotas!." & FONTTYPE_INFO)
 Exit Sub
 End If
If UserList(userindex).pos.Map = 75 Then
 Call SendData(SendTarget.toIndex, userindex, 0, "||Aqui no puedes invocar mascotas!." & FONTTYPE_INFO)
 Exit Sub
 End If
 If UserList(userindex).pos.Map = 77 Then
 Call SendData(SendTarget.toIndex, userindex, 0, "||Aqui no puedes invocar mascotas!." & FONTTYPE_INFO)
 Exit Sub
 End If
 If UserList(userindex).pos.Map = 86 Then
 Call SendData(SendTarget.toIndex, userindex, 0, "||Aqui no puedes invocar mascotas!." & FONTTYPE_INFO)
 Exit Sub
 End If
  If UserList(userindex).pos.Map = 107 Then
 Call SendData(SendTarget.toIndex, userindex, 0, "||Aqui no puedes invocar mascotas!." & FONTTYPE_INFO)
 Exit Sub
 End If
   If UserList(userindex).pos.Map = 106 Then
 Call SendData(SendTarget.toIndex, userindex, 0, "||Aqui no puedes invocar mascotas!." & FONTTYPE_INFO)
 Exit Sub
 End If
    If UserList(userindex).pos.Map = 114 Then
 Call SendData(SendTarget.toIndex, userindex, 0, "||Aqui no puedes invocar mascotas!." & FONTTYPE_INFO)
 Exit Sub
 End If
     If UserList(userindex).pos.Map = 115 Then
 Call SendData(SendTarget.toIndex, userindex, 0, "||Aqui no puedes invocar mascotas!." & FONTTYPE_INFO)
 Exit Sub
 End If
Dim h As Integer, j As Integer, ind As Integer, Index As Integer
Dim TargetPos As WorldPos


TargetPos.Map = UserList(userindex).flags.TargetMap
TargetPos.x = UserList(userindex).flags.TargetX
TargetPos.Y = UserList(userindex).flags.TargetY

h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
    
    
For j = 1 To Hechizos(h).Cant
    
    If UserList(userindex).NroMacotas < MAXMASCOTAS Then
        ind = SpawnNpc(Hechizos(h).NumNpc, TargetPos, True, False)
        If ind > 0 Then
            UserList(userindex).NroMacotas = UserList(userindex).NroMacotas + 1
            
            Index = FreeMascotaIndex(userindex)
            
            UserList(userindex).MascotasIndex(Index) = ind
            UserList(userindex).MascotasType(Index) = Npclist(ind).Numero
            
            Npclist(ind).MaestroUser = userindex
            Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion
            Npclist(ind).GiveGLD = 0
            
            Call FollowAmo(ind)
        End If
            
    Else
        Exit For
    End If
    
Next j


Call InfoHechizo(userindex)
b = True


End Sub

Sub HandleHechizoTerreno(ByVal userindex As Integer, ByVal uh As Integer)

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uInvocacion '
        Call HechizoInvocacion(userindex, b)
    Case TipoHechizo.uEstado
        Call HechizoTerrenoEstado(userindex, b)
    
End Select

If b Then
    Call SubirSkill(userindex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(userindex).Stats.MinMAN < 0 Then UserList(userindex).Stats.MinMAN = 0
    UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(userindex).Stats.MinSta < 0 Then UserList(userindex).Stats.MinSta = 0
    Call EnviarMn(userindex)
    Call EnviarSta(userindex)
End If


End Sub

Sub HandleHechizoUsuario(ByVal userindex As Integer, ByVal uh As Integer)

Dim b As Boolean
Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
       Call HechizoEstadoUsuario(userindex, b)
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
       Call HechizoPropUsuario(userindex, b)
End Select

If b Then
    Call SubirSkill(userindex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(userindex).Stats.MinMAN < 0 Then UserList(userindex).Stats.MinMAN = 0
    UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(userindex).Stats.MinSta < 0 Then UserList(userindex).Stats.MinSta = 0
    Call EnviarSta(userindex)
    Call EnviarMn(userindex)
    Call EnviarHP(UserList(userindex).flags.TargetUser)
    UserList(userindex).flags.TargetUser = 0
End If

End Sub

Sub HandleHechizoNPC(ByVal userindex As Integer, ByVal uh As Integer)
Dim h As Integer
h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)

If UserList(userindex).GuildIndex = 0 And Npclist(UserList(userindex).flags.TargetNPC).Numero = 906 Then
Call SendData(toIndex, userindex, 0, "||No tienes clan!" & FONTTYPE_INFO)
Exit Sub
Else
If UserList(userindex).GuildIndex > 0 Then
If Guilds(UserList(userindex).GuildIndex).GuildName = AlmacenaDominador And Npclist(UserList(userindex).flags.TargetNPC).Numero = 906 And Not Hechizos(h).SubeHP = 1 Then
Call SendData(toIndex, userindex, 0, "||No puedes acatar tu rey!!" & FONTTYPE_INFO)
Exit Sub
End If
End If
End If

If UserList(userindex).GuildIndex = 0 And Npclist(UserList(userindex).flags.TargetNPC).Numero = 616 Then
Call SendData(toIndex, userindex, 0, "||No tienes clan!" & FONTTYPE_INFO)
Exit Sub
Else
If UserList(userindex).GuildIndex > 0 Then
If Guilds(UserList(userindex).GuildIndex).GuildName = Lemuria And Npclist(UserList(userindex).flags.TargetNPC).Numero = 616 And Not Hechizos(h).SubeHP = 1 Then
Call SendData(toIndex, userindex, 0, "||No puedes acatar tu rey!!" & FONTTYPE_INFO)
Exit Sub
End If
End If
End If

If UserList(userindex).GuildIndex = 0 And Npclist(UserList(userindex).flags.TargetNPC).Numero = 617 Then
Call SendData(toIndex, userindex, 0, "||No tienes clan!" & FONTTYPE_INFO)
Exit Sub
Else
If UserList(userindex).GuildIndex > 0 Then
If Guilds(UserList(userindex).GuildIndex).GuildName = Tale And Npclist(UserList(userindex).flags.TargetNPC).Numero = 617 And Not Hechizos(h).SubeHP = 1 Then
Call SendData(toIndex, userindex, 0, "||No puedes acatar tu rey!!" & FONTTYPE_INFO)
Exit Sub
End If
End If
End If


If UserList(userindex).flags.demonio = True Then
If Npclist(UserList(userindex).flags.TargetNPC).Numero = 940 Then
Call SendData(toIndex, userindex, 0, "||No puedes acatar tu rey!!" & FONTTYPE_INFO)
Exit Sub
End If
End If

If UserList(userindex).flags.angel = True Then
If Npclist(UserList(userindex).flags.TargetNPC).Numero = 941 Then
Call SendData(toIndex, userindex, 0, "||No puedes acatar tu rey!!" & FONTTYPE_INFO)
Exit Sub
End If
End If

If UserList(userindex).GuildIndex = 0 And Npclist(UserList(userindex).flags.TargetNPC).Numero = 910 Then
Call SendData(toIndex, userindex, 0, "||No tienes clan!" & FONTTYPE_INFO)
Exit Sub
Else
If UserList(userindex).GuildIndex > 0 Then
If Guilds(UserList(userindex).GuildIndex).GuildName = AlmacenaDominadornix And Npclist(UserList(userindex).flags.TargetNPC).Numero = 910 And Not Hechizos(h).SubeHP = 1 Then
Call SendData(toIndex, userindex, 0, "||No puedes acatar tu rey!!" & FONTTYPE_INFO)
Exit Sub
End If
End If
End If

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
        Call HechizoEstadoNPC(UserList(userindex).flags.TargetNPC, uh, b, userindex)
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
        Call HechizoPropNPC(uh, UserList(userindex).flags.TargetNPC, userindex, b)
End Select

If b Then
    Call SubirSkill(userindex, Magia)
    UserList(userindex).flags.TargetNPC = 0
    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(userindex).Stats.MinMAN < 0 Then UserList(userindex).Stats.MinMAN = 0
    UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(userindex).Stats.MinSta < 0 Then UserList(userindex).Stats.MinSta = 0
    Call EnviarMn(userindex)
    Call EnviarSta(userindex)
End If

End Sub


Sub LanzarHechizo(Index As Integer, userindex As Integer)

Dim uh As Integer
Dim exito As Boolean

uh = UserList(userindex).Stats.UserHechizos(Index)

If UserList(userindex).flags.Desnudo = 1 Then
Call SendData(SendTarget.toIndex, userindex, 0, "||No podés atacar sin ropa." & FONTTYPE_WARNING)
Exit Sub
End If

If PuedeLanzar(userindex, uh) Then
    Select Case Hechizos(uh).Target
        
        Case TargetType.uUsuarios
            If UserList(userindex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(userindex).flags.TargetUser).pos.Y - UserList(userindex).pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(userindex, uh)
                Else
                    Call SendData(SendTarget.toIndex, userindex, 0, "||Estas demasiado lejos para lanzar este hechizo." & FONTTYPE_WARNING)
                End If
            Else
                Call SendData(SendTarget.toIndex, userindex, 0, "||Este hechizo actua solo sobre usuarios." & FONTTYPE_INFO)
            End If
        Case TargetType.uNPC
            If UserList(userindex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(userindex).flags.TargetNPC).pos.Y - UserList(userindex).pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(userindex, uh)
                Else
                    Call SendData(SendTarget.toIndex, userindex, 0, "||Estas demasiado lejos para lanzar este hechizo." & FONTTYPE_WARNING)
                End If
            Else
                Call SendData(SendTarget.toIndex, userindex, 0, "||Este hechizo solo afecta a los npcs." & FONTTYPE_INFO)
            End If
        Case TargetType.uUsuariosYnpc
            If UserList(userindex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(userindex).flags.TargetUser).pos.Y - UserList(userindex).pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(userindex, uh)
                Else
                    Call SendData(SendTarget.toIndex, userindex, 0, "||Estas demasiado lejos para lanzar este hechizo." & FONTTYPE_WARNING)
                End If
            ElseIf UserList(userindex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(userindex).flags.TargetNPC).pos.Y - UserList(userindex).pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(userindex, uh)
                Else
                    Call SendData(SendTarget.toIndex, userindex, 0, "||Estas demasiado lejos para lanzar este hechizo." & FONTTYPE_WARNING)
                End If
            Else
                Call SendData(SendTarget.toIndex, userindex, 0, "Z26")
            End If
        Case TargetType.uTerreno
            Call HandleHechizoTerreno(userindex, uh)
    End Select
    
End If



If UserList(userindex).Counters.Ocultando Then _
    UserList(userindex).Counters.Ocultando = UserList(userindex).Counters.Ocultando - 1
    
End Sub

Sub HechizoEstadoUsuario(ByVal userindex As Integer, ByRef b As Boolean)



Dim h As Integer, TU As Integer
h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
TU = UserList(userindex).flags.TargetUser

If UserList(userindex).flags.demonio = True And UserList(TU).flags.demonio = True And Not Hechizos(h).RemoverParalisis = 1 Then
Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes atacar tu bando!!" & FONTTYPE_INFO)
Exit Sub
End If

If UserList(userindex).flags.angel = True And UserList(TU).flags.angel = True And Not Hechizos(h).RemoverParalisis = 1 Then
Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes atacar tu bando!!" & FONTTYPE_INFO)
Exit Sub
End If

If Hechizos(h).Invisibilidad = 1 Then
   
   If UserList(userindex).pos.Map = 61 Then
   Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes lanzar invisibilidad en duelo!" & FONTTYPE_INFO)
   Exit Sub
   End If
   If UserList(userindex).pos.Map = 66 Then
   Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes lanzar invisibilidad en guerras!" & FONTTYPE_INFO)
   Exit Sub
   End If
   If UserList(userindex).pos.Map = 88 Then
   Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes lanzar invisibilidad en deathmatch!" & FONTTYPE_INFO)
   Exit Sub
   End If
   If UserList(userindex).pos.Map = 87 Then
   Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes lanzar invisibilidad en reto 2v2!" & FONTTYPE_INFO)
   Exit Sub
   End If
   If UserList(userindex).pos.Map = 62 Then
   Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes lanzar invisibilidad en torneo!" & FONTTYPE_INFO)
   Exit Sub
   End If
   
    If UserList(userindex).pos.Map = 86 Then
   Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes lanzar invisibilidad en sala de invocaciones!" & FONTTYPE_INFO)
   Exit Sub
   End If
   
    If UserList(userindex).pos.Map = 78 Then
   Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes lanzar invisibilidad en retos!" & FONTTYPE_INFO)
   Exit Sub
   End If
   
   If UserList(userindex).pos.Map = 79 Then
   Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes lanzar invisibilidad en torneos!" & FONTTYPE_INFO)
   Exit Sub
   End If
   
    If UserList(TU).flags.Muerto = 1 Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||¡Está muerto!" & FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
    
    If Criminal(TU) And Not Criminal(userindex) Then
        If UserList(userindex).flags.Seguro Then
            Call SendData(SendTarget.toIndex, userindex, 0, "Z6")
            Exit Sub
        Else
            Call VolverCriminal(userindex)
        End If
    End If
    
    UserList(TU).flags.Invisible = 1
#If SeguridadAlkon Then
    If EncriptarProtocolosCriticos Then
        Call SendCryptedData(SendTarget.ToMap, 0, UserList(TU).pos.Map, "NOVER" & UserList(TU).char.CharIndex & ",1")
    Else
#End If
        Call SendData(SendTarget.ToMap, 0, UserList(TU).pos.Map, "NOVER" & UserList(TU).char.CharIndex & ",1")
#If SeguridadAlkon Then
    End If
#End If
    Call InfoHechizo(userindex)
    b = True
End If

If Hechizos(h).Mimetiza = 1 Then
    If UserList(TU).flags.Muerto = 1 Then
        Exit Sub
    End If
    
    If UserList(TU).flags.Navegando = 1 Then
        Exit Sub
    End If
    If UserList(userindex).flags.Navegando = 1 Then
        Exit Sub
    End If
    
    If UserList(TU).flags.Privilegios >= PlayerType.Consejero Then
        Exit Sub
    End If
    
    If UserList(userindex).flags.Mimetizado = 1 Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||Ya te encuentras transformado. El hechizo no ha tenido efecto" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    'copio el char original al mimetizado
    
    With UserList(userindex)
        .CharMimetizado.Body = .char.Body
        .CharMimetizado.Head = .char.Head
        .CharMimetizado.CascoAnim = .char.CascoAnim
      
        .CharMimetizado.ShieldAnim = .char.ShieldAnim
        .CharMimetizado.WeaponAnim = .char.WeaponAnim
        
        .flags.Mimetizado = 1
        
        'ahora pongo local el del enemigo
        .char.Body = UserList(TU).char.Body
        .char.Head = UserList(TU).char.Head
     
        .char.CascoAnim = UserList(TU).char.CascoAnim
        .char.ShieldAnim = UserList(TU).char.ShieldAnim
        .char.WeaponAnim = UserList(TU).char.WeaponAnim
    
Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).OrigChar.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
    End With
   
   Call InfoHechizo(userindex)
   b = True
End If


If Hechizos(h).Envenena = 1 Then
        If Not PuedeAtacar(userindex, TU) Then Exit Sub
        If userindex <> TU Then
            Call UsuarioAtacadoPorUsuario(userindex, TU)
        End If
        UserList(TU).flags.Envenenado = 1
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(h).CuraVeneno = 1 Then
        UserList(TU).flags.Envenenado = 0
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(h).Maldicion = 1 Then
        If Not PuedeAtacar(userindex, TU) Then Exit Sub
        If userindex <> TU Then
            Call UsuarioAtacadoPorUsuario(userindex, TU)
        End If
        UserList(TU).flags.Maldicion = 1
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(h).RemoverMaldicion = 1 Then
        UserList(TU).flags.Maldicion = 0
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(h).Bendicion = 1 Then
        UserList(TU).flags.Bendicion = 1
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(h).Paraliza = 1 Or Hechizos(h).Inmoviliza = 1 Then
     If UserList(TU).flags.Paralizado = 0 Then
            If Not PuedeAtacar(userindex, TU) Then Exit Sub
            
            If userindex <> TU Then
                Call UsuarioAtacadoPorUsuario(userindex, TU)
            End If
            
            Call InfoHechizo(userindex)
            b = True
            If UserList(TU).Invent.HerramientaEqpObjIndex = SUPERANILLO Then
                Call SendData(SendTarget.toIndex, TU, 0, "|| Tu anillo rechaza los efectos del hechizo." & FONTTYPE_FIGHT)
                Call SendData(SendTarget.toIndex, userindex, 0, "|| ¡El hechizo no tiene efecto!" & FONTTYPE_FIGHT)
                Exit Sub
            End If
            
            UserList(TU).flags.Paralizado = 1
            UserList(TU).Counters.Paralisis = IntervaloParalizado
                Call SendData(SendTarget.toIndex, TU, 0, "PARADOW")
                Call SendData(SendTarget.toIndex, TU, 0, "PU" & UserList(TU).pos.x & "," & UserList(TU).pos.Y)

    End If
End If

If Hechizos(h).RemoverParalisis = 1 Then
    If UserList(TU).flags.Paralizado = 1 Then
        If Criminal(TU) And Not Criminal(userindex) Then
            If UserList(userindex).flags.Seguro Then
                Call SendData(SendTarget.toIndex, userindex, 0, "Z6")
                Exit Sub
            Else
                Call VolverCriminal(userindex)
            End If
        End If
        
        UserList(TU).flags.Paralizado = 0
        'no need to crypt this
        Call SendData(SendTarget.toIndex, TU, 0, "PARADOW")
        Call InfoHechizo(userindex)
        b = True
    End If
End If

If Hechizos(h).RemoverEstupidez = 1 Then
    If Not UserList(TU).flags.Estupidez = 0 Then
                UserList(TU).flags.Estupidez = 0
                'no need to crypt this
                Call SendData(SendTarget.toIndex, TU, 0, "NESTUP")
                Call InfoHechizo(userindex)
                b = True
    End If
End If


If Hechizos(h).Revivir = 1 Then
If UserList(userindex).pos.Map = 87 Then
Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes resucitar en retos" & FONTTYPE_INFO)
Exit Sub
End If
    If UserList(TU).flags.Muerto = 1 Then
        If Criminal(TU) And Not Criminal(userindex) Then
            If UserList(userindex).flags.Seguro Then
                Call SendData(SendTarget.toIndex, userindex, 0, "Z6")
                Exit Sub
            Else
                Call VolverCriminal(userindex)
            End If
        End If
 If UCase$(UserList(userindex).Clase) = "CLERIGO" Then
            If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(h).NeedStaff Then
                    Call SendData(SendTarget.toIndex, userindex, 0, "||Necesitas una mejor espada para este hechizo" & FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
            End If
            End If
        'revisamos si necesita vara
        If UCase$(UserList(userindex).Clase) = "MAGO" Then
            If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(h).NeedStaff Then
                    Call SendData(SendTarget.toIndex, userindex, 0, "||Necesitas un mejor báculo para este hechizo" & FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
            End If
        ElseIf UCase$(UserList(userindex).Clase) = "BARDO" Then
            If UserList(userindex).Invent.HerramientaEqpObjIndex <> LAUDMAGICO Then
                Call SendData(SendTarget.toIndex, userindex, 0, "||Necesitas un instrumento mágico para devolver la vida" & FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
        End If
        '/Juan Maraxus
        If Not Criminal(TU) Then
            If TU <> userindex Then
                UserList(userindex).Reputacion.NobleRep = UserList(userindex).Reputacion.NobleRep + 500
                If UserList(userindex).Reputacion.NobleRep > MAXREP Then _
                    UserList(userindex).Reputacion.NobleRep = MAXREP
                Call SendData(SendTarget.toIndex, userindex, 0, "||¡Los Dioses te sonrien, has ganado 500 puntos de nobleza!." & FONTTYPE_INFO)
            End If
        End If
        UserList(TU).Stats.MinMAN = 0
        Call EnviarMn(TU)
        '/Pablo Toxic Waste
        
        b = True
        Call InfoHechizo(userindex)
        Call RevivirUsuario(TU)
    Else
        b = False
    End If

End If

If Hechizos(h).Ceguera = 1 Then
        If Not PuedeAtacar(userindex, TU) Then Exit Sub
        If userindex <> TU Then
            Call UsuarioAtacadoPorUsuario(userindex, TU)
        End If
        UserList(TU).flags.Ceguera = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado / 3
#If SeguridadAlkon Then
        Call SendCryptedData(SendTarget.toIndex, TU, 0, "CEGU")
#Else
        Call SendData(SendTarget.toIndex, TU, 0, "CEGU")
#End If
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(h).Estupidez = 1 Then
        If Not PuedeAtacar(userindex, TU) Then Exit Sub
        If userindex <> TU Then
            Call UsuarioAtacadoPorUsuario(userindex, TU)
        End If
        UserList(TU).flags.Estupidez = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado
#If SeguridadAlkon Then
        If EncriptarProtocolosCriticos Then
            Call SendCryptedData(SendTarget.toIndex, TU, 0, "DUMB")
        Else
#End If
            Call SendData(SendTarget.toIndex, TU, 0, "DUMB")
#If SeguridadAlkon Then
        End If
#End If
        Call InfoHechizo(userindex)
        b = True
End If

End Sub
Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal userindex As Integer)

If Npclist(NpcIndex).Numero = 616 And Not Hechizos(hIndex).SubeHP = 1 Then
If Npclist(NpcIndex).Stats.MinHP > 14000 Then
        Call SendData(toAll, 0, 0, "LEMU")
        End If
        If Npclist(NpcIndex).Stats.MinHP < 3000 Then
        Call SendData(toAll, 0, 0, "LEMU")
    End If
    End If
    
    If Npclist(NpcIndex).Numero = 617 And Not Hechizos(hIndex).SubeHP = 1 Then
If Npclist(NpcIndex).Stats.MinHP > 14000 Then
        Call SendData(toAll, 0, 0, "TALE")
        End If
        If Npclist(NpcIndex).Stats.MinHP < 3000 Then
        Call SendData(toAll, 0, 0, "TALE")
    End If
    End If

If Npclist(NpcIndex).Numero = 906 And Not Hechizos(hIndex).SubeHP = 1 Then
If Npclist(NpcIndex).Stats.MinHP > 14000 Then
        Call SendData(toAll, 0, 0, "ULLA")
        End If
        If Npclist(NpcIndex).Stats.MinHP < 3000 Then
        Call SendData(toAll, 0, 0, "ULLA")
    End If
    End If
    
    If Npclist(NpcIndex).Numero = 910 And Not Hechizos(hIndex).SubeHP = 1 Then
If Npclist(NpcIndex).Stats.MinHP > 14000 Then
        Call SendData(toAll, 0, 0, "NIX")
        End If
        If Npclist(NpcIndex).Stats.MinHP < 3000 Then
        Call SendData(toAll, 0, 0, "NIX")
    End If
    End If
    
If Hechizos(hIndex).Invisibilidad = 1 Then
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Invisible = 1
   b = True
End If

If Hechizos(hIndex).Envenena = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(SendTarget.toIndex, userindex, 0, "Z7")
        Exit Sub
   End If
   
   If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
        If UserList(userindex).flags.Seguro Then
            Call SendData(SendTarget.toIndex, userindex, 0, "Z8")
            Exit Sub
        Else
            UserList(userindex).Reputacion.NobleRep = 0
            UserList(userindex).Reputacion.PlebeRep = 0
            UserList(userindex).Reputacion.AsesinoRep = UserList(userindex).Reputacion.AsesinoRep + 200
            If UserList(userindex).Reputacion.AsesinoRep > MAXREP Then _
                UserList(userindex).Reputacion.AsesinoRep = MAXREP
        End If
    End If
        
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Envenenado = 1
   b = True
End If

If Hechizos(hIndex).CuraVeneno = 1 Then
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Envenenado = 0
   b = True
End If

If Hechizos(hIndex).Maldicion = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(SendTarget.toIndex, userindex, 0, "Z7")
        Exit Sub
   End If
   
   If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
        If UserList(userindex).flags.Seguro Then
            Call SendData(SendTarget.toIndex, userindex, 0, "Z8")
            Exit Sub
        Else
            UserList(userindex).Reputacion.NobleRep = 0
            UserList(userindex).Reputacion.PlebeRep = 0
            UserList(userindex).Reputacion.AsesinoRep = UserList(userindex).Reputacion.AsesinoRep + 200
            If UserList(userindex).Reputacion.AsesinoRep > MAXREP Then _
                UserList(userindex).Reputacion.AsesinoRep = MAXREP
        End If
    End If
    
    Call InfoHechizo(userindex)
    Npclist(NpcIndex).flags.Maldicion = 1
    b = True
End If

If Hechizos(hIndex).RemoverMaldicion = 1 Then
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Maldicion = 0
   b = True
End If

If Hechizos(hIndex).Bendicion = 1 Then
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Bendicion = 1
   b = True
End If

If Hechizos(hIndex).Paraliza = 1 Then
    If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
        If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            If UserList(userindex).flags.Seguro Then
                Call SendData(SendTarget.toIndex, userindex, 0, "Z8")
                Exit Sub
            Else
                UserList(userindex).Reputacion.NobleRep = 0
                UserList(userindex).Reputacion.PlebeRep = 0
                UserList(userindex).Reputacion.AsesinoRep = UserList(userindex).Reputacion.AsesinoRep + 500
                If UserList(userindex).Reputacion.AsesinoRep > MAXREP Then _
                    UserList(userindex).Reputacion.AsesinoRep = MAXREP
            End If
        End If
        
        Call InfoHechizo(userindex)
        Npclist(NpcIndex).flags.Paralizado = 1
        Npclist(NpcIndex).flags.Inmovilizado = 0
        Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        b = True
    Else
        Call SendData(SendTarget.toIndex, userindex, 0, "Z9")
    End If
End If

'[Barrin 16-2-04]
If Hechizos(hIndex).RemoverParalisis = 1 Then
   If Npclist(NpcIndex).flags.Paralizado = 1 And Npclist(NpcIndex).MaestroUser = userindex Then
            Call InfoHechizo(userindex)
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = 0
            b = True
   Else
      Call SendData(SendTarget.toIndex, userindex, 0, "Z10")
   End If
End If
'[/Barrin]
 
If Hechizos(hIndex).Inmoviliza = 1 Then
    If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
        If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            If UserList(userindex).flags.Seguro Then
                Call SendData(SendTarget.toIndex, userindex, 0, "Z8")
                Exit Sub
            Else
                UserList(userindex).Reputacion.NobleRep = 0
                UserList(userindex).Reputacion.PlebeRep = 0
                UserList(userindex).Reputacion.AsesinoRep = UserList(userindex).Reputacion.AsesinoRep + 500
                If UserList(userindex).Reputacion.AsesinoRep > MAXREP Then _
                    UserList(userindex).Reputacion.AsesinoRep = MAXREP
            End If
        End If
        
        Npclist(NpcIndex).flags.Inmovilizado = 1
        Npclist(NpcIndex).flags.Paralizado = 0
        Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        Call InfoHechizo(userindex)
        b = True
    Else
        Call SendData(SendTarget.toIndex, userindex, 0, "Z9")
    End If
End If

End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal userindex As Integer, ByRef b As Boolean)

Dim daño As Long

If Npclist(NpcIndex).Numero = 616 And Not Hechizos(hIndex).SubeHP = 1 Then
If Npclist(NpcIndex).Stats.MinHP > 14000 Then
        Call SendData(toAll, 0, 0, "LEMU")
        End If
        If Npclist(NpcIndex).Stats.MinHP < 3000 Then
        Call SendData(toAll, 0, 0, "LEMU")
    End If
    End If
    
    If Npclist(NpcIndex).Numero = 617 And Not Hechizos(hIndex).SubeHP = 1 Then
If Npclist(NpcIndex).Stats.MinHP > 14000 Then
        Call SendData(toAll, 0, 0, "TALE")
        End If
        If Npclist(NpcIndex).Stats.MinHP < 3000 Then
        Call SendData(toAll, 0, 0, "TALE")
    End If
    End If
    
If Npclist(NpcIndex).Numero = 906 And Not Hechizos(hIndex).SubeHP = 1 Then
If Npclist(NpcIndex).Stats.MinHP > 14000 Then
        Call SendData(toAll, 0, 0, "ULLA")
        End If
        If Npclist(NpcIndex).Stats.MinHP < 3000 Then
        Call SendData(toAll, 0, 0, "ULLA")
    End If
    End If
    
    If Npclist(NpcIndex).Numero = 910 And Not Hechizos(hIndex).SubeHP = 1 Then
If Npclist(NpcIndex).Stats.MinHP > 14000 Then
        Call SendData(toAll, 0, 0, "NIX")
        End If
        If Npclist(NpcIndex).Stats.MinHP < 3000 Then
        Call SendData(toAll, 0, 0, "NIX")
    End If
    End If
'Salud

If Hechizos(hIndex).SubeHP = 1 Then
    daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(userindex).Stats.ELV)
    Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbCyan & "°+ " & daño & "!" & "°" & str(Npclist(NpcIndex).char.CharIndex))
    If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).VaraDragon = 1 And Npclist(NpcIndex).NPCtype = DRAGON Then
daño = daño * 40
End If
End If
    
    Call InfoHechizo(userindex)
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP + daño
    If Npclist(NpcIndex).Stats.MinHP > Npclist(NpcIndex).Stats.MaxHP Then _
        Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP
    Call SendData(SendTarget.toIndex, userindex, 0, "||Has curado " & daño & " puntos de salud a la criatura." & FONTTYPE_FIGHT)
    b = True
ElseIf Hechizos(hIndex).SubeHP = 2 Then
    
    If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(SendTarget.toIndex, userindex, 0, "Z7")
        b = False
        Exit Sub
    End If
    
    If Npclist(NpcIndex).NPCtype = 2 And UserList(userindex).flags.Seguro Then
        Call SendData(SendTarget.toIndex, userindex, 0, "Z8")
        b = False
        Exit Sub
    End If
    
    If Not PuedeAtacarNPC(userindex, NpcIndex) Then
        b = False
        Exit Sub
    End If
    
    daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(userindex).Stats.ELV)
    
If Npclist(NpcIndex).DefensaMagica = 1 Then
daño = daño - RandomNumber(150, 200)
End If
    
        If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).VaraDragon = 1 And Npclist(NpcIndex).NPCtype = DRAGON Then
daño = daño * 40
End If
End If

    If Hechizos(hIndex).StaffAffected Then
        If UCase$(UserList(userindex).Clase) = "MAGO" Then
            If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                daño = (daño * (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                'Aumenta daño segun el staff-
                'Daño = (Daño* (80 + BonifBáculo)) / 100
            Else
                daño = daño * 0.7 'Baja daño a 80% del original
            End If
        End If
    End If
    If UserList(userindex).Invent.HerramientaEqpObjIndex = LAUDMAGICO Then
        daño = daño * 1.04  'laud magico de los bardos
    End If
    


    Call InfoHechizo(userindex)
    b = True
    Call NpcAtacado(NpcIndex, userindex)
    If Npclist(NpcIndex).flags.Snd2 > 0 Then Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & Npclist(NpcIndex).flags.Snd2)
    
    Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbYellow & "°- " & daño & "!" & "°" & str(Npclist(NpcIndex).char.CharIndex))
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - daño
    SendData SendTarget.toIndex, userindex, 0, "||Le has causado " & daño & " puntos de daño a la criatura!" & FONTTYPE_FIGHT
    Call CalcularDarExp(userindex, NpcIndex, daño)

If Npclist(NpcIndex).Stats.MinHP < 1 Then
        Npclist(NpcIndex).Stats.MinHP = 0
        Call MuereNpc(NpcIndex, userindex)
Else
    'Mascotas atacan a la criatura.
    Call CheckPets(NpcIndex, userindex, True)
End If
End If

End Sub

Sub InfoHechizo(ByVal userindex As Integer)


    Dim h As Integer
    h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
    
    
    Call DecirPalabrasMagicas(Hechizos(h).PalabrasMagicas, userindex)
    
    If UserList(userindex).flags.TargetUser > 0 Then
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(UserList(userindex).flags.TargetUser).char.CharIndex & "," & Hechizos(h).FXgrh & "," & Hechizos(h).loops)
        Call SendData(SendTarget.ToPCArea, UserList(userindex).flags.TargetUser, UserList(userindex).pos.Map, "TW" & Hechizos(h).WAV)
    ElseIf UserList(userindex).flags.TargetNPC > 0 Then
        Call SendData(SendTarget.ToNPCArea, UserList(userindex).flags.TargetNPC, Npclist(UserList(userindex).flags.TargetNPC).pos.Map, "CFX" & Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex & "," & Hechizos(h).FXgrh & "," & Hechizos(h).loops)
        Call SendData(SendTarget.ToNPCArea, UserList(userindex).flags.TargetNPC, UserList(userindex).pos.Map, "TW" & Hechizos(h).WAV)
    End If
    
    If UserList(userindex).flags.TargetUser > 0 Then
        If userindex <> UserList(userindex).flags.TargetUser Then
            Call SendData(SendTarget.toIndex, userindex, 0, "||" & Hechizos(h).HechizeroMsg & " " & UserList(UserList(userindex).flags.TargetUser).name & FONTTYPE_FIGHT)
            Call SendData(SendTarget.toIndex, UserList(userindex).flags.TargetUser, 0, "||" & UserList(userindex).name & " " & Hechizos(h).TargetMsg & FONTTYPE_FIGHT)
        Else
            Call SendData(SendTarget.toIndex, userindex, 0, "||" & Hechizos(h).PropioMsg & FONTTYPE_FIGHT)
        End If
    ElseIf UserList(userindex).flags.TargetNPC > 0 Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||" & Hechizos(h).HechizeroMsg & " " & "la criatura." & FONTTYPE_FIGHT)
    End If

End Sub

Sub HechizoPropUsuario(ByVal userindex As Integer, ByRef b As Boolean)

Dim h As Integer
Dim daño As Integer
Dim tempChr As Integer
    
    
h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
tempChr = UserList(userindex).flags.TargetUser
      
      If UserList(userindex).flags.demonio = True And UserList(tempChr).flags.demonio = True And Not Hechizos(h).SubeFuerza = 1 And Not Hechizos(h).SubeAgilidad = 1 Then
Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes atacar tu bando!!" & FONTTYPE_INFO)
Exit Sub
End If

If UserList(userindex).flags.angel = True And UserList(tempChr).flags.angel = True And Not Hechizos(h).SubeFuerza = 1 And Not Hechizos(h).SubeAgilidad = 1 Then
Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes atacar tu bando!!" & FONTTYPE_INFO)
Exit Sub
End If
'If UserList(UserIndex).Name = "EL OSO" Then
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| le tiro el hechizo " & H & " a " & UserList(tempChr).Name & FONTTYPE_VENENO)
'End If
      
      
'Hambre
If Hechizos(h).SubeHam = 1 Then
    
    Call InfoHechizo(userindex)
    
    daño = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
    
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam + daño
    If UserList(tempChr).Stats.MinHam > UserList(tempChr).Stats.MaxHam Then _
        UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MaxHam
    
    If userindex <> tempChr Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||Le has restaurado " & daño & " puntos de hambre a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.toIndex, tempChr, 0, "||" & UserList(userindex).name & " te ha restaurado " & daño & " puntos de hambre." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.toIndex, userindex, 0, "||Te has restaurado " & daño & " puntos de hambre." & FONTTYPE_FIGHT)
    End If
    
    Call EnviarHambreYsed(tempChr)
    b = True
    
ElseIf Hechizos(h).SubeHam = 2 Then
    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    Else
        Exit Sub
    End If
    
    Call InfoHechizo(userindex)
    
    daño = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
    
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - daño
    
    If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
    
    If userindex <> tempChr Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||Le has quitado " & daño & " puntos de hambre a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.toIndex, tempChr, 0, "||" & UserList(userindex).name & " te ha quitado " & daño & " puntos de hambre." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.toIndex, userindex, 0, "||Te has quitado " & daño & " puntos de hambre." & FONTTYPE_FIGHT)
    End If
    
    Call EnviarHambreYsed(tempChr)
    
    b = True
    
    If UserList(tempChr).Stats.MinHam < 1 Then
        UserList(tempChr).Stats.MinHam = 0
        UserList(tempChr).flags.Hambre = 1
    End If
    
End If

'Sed
If Hechizos(h).SubeSed = 1 Then
    
    Call InfoHechizo(userindex)
    
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU + daño
    If UserList(tempChr).Stats.MinAGU > UserList(tempChr).Stats.MaxAGU Then _
        UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MaxAGU
         
    If userindex <> tempChr Then
      Call SendData(SendTarget.toIndex, userindex, 0, "||Le has restaurado " & daño & " puntos de sed a " & UserList(tempChr).name & FONTTYPE_FIGHT)
      Call SendData(SendTarget.toIndex, tempChr, 0, "||" & UserList(userindex).name & " te ha restaurado " & daño & " puntos de sed." & FONTTYPE_FIGHT)
    Else
      Call SendData(SendTarget.toIndex, userindex, 0, "||Te has restaurado " & daño & " puntos de sed." & FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(h).SubeSed = 2 Then
    
    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    End If
    
    Call InfoHechizo(userindex)
    
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - daño
    
    If userindex <> tempChr Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||Le has quitado " & daño & " puntos de sed a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.toIndex, tempChr, 0, "||" & UserList(userindex).name & " te ha quitado " & daño & " puntos de sed." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.toIndex, userindex, 0, "||Te has quitado " & daño & " puntos de sed." & FONTTYPE_FIGHT)
    End If
    
    If UserList(tempChr).Stats.MinAGU < 1 Then
            UserList(tempChr).Stats.MinAGU = 0
            UserList(tempChr).flags.Sed = 1
    End If
    
    b = True
End If

' <-------- Agilidad ---------->
If Hechizos(h).SubeAgilidad = 1 Then
    If Criminal(tempChr) And Not Criminal(userindex) Then
        If UserList(userindex).flags.Seguro Then
            Call SendData(SendTarget.toIndex, userindex, 0, "Z6")
            Exit Sub
        Else
            Call DisNobAuBan(userindex, UserList(userindex).Reputacion.NobleRep * 0.5, 10000)
        End If
    End If
    
    Call InfoHechizo(userindex)
    daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
    
    UserList(tempChr).flags.DuracionEfecto = 1200
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2) Then _
        UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2)
    UserList(tempChr).flags.TomoPocion = True
    b = True
    Call EnviarDopa(tempChr)
ElseIf Hechizos(h).SubeAgilidad = 2 Then
    
    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    End If
    
    Call InfoHechizo(userindex)
    
    UserList(tempChr).flags.TomoPocion = True
    daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
    b = True
    Call EnviarDopa(tempChr)
End If

' <-------- Fuerza ---------->
If Hechizos(h).SubeFuerza = 1 Then
    If Criminal(tempChr) And Not Criminal(userindex) Then
        If UserList(userindex).flags.Seguro Then
            Call SendData(SendTarget.toIndex, userindex, 0, "Z6")
            Exit Sub
        Else
            Call DisNobAuBan(userindex, UserList(userindex).Reputacion.NobleRep * 0.5, 10000)
        End If
    End If
    
    Call InfoHechizo(userindex)
    daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    
    UserList(tempChr).flags.DuracionEfecto = 1200

    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) * 2) Then _
        UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) * 2)
    
    UserList(tempChr).flags.TomoPocion = True
    b = True
    Call EnviarDopa(tempChr)
ElseIf Hechizos(h).SubeFuerza = 2 Then

    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    End If
    
    Call InfoHechizo(userindex)
    
    UserList(tempChr).flags.TomoPocion = True
    
    daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
    b = True
    Call EnviarDopa(tempChr)
End If

'Salud
If Hechizos(h).SubeHP = 1 Then
    If UserList(userindex).pos.Map = 1 Then
     Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes curar en ullathorpe")
    Exit Sub
    End If
    If Criminal(tempChr) And Not Criminal(userindex) Then
        If UserList(userindex).flags.Seguro Then
            Call SendData(SendTarget.toIndex, userindex, 0, "Z6")
            Exit Sub
        Else
            Call DisNobAuBan(userindex, UserList(userindex).Reputacion.NobleRep * 0.5, 10000)
        End If
    End If
    
    
    daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(userindex).Stats.ELV)
    
    Call InfoHechizo(userindex)

    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP + daño
    If UserList(tempChr).Stats.MinHP > UserList(tempChr).Stats.MaxHP Then _
        UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MaxHP
    
    If userindex <> tempChr Then
    
        Call SendData(SendTarget.toIndex, userindex, 0, "||Le has restaurado " & daño & " puntos de vida a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.toIndex, tempChr, 0, "||" & UserList(userindex).name & " te ha restaurado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.toIndex, userindex, 0, "||Te has restaurado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
    End If
    
    b = True
ElseIf Hechizos(h).SubeHP = 2 Then
    
        'If UserList(UserIndex).flags.SeguroClan Then
        'If Guilds(UserList(tempChr).GuildIndex).GuildName = Guilds(UserList(UserIndex).GuildIndex).GuildName And Guilds(UserList(UserIndex).GuildIndex).GuildName <> "" Then
        '    Call SendData(ToIndex, UserIndex, 0, "||No puedes atacar a tu propio Clan con el seguro activado, escribe /SEGCLAN para desactivarlo." & FONTTYPE_FIGHT)
        '    Exit Sub
        'End If
    'End If
    
    If userindex = tempChr Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||No podes atacarte a vos mismo." & FONTTYPE_FIGHT)
        Exit Sub
    End If
    
    daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
    
'If UserList(UserIndex).Name = "EL OSO" Then
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| danio, minhp, maxhp " & daño & " " & Hechizos(H).MinHP & " " & Hechizos(H).MaxHP & FONTTYPE_VENENO)
'End If
    
    
    daño = daño + Porcentaje(daño, 3 * UserList(userindex).Stats.ELV)
    
'If UserList(UserIndex).Name = "EL OSO" Then
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| daño, ELV " & daño & " " & UserList(UserIndex).Stats.ELV & FONTTYPE_VENENO)
'End If
    
    
    If Hechizos(h).StaffAffected Then
        If UCase$(UserList(userindex).Clase) = "MAGO" Then
            If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                daño = (daño * (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
            Else
                daño = daño * 0.7 'Baja daño a 70% del original
            End If
        End If
    End If
    
    If UserList(userindex).Invent.HerramientaEqpObjIndex = LAUDMAGICO Then
        daño = daño * 1.04  'laud magico de los bardos
    End If
    
    'cascos antimagia
    If (UserList(tempChr).Invent.CascoEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMax)
    End If
    
    If UserList(tempChr).Invent.ArmourEqpObjIndex > 0 Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).DefensaMagicaMax)
        End If
    
    'anillos
    If (UserList(tempChr).Invent.HerramientaEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.HerramientaEqpObjIndex).DefensaMagicaMax)
    End If
    
    If daño < 0 Then daño = 0
    
    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    End If
    
    Call InfoHechizo(userindex)
    
    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - daño
    Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbRed & "°- " & daño & "!" & "°" & str(UserList(tempChr).char.CharIndex))
    Call SendData(SendTarget.toIndex, userindex, 0, "||Le has quitado " & daño & " puntos de vida a " & UserList(tempChr).name & FONTTYPE_FIGHT)
    Call SendData(SendTarget.toIndex, tempChr, 0, "||" & UserList(userindex).name & " te ha quitado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
    
    'Muere
    If UserList(tempChr).Stats.MinHP < 1 Then
        Call ContarMuerte(tempChr, userindex)
        UserList(tempChr).Stats.MinHP = 0
        Call ActStats(tempChr, userindex)
        Call UserDie(tempChr)
    End If
    
    b = True
End If

'Mana
If Hechizos(h).SubeMana = 1 Then
    
    Call InfoHechizo(userindex)
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN + daño
    If UserList(tempChr).Stats.MinMAN > UserList(tempChr).Stats.MaxMAN Then _
        UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MaxMAN
    
    If userindex <> tempChr Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||Le has restaurado " & daño & " puntos de mana a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.toIndex, tempChr, 0, "||" & UserList(userindex).name & " te ha restaurado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.toIndex, userindex, 0, "||Te has restaurado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(h).SubeMana = 2 Then
    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    End If
    
    Call InfoHechizo(userindex)
    
    If userindex <> tempChr Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||Le has quitado " & daño & " puntos de mana a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.toIndex, tempChr, 0, "||" & UserList(userindex).name & " te ha quitado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.toIndex, userindex, 0, "||Te has quitado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - daño
    If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0
    b = True
    
End If

'Stamina
If Hechizos(h).SubeSta = 1 Then
    Call InfoHechizo(userindex)
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta + daño
    If UserList(tempChr).Stats.MinSta > UserList(tempChr).Stats.MaxSta Then _
        UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MaxSta
    If userindex <> tempChr Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||Le has restaurado " & daño & " puntos de vitalidad a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.toIndex, tempChr, 0, "||" & UserList(userindex).name & " te ha restaurado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.toIndex, userindex, 0, "||Te has restaurado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    b = True
ElseIf Hechizos(h).SubeMana = 2 Then
    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    End If
    
    Call InfoHechizo(userindex)
    
    If userindex <> tempChr Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||Le has quitado " & daño & " puntos de vitalidad a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.toIndex, tempChr, 0, "||" & UserList(userindex).name & " te ha quitado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.toIndex, userindex, 0, "||Te has quitado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - daño
    Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbRed & "°- " & daño & "!" & "°" & str(UserList(tempChr).char.CharIndex))
    If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0
    b = True
End If


End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal userindex As Integer, ByVal Slot As Byte)

'Call LogTarea("Sub UpdateUserHechizos")

Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(userindex).Stats.UserHechizos(Slot) > 0 Then
        Call ChangeUserHechizo(userindex, Slot, UserList(userindex).Stats.UserHechizos(Slot))
    Else
        Call ChangeUserHechizo(userindex, Slot, 0)
    End If

Else

'Actualiza todos los slots
For LoopC = 1 To MAXUSERHECHIZOS

        'Actualiza el inventario
        If UserList(userindex).Stats.UserHechizos(LoopC) > 0 Then
            Call ChangeUserHechizo(userindex, LoopC, UserList(userindex).Stats.UserHechizos(LoopC))
        Else
            Call ChangeUserHechizo(userindex, LoopC, 0)
        End If

Next LoopC

End If

End Sub

Sub ChangeUserHechizo(ByVal userindex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)

'Call LogTarea("ChangeUserHechizo")

UserList(userindex).Stats.UserHechizos(Slot) = Hechizo


If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then

    Call SendData(SendTarget.toIndex, userindex, 0, "SHS" & Slot & "," & Hechizo & "," & Hechizos(Hechizo).nombre)

Else

    Call SendData(SendTarget.toIndex, userindex, 0, "SHS" & Slot & "," & "0" & "," & "(Vacío)")

End If


End Sub


Public Sub DesplazarHechizo(ByVal userindex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)

If Not (Dire >= 1 And Dire <= 2) Then Exit Sub
If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

Dim TempHechizo As Integer

If Dire = 1 Then 'Mover arriba
    If CualHechizo = 1 Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes mover el hechizo en esa direccion." & FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(userindex).Stats.UserHechizos(CualHechizo)
        UserList(userindex).Stats.UserHechizos(CualHechizo) = UserList(userindex).Stats.UserHechizos(CualHechizo - 1)
        UserList(userindex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo
        
        Call UpdateUserHechizos(False, userindex, CualHechizo - 1)
    End If
Else 'mover abajo
    If CualHechizo = MAXUSERHECHIZOS Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes mover el hechizo en esa direccion." & FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(userindex).Stats.UserHechizos(CualHechizo)
        UserList(userindex).Stats.UserHechizos(CualHechizo) = UserList(userindex).Stats.UserHechizos(CualHechizo + 1)
        UserList(userindex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo
        
        Call UpdateUserHechizos(False, userindex, CualHechizo + 1)
    End If
End If
Call UpdateUserHechizos(False, userindex, CualHechizo)

End Sub


Public Sub DisNobAuBan(ByVal userindex As Integer, NoblePts As Long, BandidoPts As Long)
'disminuye la nobleza NoblePts puntos y aumenta el bandido BandidoPts puntos

    'Si estamos en la arena no hacemos nada
    If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).trigger = 6 Then Exit Sub
    
    'pierdo nobleza...
    UserList(userindex).Reputacion.NobleRep = UserList(userindex).Reputacion.NobleRep - NoblePts
    If UserList(userindex).Reputacion.NobleRep < 0 Then
        UserList(userindex).Reputacion.NobleRep = 0
    End If
    
    'gano bandido...
    UserList(userindex).Reputacion.BandidoRep = UserList(userindex).Reputacion.BandidoRep + BandidoPts
    If UserList(userindex).Reputacion.BandidoRep > MAXREP Then _
        UserList(userindex).Reputacion.BandidoRep = MAXREP
    Call SendData(SendTarget.toIndex, userindex, 0, "PN")
    If Criminal(userindex) Then If UserList(userindex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(userindex)
End Sub
