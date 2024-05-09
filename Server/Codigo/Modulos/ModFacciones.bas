Attribute VB_Name = "ModFacciones"



Option Explicit

Public ArmaduraImperial1 As Integer 'Primer jerarquia
Public ArmaduraImperial2 As Integer 'Segunda jerarquía
Public ArmaduraImperial3 As Integer 'Enanos
Public TunicaMagoImperial As Integer 'Magos
Public TunicaMagoImperialEnanos As Integer 'Magos

Public VestimentaImperialHumano As Integer
Public VestimentaImperialEnano As Integer
Public TunicaConspicuaHumano As Integer
Public TunicaConspicuaEnano As Integer
Public ArmaduraNobilisimaHumano As Integer
Public ArmaduraNobilisimaEnano As Integer
Public ArmaduraGranSacerdote As Integer

Public VestimentaLegionHumano As Integer
Public VestimentaLegionEnano As Integer
Public TunicaLobregaHumano As Integer
Public TunicaLobregaEnano As Integer
Public TunicaEgregiaHumano As Integer
Public TunicaEgregiaEnano As Integer
Public SacerdoteDemoniaco As Integer

Public ArmaduraCaos1 As Integer
Public TunicaMagoCaos As Integer
Public TunicaMagoCaosEnanos As Integer
Public ArmaduraCaos2 As Integer
Public ArmaduraCaos3 As Integer

Public Const ExpAlUnirse As Long = 50000
Public Const ExpX100 As Integer = 5000


Public Sub EnlistarArmadaReal(ByVal userindex As Integer)

If UserList(userindex).Faccion.ArmadaReal = 1 Then
    Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "Ya perteneces a las tropas reales!!! Ve a combatir criminales!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

If UserList(userindex).Faccion.FuerzasCaos = 1 Then
    Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "Maldito insolente!!! vete de aqui seguidor de las sombras!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

If Criminal(userindex) Then
    Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "No se permiten criminales en el ejercito imperial!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

If UserList(userindex).Faccion.CriminalesMatados < 10 Then
    Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes matar al menos 10 criminales, solo has matado " & UserList(userindex).Faccion.CriminalesMatados & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

If UserList(userindex).Stats.ELV < 25 Then
    Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If
 
If UserList(userindex).Faccion.CiudadanosMatados > 0 Then
    Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

If UserList(userindex).Faccion.Reenlistadas > 4 Then
    Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "Has sido expulsado de las fuerzas reales demasiadas veces!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

UserList(userindex).Faccion.ArmadaReal = 1
UserList(userindex).Faccion.Reenlistadas = UserList(userindex).Faccion.Reenlistadas + 1

UserList(userindex).Faccion.RecompensasReal = UserList(userindex).Faccion.CriminalesMatados \ 100

Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "¡¡¡Bienvenido a al Ejercito Imperial!!!, aqui tienes tus vestimentas. Por cada centena de criminales que acabes te daré un recompensa, buena suerte soldado!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))

If UserList(userindex).Faccion.RecibioArmaduraReal = 0 Then
    Dim MiObj As Obj
    Dim MiObj2 As Obj
    MiObj.Amount = 1
    MiObj2.Amount = 1
    
    
    
    
'Public VestimentaImperialHumano As Integer
'Public VestimentaImperialEnano As Integer
'Public TunicaConspicuaHumano As Integer
'Public TunicaConspicuaEnano As Integer
'Public ArmaduraNobilisimaHumano As Integer
'Public ArmaduraNobilisimaEnano As Integer
'Public ArmaduraGranSacerdote As Integer

'Public VestimentaLegionHumano As Integer
'Public VestimentaLegionEnano As Integer
'Public TunicaLobregaHumano As Integer
'Public TunicaLobregaEnano As Integer
'Public TunicaEgregiaHumano As Integer
'Public TunicaEgregiaEnano As Integer
'Public SacerdoteDemoniaco As Integer
'
    
        
    If UCase$(UserList(userindex).Raza) = "ENANO" Or UCase$(UserList(userindex).Raza) = "GNOMO" Then
        MiObj.ObjIndex = VestimentaImperialEnano
        Select Case UCase$(UserList(userindex).Clase)
            Case "MAGO"
                MiObj2.ObjIndex = TunicaConspicuaEnano
            Case Else
                MiObj2.ObjIndex = ArmaduraNobilisimaEnano
        End Select
    Else
        MiObj.ObjIndex = VestimentaImperialHumano
        Select Case UCase$(UserList(userindex).Clase)
            Case "MAGO"
                MiObj2.ObjIndex = TunicaConspicuaHumano
            Case "CLERIGO", "DRUIDA", "BARDO"
                MiObj2.ObjIndex = ArmaduraGranSacerdote
            Case Else
                MiObj2.ObjIndex = ArmaduraNobilisimaHumano
        End Select
    End If
    
    If Not MeterItemEnInventario(userindex, MiObj) Then
            Call TirarItemAlPiso(UserList(userindex).pos, MiObj)
    End If
    If Not MeterItemEnInventario(userindex, MiObj2) Then
            Call TirarItemAlPiso(UserList(userindex).pos, MiObj2)
    End If
    
    UserList(userindex).Faccion.RecibioArmaduraReal = 1
End If

If UserList(userindex).Faccion.RecibioExpInicialReal = 0 Then
    UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + ExpAlUnirse
    If UserList(userindex).Stats.Exp > MAXEXP Then _
        UserList(userindex).Stats.Exp = MAXEXP
    Call SendData(SendTarget.toIndex, userindex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(userindex).Faccion.RecibioExpInicialReal = 1
    Call CheckUserLevel(userindex)
End If


Call LogEjercitoReal(UserList(userindex).name)

End Sub

Public Sub RecompensaArmadaReal(ByVal userindex As Integer)

If UserList(userindex).Faccion.CriminalesMatados \ 100 = _
   UserList(userindex).Faccion.RecompensasReal Then
    Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, mata 100 crinales mas para recibir la proxima!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
Else
    Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
    UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + ExpX100
    If UserList(userindex).Stats.Exp > MAXEXP Then _
        UserList(userindex).Stats.Exp = MAXEXP
    Call SendData(SendTarget.toIndex, userindex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(userindex).Faccion.RecompensasReal = UserList(userindex).Faccion.RecompensasReal + 1
    Call CheckUserLevel(userindex)
End If

End Sub

Public Sub ExpulsarFaccionReal(ByVal userindex As Integer)

    UserList(userindex).Faccion.ArmadaReal = 0
    'Call PerderItemsFaccionarios(UserIndex)
    Call SendData(SendTarget.toIndex, userindex, 0, "||Has sido expulsado de las tropas reales!!!." & FONTTYPE_FIGHT)
    'Desequipamos la armadura real si está equipada
    If ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(userindex, UserList(userindex).Invent.ArmourEqpSlot)
End Sub

Public Sub ExpulsarFaccionCaos(ByVal userindex As Integer)

    UserList(userindex).Faccion.FuerzasCaos = 0
    'Call PerderItemsFaccionarios(UserIndex)
    Call SendData(SendTarget.toIndex, userindex, 0, "||Has sido expulsado de la legión oscura!!!." & FONTTYPE_FIGHT)
    'Desequipamos la armadura real si está equipada
    If ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(userindex, UserList(userindex).Invent.ArmourEqpSlot)
End Sub

Public Function TituloReal(ByVal userindex As Integer) As String

Select Case UserList(userindex).Faccion.RecompensasReal
  Case 0
        TituloReal = "Estudiante"
    Case 1
        TituloReal = "Director"
    Case 2
        TituloReal = "Oficial del bien"
    Case 3
        TituloReal = "Flasheador"
    Case 4
        TituloReal = "Adorador del bien"
    Case 5
        TituloReal = "Jefe"
    Case 6
        TituloReal = "Intinishko"
    Case 7
        TituloReal = "Agente"
    Case 8
        TituloReal = "Seguridad"
    Case 9
        TituloReal = "AntiCaos"
    Case Else
        TituloReal = "Iluminado"
End Select

End Function

Public Sub EnlistarCaos(ByVal userindex As Integer)

If Not Criminal(userindex) Then
    Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "Largate de aqui, bufon!!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

If UserList(userindex).Faccion.FuerzasCaos = 1 Then
    Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "Ya perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

If UserList(userindex).Faccion.ArmadaReal = 1 Then
    Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "Las sombras reinaran en Argentum, largate de aqui estupido ciudadano.!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

'[Barrin 17-12-03] Si era miembro de la Armada Real no se puede enlistar
If UserList(userindex).Faccion.RecibioExpInicialReal = 1 Then 'Tomamos el valor de ahí: ¿Recibio la experiencia para entrar?
    Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "No permitiré que ningún insecto real ingrese ¡Traidor del Rey!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If
'[/Barrin]

If Not Criminal(userindex) Then
    Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "Ja ja ja tu no eres bienvenido aqui!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

If UserList(userindex).Faccion.CiudadanosMatados < 10 Then
    Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes matar al menos 10 ciudadanos, solo has matado " & UserList(userindex).Faccion.CiudadanosMatados & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

If UserList(userindex).Stats.ELV < 25 Then
    Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If


If UserList(userindex).Faccion.Reenlistadas > 4 Then
    If UserList(userindex).Faccion.Reenlistadas = 200 Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "Has sido expulsado de las fuerzas oscuras y durante tu rebeldía has atacado a mi ejército. Vete de aquí!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
    Else
        Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "Has sido expulsado de las fuerzas oscuras demasiadas veces!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
    End If
    Exit Sub
End If

UserList(userindex).Faccion.Reenlistadas = UserList(userindex).Faccion.Reenlistadas + 1
UserList(userindex).Faccion.FuerzasCaos = 1
UserList(userindex).Faccion.RecompensasCaos = UserList(userindex).Faccion.CiudadanosMatados \ 100

Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "Bienvenido a al lado oscuro!!!, aqui tienes tu armadura. Por cada centena de ciudadanos que acabes te daré un recompensa, buena suerte soldado!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))

If UserList(userindex).Faccion.RecibioArmaduraCaos = 0 Then
    Dim MiObj As Obj
    Dim MiObj2 As Obj
    MiObj.Amount = 1
    MiObj2.Amount = 1
    
    If UCase$(UserList(userindex).Raza) = "ENANO" Or UCase$(UserList(userindex).Raza) = "GNOMO" Then
        MiObj.ObjIndex = VestimentaLegionEnano
        Select Case UCase$(UserList(userindex).Clase)
            Case "MAGO"
                MiObj2.ObjIndex = TunicaEgregiaEnano
            Case Else
                MiObj2.ObjIndex = TunicaLobregaEnano
        End Select
    Else
        MiObj.ObjIndex = VestimentaLegionHumano
        Select Case UCase$(UserList(userindex).Clase)
            Case "MAGO"
                MiObj2.ObjIndex = TunicaEgregiaHumano
            Case "CLERIGO", "DRUIDA", "BARDO"
                MiObj2.ObjIndex = SacerdoteDemoniaco
            Case Else
                MiObj2.ObjIndex = TunicaLobregaHumano
        End Select
    End If
    
    If Not MeterItemEnInventario(userindex, MiObj) Then
            Call TirarItemAlPiso(UserList(userindex).pos, MiObj)
    End If
    If Not MeterItemEnInventario(userindex, MiObj2) Then
            Call TirarItemAlPiso(UserList(userindex).pos, MiObj2)
    End If
    
    UserList(userindex).Faccion.RecibioArmaduraCaos = 1
    


End If

If UserList(userindex).Faccion.RecibioExpInicialCaos = 0 Then
    UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + ExpAlUnirse
    If UserList(userindex).Stats.Exp > MAXEXP Then _
        UserList(userindex).Stats.Exp = MAXEXP
    Call SendData(SendTarget.toIndex, userindex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(userindex).Faccion.RecibioExpInicialCaos = 1
    Call CheckUserLevel(userindex)
End If


Call LogEjercitoCaos(UserList(userindex).name)

End Sub

Public Sub RecompensaCaos(ByVal userindex As Integer)

If UserList(userindex).Faccion.CiudadanosMatados \ 100 = _
   UserList(userindex).Faccion.RecompensasCaos Then
    Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, mata 100 ciudadanos mas para recibir la proxima!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
Else
    Call SendData(SendTarget.toIndex, userindex, 0, "||" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
    UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + ExpX100
    If UserList(userindex).Stats.Exp > MAXEXP Then _
        UserList(userindex).Stats.Exp = MAXEXP
    Call SendData(SendTarget.toIndex, userindex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(userindex).Faccion.RecompensasCaos = UserList(userindex).Faccion.RecompensasCaos + 1
    Call CheckUserLevel(userindex)
End If


End Sub

Public Function TituloCaos(ByVal userindex As Integer) As String
Select Case UserList(userindex).Faccion.RecompensasCaos
    Case 0
        TituloCaos = "Novato"
    Case 1
        TituloCaos = "Servidor indulgente"
    Case 2
        TituloCaos = "Principe oscuro"
    Case 3
        TituloCaos = "Mutilante sombrio"
    Case 4
        TituloCaos = "Indulgente"
    Case 5
        TituloCaos = "Triturador"
    Case 6
        TituloCaos = "Arrogante"
    Case 7
        TituloCaos = "Heraldo Hástico"
    Case 8
        TituloCaos = "Anikilador"
    Case Else
        TituloCaos = "Exorcista"
End Select


End Function

'[Barrin 17-12-03]
'Sub PerderItemsFaccionarios(ByVal UserIndex As Integer)
'Dim i As Byte
'Dim MiObj As Obj
'Dim ItemIndex As Integer
'
'For i = 1 To MAX_INVENTORY_SLOTS
'  ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
'  If ItemIndex > 0 Then
'         If ObjData(ItemIndex).Real = 1 Or ObjData(ItemIndex).Caos = 1 Then
'            Call QuitarUserInvItem(UserIndex, i, UserList(UserIndex).Invent.Object(i).Amount)
'            Call UpdateUserInv(False, UserIndex, i)
'            If ObjData(ItemIndex).ObjType = eOBJType.Armour Then
'                If ObjData(ItemIndex).Real = 1 Then UserList(UserIndex).Faccion.RecibioArmaduraReal = 0
'                If ObjData(ItemIndex).Caos = 1 Then UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0
'            Else
'                UserList(UserIndex).Faccion.RecibioItemFaccionario = 0
'            End If
'         End If
'
'  End If
'Next i
'
'End Sub
'[/Barrin]
