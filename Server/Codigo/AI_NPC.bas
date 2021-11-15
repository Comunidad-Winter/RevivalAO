Attribute VB_Name = "AI"


Option Explicit

Public Enum TipoAI
    ESTATICO = 1
    MueveAlAzar = 2
    NpcMaloAtacaUsersBuenos = 3
    NPCDEFENSA = 4
    GuardiasAtacanCriminales = 5
    SigueAmo = 8
    NpcAtacaNpc = 9
    NpcPathfinding = 10
End Enum

Public Const ELEMENTALFUEGO As Integer = 93
Public Const ELEMENTALTIERRA As Integer = 94
Public Const ELEMENTALAGUA As Integer = 92

'Damos a los NPCs el mismo rango de visión que un PJ
Public Const RANGO_VISION_X As Byte = 8
Public Const RANGO_VISION_Y As Byte = 6

Public Enum e_Alineacion
    ninguna = 0
    Real = 1
    Caos = 2
    Neutro = 3
End Enum

Public Enum e_Personalidad
''Inerte: no tiene objetivos de ningun tipo (npcs vendedores, curas, etc)
''Agresivo no magico: Su objetivo es acercarse a las victimas para atacarlas
''Agresivo magico: Su objetivo es mantenerse lo mas lejos posible de sus victimas y atacarlas con magia
''Mascota: Solo ataca a quien ataque a su amo.
''Pacifico: No ataca.
    ninguna = 0
    Inerte = 1
    AgresivoNoMagico = 2
    AgresivoMagico = 3
    Macota = 4
    Pacifico = 5
End Enum

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo AI_NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'AI de los NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Private Sub HandleAlineacion(ByVal NpcIndex As Integer)
Dim Al As e_Alineacion
Dim Pe As e_Personalidad
Dim TargetPJ As Integer
Dim TargetNPC As Integer
Dim TieneTarget As Boolean
Dim EsNpc As Boolean

    TieneTarget = False
    Al = Npclist(NpcIndex).flags.AIAlineacion
    TargetPJ = Npclist(NpcIndex).flags.AtacaAPJ
    TargetNPC = Npclist(NpcIndex).flags.AtacaANPC
    
    
    Select Case Al
        Case e_Alineacion.Caos
            If TargetPJ > 0 Then
                If InRangoVisionNPC(NpcIndex, UserList(TargetPJ).pos.x, UserList(TargetPJ).pos.y) Then
                    If Not Criminal(TargetPJ) Then
                        TieneTarget = True
                    Else
                        Npclist(NpcIndex).flags.AtacaAPJ = 0
                    End If
                Else
                    Npclist(NpcIndex).flags.AtacaAPJ = 0
                End If
            End If
            If TargetNPC > 0 Then
                If InRangoVisionNPC(NpcIndex, Npclist(TargetNPC).pos.x, Npclist(TargetNPC).pos.y) Then
                    TieneTarget = True
                Else
                    Npclist(NpcIndex).flags.AtacaANPC = 0
                End If
            End If
        Case e_Alineacion.Neutro
            If TargetPJ > 0 Then
                If InRangoVisionNPC(NpcIndex, UserList(TargetPJ).pos.x, UserList(TargetPJ).pos.y) Then
                    TieneTarget = True
                Else
                    Npclist(NpcIndex).flags.AtacaAPJ = 0
                End If
            End If
            If TargetNPC > 0 Then
                If InRangoVisionNPC(NpcIndex, Npclist(TargetNPC).pos.x, Npclist(TargetNPC).pos.y) Then
                    TieneTarget = True
                Else
                    Npclist(NpcIndex).flags.AtacaANPC = 0
                End If
            End If
        Case e_Alineacion.ninguna
            Exit Sub
        Case e_Alineacion.Real
            If TargetPJ > 0 Then
                If InRangoVisionNPC(NpcIndex, UserList(TargetPJ).pos.x, UserList(TargetPJ).pos.y) Then
                    If Criminal(TargetPJ) Then
                        TieneTarget = True
                    Else
                        Npclist(NpcIndex).flags.AtacaAPJ = 0
                    End If
                Else
                    Npclist(NpcIndex).flags.AtacaAPJ = 0
                End If
            End If
            If TargetNPC > 0 Then
                If InRangoVisionNPC(NpcIndex, Npclist(TargetNPC).pos.x, Npclist(TargetNPC).pos.y) Then
                    TieneTarget = True
                Else
                    Npclist(NpcIndex).flags.AtacaANPC = 0
                End If
            End If
    End Select
    
    If Not TieneTarget Then
        
    
    End If

End Sub

Private Function AcquireNewTargetForAlignment(ByVal NpcIndex As Integer, ByRef EsNpc As Boolean) As Integer
Dim r As Byte
Dim NPCPosX As Byte
Dim NPCPosY As Byte
Dim NpcBestTarget As Integer
Dim PJBestTarget As Integer
Dim PJ As Integer
Dim npc As Integer

Dim x As Integer
Dim y As Integer
Dim m As Integer

    NPCPosX = Npclist(NpcIndex).pos.x
    NPCPosY = Npclist(NpcIndex).pos.y
    m = Npclist(NpcIndex).pos.Map
    
    For r = 1 To MinYBorder
        For x = NPCPosX - r To NPCPosX + r
            For y = NPCPosY - r To NPCPosY + r
                PJ = MapData(m, x, y).UserIndex
                npc = MapData(m, x, y).NpcIndex
                
                If PJ > 0 Then
                    Select Case Npclist(NpcIndex).flags.AIAlineacion
                        Case e_Alineacion.Caos
                            If Not Criminal(PJ) And Not UserList(PJ).flags.Muerto And Not UserList(PJ).flags.Invisible And Not UserList(PJ).flags.Oculto And UserList(PJ).flags.Privilegios = PlayerType.User Then
                                PJBestTarget = PJ
                            End If
                        Case e_Alineacion.Real
                        
                        Case e_Alineacion.Neutro
                    
                    End Select
                
                End If
                If MapData(m, x, y).NpcIndex > 0 Then
                
                End If
            Next y
        Next x
        If PJBestTarget > 0 Then
            EsNpc = False
            AcquireNewTargetForAlignment = PJBestTarget
            Exit Function
        End If
        If NpcBestTarget > 0 Then
            EsNpc = True
            AcquireNewTargetForAlignment = NpcBestTarget
            Exit Function
        End If
        
    Next r
            

End Function


Private Sub GuardiasAI(ByVal NpcIndex As Integer, Optional ByVal DelCaos As Boolean = False)
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer
Dim UI As Integer

For headingloop = eHeading.NORTH To eHeading.WEST
    nPos = Npclist(NpcIndex).pos
    If Npclist(NpcIndex).flags.Inmovilizado = 0 Or headingloop = Npclist(NpcIndex).char.Heading Then
        Call HeadtoPos(headingloop, nPos)
        If InMapBounds(nPos.Map, nPos.x, nPos.y) Then
            UI = MapData(nPos.Map, nPos.x, nPos.y).UserIndex
            If UI > 0 Then
                  If UserList(UI).flags.Muerto = 0 Then
                         '¿ES CRIMINAL?
                         If Not DelCaos Then
                            If Criminal(UI) Then
                                   If NpcAtacaUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(SendTarget.ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).char.Body, Npclist(NpcIndex).char.Head, headingloop)
                                   End If
                                   Exit Sub
                            ElseIf Npclist(NpcIndex).flags.AttackedBy = UserList(UI).name _
                                      And Not Npclist(NpcIndex).flags.Follow Then
                                  
                                  If NpcAtacaUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(SendTarget.ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).char.Body, Npclist(NpcIndex).char.Head, headingloop)
                                  End If
                                  Exit Sub
                            End If
                        Else
                            If Not Criminal(UI) Then
                                   
                                   If NpcAtacaUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(SendTarget.ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).char.Body, Npclist(NpcIndex).char.Head, headingloop)
                                   End If
                                   Exit Sub
                            ElseIf Npclist(NpcIndex).flags.AttackedBy = UserList(UI).name _
                                      And Not Npclist(NpcIndex).flags.Follow Then
                                  
                                  If NpcAtacaUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(SendTarget.ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).char.Body, Npclist(NpcIndex).char.Head, headingloop)
                                  End If
                                  Exit Sub
                            End If
                        End If
                  End If
            End If
        End If
    End If  'not inmovil
Next headingloop

Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub HostilMalvadoAI(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer
Dim UI As Integer
Dim NPCI As Integer
Dim atacoPJ As Boolean

atacoPJ = False

For headingloop = eHeading.NORTH To eHeading.WEST
    nPos = Npclist(NpcIndex).pos
    If Npclist(NpcIndex).flags.Inmovilizado = 0 Or Npclist(NpcIndex).char.Heading = headingloop Then
        Call HeadtoPos(headingloop, nPos)
        If InMapBounds(nPos.Map, nPos.x, nPos.y) Then
            UI = MapData(nPos.Map, nPos.x, nPos.y).UserIndex
            NPCI = MapData(nPos.Map, nPos.x, nPos.y).NpcIndex
            If UI > 0 And Not atacoPJ Then
                If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Privilegios = PlayerType.User Then
                    atacoPJ = True
                    If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then
                        Call NpcLanzaUnSpell(NpcIndex, UI)
                    End If
                    If NpcAtacaUser(NpcIndex, MapData(nPos.Map, nPos.x, nPos.y).UserIndex) Then
                        Call ChangeNPCChar(SendTarget.ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).char.Body, Npclist(NpcIndex).char.Head, headingloop)
                    End If
                    Exit Sub
                End If
            ElseIf NPCI > 0 Then
                    If Npclist(NPCI).MaestroUser > 0 And Npclist(NPCI).flags.Paralizado = 0 Then
                        Call ChangeNPCChar(SendTarget.ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).char.Body, Npclist(NpcIndex).char.Head, headingloop)
                        Call SistemaCombate.NpcAtacaNpc(NpcIndex, NPCI, False)
                        Exit Sub
                    End If
            End If
        End If
    End If  'inmo
Next headingloop

Call RestoreOldMovement(NpcIndex)

End Sub


Private Sub HostilBuenoAI(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim headingloop As eHeading
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer
Dim UI As Integer
For headingloop = eHeading.NORTH To eHeading.WEST
    nPos = Npclist(NpcIndex).pos
    If Npclist(NpcIndex).flags.Inmovilizado = 0 Or Npclist(NpcIndex).char.Heading = headingloop Then
        Call HeadtoPos(headingloop, nPos)
        If InMapBounds(nPos.Map, nPos.x, nPos.y) Then
            UI = MapData(nPos.Map, nPos.x, nPos.y).UserIndex
            If UI > 0 Then
                If UserList(UI).name = Npclist(NpcIndex).flags.AttackedBy Then
                    If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Privilegios = PlayerType.User Then
                            If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                              Call NpcLanzaUnSpell(NpcIndex, UI)
                            End If
                            If NpcAtacaUser(NpcIndex, UI) Then
                                Call ChangeNPCChar(SendTarget.ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).char.Body, Npclist(NpcIndex).char.Head, headingloop)
                            End If
                            Exit Sub
                    End If
                End If
            End If
        End If
    End If
Next headingloop

Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub IrUsuarioCercano(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer
Dim UI As Integer
Dim SignoNS As Integer
Dim SignoEO As Integer

If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
    Select Case Npclist(NpcIndex).char.Heading
        Case eHeading.NORTH
            SignoNS = -1
            SignoEO = 0
        Case eHeading.EAST
            SignoNS = 0
            SignoEO = 1
        Case eHeading.SOUTH
            SignoNS = 1
            SignoEO = 0
        Case eHeading.WEST
            SignoEO = -1
            SignoNS = 0
    End Select
    
    For y = Npclist(NpcIndex).pos.y To Npclist(NpcIndex).pos.y + SignoNS * RANGO_VISION_Y Step IIf(SignoNS = 0, 1, SignoNS)
        For x = Npclist(NpcIndex).pos.x To Npclist(NpcIndex).pos.x + SignoEO * RANGO_VISION_X Step IIf(SignoEO = 0, 1, SignoEO)
            
            If x >= MinXBorder And x <= MaxXBorder And y >= MinYBorder And y <= MaxYBorder Then
                   UI = MapData(Npclist(NpcIndex).pos.Map, x, y).UserIndex
                   If UI > 0 Then
                      If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Privilegios = PlayerType.User Then
                            If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                            Exit Sub
                      End If
                   End If
            End If
            
        Next x
    Next y
    
Else
    For y = Npclist(NpcIndex).pos.y - RANGO_VISION_Y To Npclist(NpcIndex).pos.y + RANGO_VISION_Y
        For x = Npclist(NpcIndex).pos.x - RANGO_VISION_X To Npclist(NpcIndex).pos.x + RANGO_VISION_X
            If x >= MinXBorder And x <= MaxXBorder And y >= MinYBorder And y <= MaxYBorder Then
                UI = MapData(Npclist(NpcIndex).pos.Map, x, y).UserIndex
                If UI > 0 Then
                     If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.Privilegios = PlayerType.User Then
                         If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                         tHeading = FindDirection(Npclist(NpcIndex).pos, UserList(MapData(Npclist(NpcIndex).pos.Map, x, y).UserIndex).pos)
                         Call MoveNPCChar(NpcIndex, tHeading)
                         Exit Sub
                     End If
                End If
            End If
        Next x
    Next y
End If

Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub SeguirAgresor(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer
Dim UI As Integer

Dim SignoNS As Integer
Dim SignoEO As Integer

If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
    Select Case Npclist(NpcIndex).char.Heading
        Case eHeading.NORTH
            SignoNS = -1
            SignoEO = 0
        Case eHeading.EAST
            SignoNS = 0
            SignoEO = 1
        Case eHeading.SOUTH
            SignoNS = 1
            SignoEO = 0
        Case eHeading.WEST
            SignoEO = -1
            SignoNS = 0
    End Select
    
    For y = Npclist(NpcIndex).pos.y To Npclist(NpcIndex).pos.y + SignoNS * RANGO_VISION_Y Step IIf(SignoNS = 0, 1, SignoNS)
        For x = Npclist(NpcIndex).pos.x To Npclist(NpcIndex).pos.x + SignoEO * RANGO_VISION_X Step IIf(SignoEO = 0, 1, SignoEO)

            If x >= MinXBorder And x <= MaxXBorder And y >= MinYBorder And y <= MaxYBorder Then
                UI = MapData(Npclist(NpcIndex).pos.Map, x, y).UserIndex
                If UI > 0 Then
                    If UserList(UI).name = Npclist(NpcIndex).flags.AttackedBy Then
                        If Npclist(NpcIndex).MaestroUser > 0 Then
                            If Not Criminal(Npclist(NpcIndex).MaestroUser) And Not Criminal(UI) And (UserList(Npclist(NpcIndex).MaestroUser).flags.Seguro Or UserList(Npclist(NpcIndex).MaestroUser).Faccion.ArmadaReal = 1) Then
                                Call SendData(SendTarget.ToIndex, Npclist(NpcIndex).MaestroUser, 0, "||La mascota no atacará a ciudadanos si eres miembro de la Armada Real o tienes el seguro activado" & FONTTYPE_INFO)
                                Npclist(NpcIndex).flags.AttackedBy = ""
                                Exit Sub
                            End If
                        End If
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.Privilegios = PlayerType.User Then
                             If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                                  Call NpcLanzaUnSpell(NpcIndex, UI)
                             End If
                             Exit Sub
                        End If
                    End If
                End If
            End If
        Next x
    Next y
Else
    For y = Npclist(NpcIndex).pos.y - RANGO_VISION_Y To Npclist(NpcIndex).pos.y + RANGO_VISION_Y
        For x = Npclist(NpcIndex).pos.x - RANGO_VISION_X To Npclist(NpcIndex).pos.x + RANGO_VISION_X
            If x >= MinXBorder And x <= MaxXBorder And y >= MinYBorder And y <= MaxYBorder Then
                UI = MapData(Npclist(NpcIndex).pos.Map, x, y).UserIndex
                If UI > 0 Then
                    If UserList(UI).name = Npclist(NpcIndex).flags.AttackedBy Then
                        If Npclist(NpcIndex).MaestroUser > 0 Then
                            If Not Criminal(Npclist(NpcIndex).MaestroUser) And Not Criminal(UI) And (UserList(Npclist(NpcIndex).MaestroUser).flags.Seguro Or UserList(Npclist(NpcIndex).MaestroUser).Faccion.ArmadaReal = 1) Then
                                Call SendData(SendTarget.ToIndex, Npclist(NpcIndex).MaestroUser, 0, "||La mascota no atacará a ciudadanos si eres miembro de la Armada Real o tienes el seguro activado" & FONTTYPE_INFO)
                                Npclist(NpcIndex).flags.AttackedBy = ""
                                Call FollowAmo(NpcIndex)
                                Exit Sub
                            End If
                        End If
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.Privilegios = PlayerType.User Then
                             If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                                  Call NpcLanzaUnSpell(NpcIndex, UI)
                             End If
                             tHeading = FindDirection(Npclist(NpcIndex).pos, UserList(MapData(Npclist(NpcIndex).pos.Map, x, y).UserIndex).pos)
                             Call MoveNPCChar(NpcIndex, tHeading)
                             Exit Sub
                        End If
                    End If
                End If
            End If
        Next x
    Next y
End If
Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)

If Npclist(NpcIndex).MaestroUser = 0 Then
    Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
    Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
    Npclist(NpcIndex).flags.AttackedBy = ""
End If

End Sub

Private Sub PersigueCiudadano(ByVal NpcIndex As Integer)
Dim UI As Integer
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer
For y = Npclist(NpcIndex).pos.y - RANGO_VISION_Y To Npclist(NpcIndex).pos.y + RANGO_VISION_Y
    For x = Npclist(NpcIndex).pos.x - RANGO_VISION_X To Npclist(NpcIndex).pos.x + RANGO_VISION_X
        If x >= MinXBorder And x <= MaxXBorder And y >= MinYBorder And y <= MaxYBorder Then
           UI = MapData(Npclist(NpcIndex).pos.Map, x, y).UserIndex
           If UI > 0 Then
                If Not Criminal(UI) Then
                   If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.Privilegios = PlayerType.User Then
                        If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                              Call NpcLanzaUnSpell(NpcIndex, UI)
                        End If
                        tHeading = FindDirection(Npclist(NpcIndex).pos, UserList(MapData(Npclist(NpcIndex).pos.Map, x, y).UserIndex).pos)
                        Call MoveNPCChar(NpcIndex, tHeading)
                        Exit Sub
                   End If
                End If
           End If
        End If
    Next x
Next y

Call RestoreOldMovement(NpcIndex)

End Sub


Private Sub PersigueCriminal(ByVal NpcIndex As Integer)
Dim UI As Integer
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer
Dim SignoNS As Integer
Dim SignoEO As Integer

If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
    Select Case Npclist(NpcIndex).char.Heading
        Case eHeading.NORTH
            SignoNS = -1
            SignoEO = 0
        Case eHeading.EAST
            SignoNS = 0
            SignoEO = 1
        Case eHeading.SOUTH
            SignoNS = 1
            SignoEO = 0
        Case eHeading.WEST
            SignoEO = -1
            SignoNS = 0
    End Select
    
    For y = Npclist(NpcIndex).pos.y To Npclist(NpcIndex).pos.y + SignoNS * RANGO_VISION_Y Step IIf(SignoNS = 0, 1, SignoNS)
        For x = Npclist(NpcIndex).pos.x To Npclist(NpcIndex).pos.x + SignoEO * RANGO_VISION_X Step IIf(SignoEO = 0, 1, SignoEO)


            If x >= MinXBorder And x <= MaxXBorder And y >= MinYBorder And y <= MaxYBorder Then
               UI = MapData(Npclist(NpcIndex).pos.Map, x, y).UserIndex
               If UI > 0 Then
                    If Criminal(UI) Then
                       If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.Privilegios = PlayerType.User Then
                            If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                                  Call NpcLanzaUnSpell(NpcIndex, UI)
                            End If
                            Exit Sub
                       End If
                    End If
               End If
            End If


        Next x
    Next y
Else
    For y = Npclist(NpcIndex).pos.y - RANGO_VISION_Y To Npclist(NpcIndex).pos.y + RANGO_VISION_Y
        For x = Npclist(NpcIndex).pos.x - RANGO_VISION_X To Npclist(NpcIndex).pos.x + RANGO_VISION_X
            If x >= MinXBorder And x <= MaxXBorder And y >= MinYBorder And y <= MaxYBorder Then
               UI = MapData(Npclist(NpcIndex).pos.Map, x, y).UserIndex
               If UI > 0 Then
                    If Criminal(UI) Then
                       If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.Privilegios = PlayerType.User Then
                            If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                                  Call NpcLanzaUnSpell(NpcIndex, UI)
                            End If
                            If Npclist(NpcIndex).flags.Inmovilizado = 1 Then Exit Sub
                            tHeading = FindDirection(Npclist(NpcIndex).pos, UserList(MapData(Npclist(NpcIndex).pos.Map, x, y).UserIndex).pos)
                            Call MoveNPCChar(NpcIndex, tHeading)
                            Exit Sub
                       End If
                    End If
               End If
            End If
        Next x
    Next y
End If
Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub SeguirAmo(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer
Dim UI As Integer
For y = Npclist(NpcIndex).pos.y - 10 To Npclist(NpcIndex).pos.y + 10
    For x = Npclist(NpcIndex).pos.x - 10 To Npclist(NpcIndex).pos.x + 10
        If x >= MinXBorder And x <= MaxXBorder And y >= MinYBorder And y <= MaxYBorder Then
            If Npclist(NpcIndex).Target = 0 And Npclist(NpcIndex).TargetNPC = 0 Then
                UI = MapData(Npclist(NpcIndex).pos.Map, x, y).UserIndex
                If UI > 0 Then
                   If UserList(UI).flags.Muerto = 0 _
                   And UserList(UI).flags.Invisible = 0 _
                   And UserList(UI).flags.Oculto = 0 _
                   And UI = Npclist(NpcIndex).MaestroUser _
                   And Distancia(Npclist(NpcIndex).pos, UserList(UI).pos) > 3 Then
                        tHeading = FindDirection(Npclist(NpcIndex).pos, UserList(MapData(Npclist(NpcIndex).pos.Map, x, y).UserIndex).pos)
                        Call MoveNPCChar(NpcIndex, tHeading)
                        Exit Sub
                   End If
                End If
            End If
        End If
    Next x
Next y

Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub AiNpcAtacaNpc(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer
Dim NI As Integer
Dim bNoEsta As Boolean

Dim SignoNS As Integer
Dim SignoEO As Integer

If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
    Select Case Npclist(NpcIndex).char.Heading
        Case eHeading.NORTH
            SignoNS = -1
            SignoEO = 0
        Case eHeading.EAST
            SignoNS = 0
            SignoEO = 1
        Case eHeading.SOUTH
            SignoNS = 1
            SignoEO = 0
        Case eHeading.WEST
            SignoEO = -1
            SignoNS = 0
    End Select
    
    For y = Npclist(NpcIndex).pos.y To Npclist(NpcIndex).pos.y + SignoNS * RANGO_VISION_Y Step IIf(SignoNS = 0, 1, SignoNS)
        For x = Npclist(NpcIndex).pos.x To Npclist(NpcIndex).pos.x + SignoEO * RANGO_VISION_X Step IIf(SignoEO = 0, 1, SignoEO)
            If x >= MinXBorder And x <= MaxXBorder And y >= MinYBorder And y <= MaxYBorder Then
               NI = MapData(Npclist(NpcIndex).pos.Map, x, y).NpcIndex
               If NI > 0 Then
                    If Npclist(NpcIndex).TargetNPC = NI Then
                         bNoEsta = True
                         If Npclist(NpcIndex).Numero = ELEMENTALFUEGO Then
                             Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)
                             If Npclist(NI).NPCtype = DRAGON Then
                                Npclist(NI).CanAttack = 1
                                Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)
                             End If
                         Else
                            'aca verificamosss la distancia de ataque
                            
                                Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)
                            
                         End If
                         Exit Sub
                    End If
               End If
            End If
        Next x
    Next y
Else
    For y = Npclist(NpcIndex).pos.y - RANGO_VISION_Y To Npclist(NpcIndex).pos.y + RANGO_VISION_Y
        For x = Npclist(NpcIndex).pos.x - RANGO_VISION_Y To Npclist(NpcIndex).pos.x + RANGO_VISION_Y
            If x >= MinXBorder And x <= MaxXBorder And y >= MinYBorder And y <= MaxYBorder Then
               NI = MapData(Npclist(NpcIndex).pos.Map, x, y).NpcIndex
               If NI > 0 Then
                    If Npclist(NpcIndex).TargetNPC = NI Then
                         bNoEsta = True
                         If Npclist(NpcIndex).Numero = ELEMENTALFUEGO Then
                             Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)
                             If Npclist(NI).NPCtype = DRAGON Then
                                Npclist(NI).CanAttack = 1
                                Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)
                             End If
                         Else
                            'aca verificamosss la distancia de ataque
                            
                                Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)
                     
                         End If
                         If Npclist(NpcIndex).flags.Inmovilizado = 1 Then Exit Sub
                         tHeading = FindDirection(Npclist(NpcIndex).pos, Npclist(MapData(Npclist(NpcIndex).pos.Map, x, y).NpcIndex).pos)
                         Call MoveNPCChar(NpcIndex, tHeading)
                         Exit Sub
                    End If
               End If
            End If
        Next x
    Next y
End If

If Not bNoEsta Then
    If Npclist(NpcIndex).MaestroUser > 0 Then
        Call FollowAmo(NpcIndex)
    Else
        Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
        Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
    End If
End If
    
End Sub

Function NPCAI(ByVal NpcIndex As Integer)
On Error GoTo ErrorHandler
        '<<<<<<<<<<< Ataques >>>>>>>>>>>>>>>>
        If Npclist(NpcIndex).MaestroUser = 0 Then
            'Busca a alguien para atacar
            '¿Es un guardia?
            If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                    Call GuardiasAI(NpcIndex)
            ElseIf Npclist(NpcIndex).NPCtype = eNPCType.Guardiascaos Then
                    Call GuardiasAI(NpcIndex, True)
            ElseIf Npclist(NpcIndex).Hostile And Npclist(NpcIndex).Stats.Alineacion <> 0 Then
                    Call HostilMalvadoAI(NpcIndex)
            ElseIf Npclist(NpcIndex).Hostile And Npclist(NpcIndex).Stats.Alineacion = 0 Then
                    Call HostilBuenoAI(NpcIndex)
            End If
        Else
            If False Then Exit Function
            'Evitamos que ataque a su amo, a menos
            'que el amo lo ataque.
            'Call HostilBuenoAI(NpcIndex)
        End If
        
        
        
        
        '<<<<<<<<<<<Movimiento>>>>>>>>>>>>>>>>
        Select Case Npclist(NpcIndex).Movement
            Case TipoAI.MueveAlAzar
                If Npclist(NpcIndex).flags.Inmovilizado = 1 Then Exit Function
                If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                    End If
                    Call PersigueCriminal(NpcIndex)
                ElseIf Npclist(NpcIndex).NPCtype = eNPCType.Guardiascaos Then
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                    End If
                    Call PersigueCiudadano(NpcIndex)
                Else
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                    End If
                End If
            'Va hacia el usuario cercano
            Case TipoAI.NpcMaloAtacaUsersBuenos
                Call IrUsuarioCercano(NpcIndex)
            'Va hacia el usuario que lo ataco(FOLLOW)
            Case TipoAI.NPCDEFENSA
                Call SeguirAgresor(NpcIndex)
            'Persigue criminales
            Case TipoAI.GuardiasAtacanCriminales
                Call PersigueCriminal(NpcIndex)
            Case TipoAI.SigueAmo
                If Npclist(NpcIndex).flags.Inmovilizado = 1 Then Exit Function
                Call SeguirAmo(NpcIndex)
                If RandomNumber(1, 12) = 3 Then
                    Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                End If
            Case TipoAI.NpcAtacaNpc
                Call AiNpcAtacaNpc(NpcIndex)
            Case TipoAI.NpcPathfinding
                If Npclist(NpcIndex).flags.Inmovilizado = 1 Then Exit Function
                If ReCalculatePath(NpcIndex) Then
                    Call PathFindingAI(NpcIndex)
                    'Existe el camino?
                    If Npclist(NpcIndex).PFINFO.NoPath Then 'Si no existe nos movemos al azar
                        'Move randomly
                        Call MoveNPCChar(NpcIndex, RandomNumber(eHeading.NORTH, eHeading.WEST))
                    End If
                Else
                    If Not PathEnd(NpcIndex) Then
                        Call FollowPath(NpcIndex)
                    Else
                        Npclist(NpcIndex).PFINFO.PathLenght = 0
                    End If
                End If

        End Select


Exit Function


ErrorHandler:
    Call LogError("NPCAI " & Npclist(NpcIndex).name & " " & Npclist(NpcIndex).MaestroUser & " " & Npclist(NpcIndex).MaestroNpc & " mapa:" & Npclist(NpcIndex).pos.Map & " x:" & Npclist(NpcIndex).pos.x & " y:" & Npclist(NpcIndex).pos.y & " Mov:" & Npclist(NpcIndex).Movement & " TargU:" & Npclist(NpcIndex).Target & " TargN:" & Npclist(NpcIndex).TargetNPC)
    Dim MiNPC As npc
    MiNPC = Npclist(NpcIndex)
    Call QuitarNPC(NpcIndex)
    Call ReSpawnNpc(MiNPC)
    
End Function


Function UserNear(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Returns True if there is an user adjacent to the npc position.
'#################################################################
UserNear = Not Int(Distance(Npclist(NpcIndex).pos.x, Npclist(NpcIndex).pos.y, UserList(Npclist(NpcIndex).PFINFO.TargetUser).pos.x, UserList(Npclist(NpcIndex).PFINFO.TargetUser).pos.y)) > 1
End Function

Function ReCalculatePath(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Returns true if we have to seek a new path
'#################################################################
If Npclist(NpcIndex).PFINFO.PathLenght = 0 Then
    ReCalculatePath = True
ElseIf Not UserNear(NpcIndex) And Npclist(NpcIndex).PFINFO.PathLenght = Npclist(NpcIndex).PFINFO.CurPos - 1 Then
    ReCalculatePath = True
End If
End Function

Function SimpleAI(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Old Ore4 AI function
'#################################################################
Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer

For y = Npclist(NpcIndex).pos.y - 5 To Npclist(NpcIndex).pos.y + 5    'Makes a loop that looks at
    For x = Npclist(NpcIndex).pos.x - 5 To Npclist(NpcIndex).pos.x + 5   '5 tiles in every direction
           'Make sure tile is legal
            If x > MinXBorder And x < MaxXBorder And y > MinYBorder And y < MaxYBorder Then
                'look for a user
                If MapData(Npclist(NpcIndex).pos.Map, x, y).UserIndex > 0 Then
                    'Move towards user
                    tHeading = FindDirection(Npclist(NpcIndex).pos, UserList(MapData(Npclist(NpcIndex).pos.Map, x, y).UserIndex).pos)
                    MoveNPCChar NpcIndex, tHeading
                    'Leave
                    Exit Function
                End If
            End If
    Next x
Next y

End Function

Function PathEnd(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock
'Returns if the npc has arrived to the end of its path
'#################################################################
PathEnd = Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.PathLenght
End Function

Function FollowPath(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock
'Moves the npc.
'#################################################################

Dim tmpPos As WorldPos
Dim tHeading As Byte

tmpPos.Map = Npclist(NpcIndex).pos.Map
tmpPos.x = Npclist(NpcIndex).PFINFO.Path(Npclist(NpcIndex).PFINFO.CurPos).y ' invertí las coordenadas
tmpPos.y = Npclist(NpcIndex).PFINFO.Path(Npclist(NpcIndex).PFINFO.CurPos).x

'Debug.Print "(" & tmpPos.X & "," & tmpPos.Y & ")"

tHeading = FindDirection(Npclist(NpcIndex).pos, tmpPos)

MoveNPCChar NpcIndex, tHeading

Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.CurPos + 1

End Function

Function PathFindingAI(ByVal NpcIndex As Integer) As Boolean

Dim nPos As WorldPos
Dim headingloop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer

For y = Npclist(NpcIndex).pos.y - 10 To Npclist(NpcIndex).pos.y + 10    'Makes a loop that looks at
     For x = Npclist(NpcIndex).pos.x - 10 To Npclist(NpcIndex).pos.x + 10   '5 tiles in every direction

         'Make sure tile is legal
         If x > MinXBorder And x < MaxXBorder And y > MinYBorder And y < MaxYBorder Then
         
             'look for a user
             If MapData(Npclist(NpcIndex).pos.Map, x, y).UserIndex > 0 Then
                 'Move towards user
                  Dim tmpUserIndex As Integer
                  tmpUserIndex = MapData(Npclist(NpcIndex).pos.Map, x, y).UserIndex
                  If UserList(tmpUserIndex).flags.Muerto = 0 And UserList(tmpUserIndex).flags.Invisible = 0 And UserList(tmpUserIndex).flags.Oculto = 0 And UserList(tmpUserIndex).flags.Privilegios = PlayerType.User Then
                    'We have to invert the coordinates, this is because
                    'ORE refers to maps in converse way of my pathfinding
                    'routines.
                    Npclist(NpcIndex).PFINFO.Target.x = UserList(tmpUserIndex).pos.y
                    Npclist(NpcIndex).PFINFO.Target.y = UserList(tmpUserIndex).pos.x 'ops!
                    Npclist(NpcIndex).PFINFO.TargetUser = tmpUserIndex
                    Call SeekPath(NpcIndex)
                    Exit Function
                  End If
             End If
             
         End If
              
     Next x
 Next y
End Function


Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Or Not UserList(UserIndex).flags.Privilegios = PlayerType.User Then Exit Sub

Dim k As Integer
k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
Call NpcLanzaSpellSobreUser(NpcIndex, UserIndex, Npclist(NpcIndex).Spells(k))

End Sub


Sub NpcLanzaUnSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer)

Dim k As Integer
    k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNPC, Npclist(NpcIndex).Spells(k))

End Sub


