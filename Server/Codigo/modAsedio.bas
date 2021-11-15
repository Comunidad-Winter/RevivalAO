Attribute VB_Name = "modAsedio"
Option Explicit
Public Const ItemMuralla As Integer = 4297
Public Const ReyNPC As Integer = 1018
Public Const MurallaNPC As Integer = 1019

Private Const Muralla_Max As Integer = 32748
Private Const Muralla_Medio As Integer = 32749
Private Const Muralla_Min As Integer = 32750

Private Muralla_Position(1 To 4) As tAsedioPos

Public ReyIndex As Integer

Public Muralla(0 To 6, 1 To 4) As Integer

Public UserAsedio() As Integer
Public Enum Equipos
    Verde = 1
    Negro = 2
    Azul = 3
    Rojo = 4
End Enum

Public Enum AStatus
    Finalizada = 0
    Inscripcion = 1
    Curso = 2
End Enum

Public Type tAsedio
    Estado As AStatus
    MaxSlots As Integer
    Slots As Integer
    Costo As Long
    Premio As Long
    Tiempo As Long
End Type

Public Type flagsAsedio
    Participando As Boolean
    Slot As Integer
    Team As Integer
End Type

Private Type tAsedioPos
    Map As Byte
    x As Byte
    Y As Byte
End Type

Public ReyTeam As Byte

Public Asedio As tAsedio
Public Sub WarpUserCharX(ByVal userindex As Integer, ByVal mapa As Integer, ByVal x As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)
    Dim NuevaPos As WorldPos
    Dim FuturePos As WorldPos
    FuturePos.Map = mapa
    FuturePos.x = x
    FuturePos.Y = Y
    Call ClosestLegalPos(FuturePos, NuevaPos)
          
    If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(userindex, NuevaPos.Map, NuevaPos.x, NuevaPos.Y, FX)
End Sub
Public Sub Iniciar_Asedio(ByVal userindex As Integer, ByVal MaxSlot As Integer, ByVal Costo As Long, ByVal Tiempo As Long)
    
    Select Case Asedio.Estado
        Case AStatus.Inscripcion
            If userindex > 0 Then Call SendData(SendTarget.toIndex, userindex, 0, "||¡El evento esta en su inscripcion! ~255~255~255~1~0~")
            Exit Sub
        Case AStatus.Curso
            If userindex > 0 Then Call SendData(SendTarget.toIndex, userindex, 0, "||¡El evento ya ha comenzado! ~255~255~255~1~0~")
            Exit Sub
    End Select
    
    If MaxSlot Mod 4 <> 0 Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||El maximo slots que este evento tolera son multiplos de 4, ejemplo maxslot de 4, 8, 12, 16, 20, 24, 28, 32, 36, 40, etc. ~255~255~255~1~0~")
        Exit Sub
    End If
    
    If Tiempo < 5 Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||El tiempo minimo es de 30 minutos ~255~255~255~1~0~")
        Exit Sub
    End If
    
    Asedio.MaxSlots = MaxSlot
    Asedio.Slots = 0
    Asedio.Costo = Costo
    Asedio.Premio = 7000000
    Asedio.Estado = AStatus.Curso
    Asedio.Tiempo = Tiempo
    
    If ReyIndex > 0 Then
        If Npclist(ReyIndex).Numero = ReyNPC Then
            Call QuitarNPC(ReyIndex)
        End If
    End If
    
    With Muralla_Position(Equipos.Azul)
        .Map = 114
        .x = 14
        .Y = 53
    End With
    
    With Muralla_Position(Equipos.Verde)
        .Map = 114
        .x = 47
        .Y = 72
    End With
    
    With Muralla_Position(Equipos.Rojo)
        .Map = 114
        .x = 79
        .Y = 48
    End With
    
    With Muralla_Position(Equipos.Negro)
        .Map = 114
        .x = 47
        .Y = 27
    End With
    

    Dim i As Byte
    Dim j As Byte
    Dim Position As WorldPos
    
    For i = 0 To 6
        For j = 1 To 4
            Position.Map = Muralla_Position(j).Map
            Position.x = Muralla_Position(j).x + i
            Position.Y = Muralla_Position(j).Y
            Muralla(i, j) = SpawnNpc(MurallaNPC, Position, False, False)
            Npclist(Muralla(i, j)).MurallaEquipo = j
            Npclist(Muralla(i, j)).MurallaIndex = i
            Call CalcularGrafico(Muralla(i, j))
        Next j
    Next i
    
    Dim PosRey As WorldPos
    PosRey.Map = 115
    PosRey.x = 46
    PosRey.Y = 58
    
    ReyIndex = SpawnNpc(ReyNPC, PosRey, False, False)
    
    ReDim UserAsedio(1 To Asedio.MaxSlots, 1 To 4) As Integer
    
    Call SendData(SendTarget.toAll, 0, 0, "||Se ha dado comienzo al Evento Asedio, para ingresar /ASEDIO ~255~255~255~1~0~")
    Call SendData(SendTarget.toAll, 0, 0, "||Costo inscripcion: " & Asedio.Costo & "~255~255~255~1~0~")
    Call SendData(SendTarget.toAll, 0, 0, "||Tiempo: " & Asedio.Tiempo & "~255~255~255~1~0~")
     Call SendData(SendTarget.toAll, 0, 0, "TW48")
    
End Sub
Public Sub Inscribir_Asedio(ByVal userindex As Integer)
        If UserList(userindex).flags.EstaDueleando1 = True Then Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes ir a torneo estando plantes!." & FONTTYPE_WARNING): Exit Sub
        
        If userindex = Team.Pj1 Or userindex = Team.Pj2 Then Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes participar en eventos si esperas retos!!!" & FONTTYPE_INFO): Exit Sub
        
        If UserList(userindex).pos.Map = 66 Then Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes ir a guerra estando en duelos." & FONTTYPE_WARNING): Exit Sub
        
        If UserList(userindex).pos.Map = 61 Then Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes ir a torneo estando en duelos." & FONTTYPE_WARNING): Exit Sub
        
        If UserList(userindex).pos.Map = 79 Then Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes ir a torneo estando en torneos." & FONTTYPE_WARNING): Exit Sub
        
        If UserList(userindex).pos.Map = 88 Then Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes ir a duelo estando en deatmatch." & FONTTYPE_WARNING): Exit Sub
        
        If UserList(userindex).pos.Map = 87 Then Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes ir a duelo estando en retos." & FONTTYPE_WARNING): Exit Sub
        
        If UserList(userindex).pos.Map = 78 Then Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes ir a torneo estando en retos." & FONTTYPE_WARNING): Exit Sub
        
        If UserList(userindex).pos.Map = 67 Then Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes ir a torneo estando en la carcel." & FONTTYPE_WARNING): Exit Sub
    
        If UserList(userindex).Asedio.Participando Then
            Call SendData(SendTarget.toIndex, userindex, 0, "||¡Ya estas en el evento!" & FONTTYPE_WARNING)
            Exit Sub
        End If
        
        If Asedio.Slots = Asedio.MaxSlots Then
            Call SendData(SendTarget.toIndex, userindex, 0, "||¡Cupo lleno!" & FONTTYPE_WARNING)
            Exit Sub
        End If
        
        If UserList(userindex).Stats.GLD - Asedio.Costo < 0 Then
            Call SendData(SendTarget.toIndex, userindex, 0, "||¡No tienes suficiente oro!" & FONTTYPE_WARNING)
            Exit Sub
        End If
        
        Static NumTeam As Integer
        Dim i As Long
        
        If NumTeam = 4 Then NumTeam = 0
        NumTeam = NumTeam + 1
        
        For i = 1 To Asedio.MaxSlots
            If UserAsedio(i, NumTeam) = 0 Then
                UserList(userindex).Asedio.Slot = i
                UserList(userindex).Asedio.Team = NumTeam
                UserAsedio(i, NumTeam) = userindex
                Exit For
            End If
        Next i
        
        Asedio.Slots = Asedio.Slots + 1 'Primero prueba el asedio con 100 slots
        UserList(userindex).Asedio.Participando = True
        UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - Asedio.Costo
        Asedio.Premio = Asedio.Premio + Asedio.Costo
        Call SendUserStatsBox(userindex)
        Call SendData(SendTarget.toIndex, userindex, 0, "||¡Te has inscripto! ¡Estas en el equipo " & NombreEquipo(NumTeam) & "!" & FONTTYPE_WARNING)
        
        Dim User_Position As tAsedioPos
        User_Position = PosBase(NumTeam)
        With User_Position
            Call WarpUserCharX(userindex, .Map, .x, .Y, True)
        End With
        Call EnviarAsedio("SSED" & Asedio.Tiempo)
        Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, val(userindex), UserList(userindex).char.Body, UserList(userindex).OrigChar.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
End Sub

Private Function PosBase(ByVal Team As Byte) As tAsedioPos
If ReyTeam = Team Then
    PosBase.Map = 115
    PosBase.x = 51
    PosBase.Y = 15
    Exit Function
End If

Select Case Team
    Case 1
        PosBase.Map = 114
        PosBase.x = 80
        PosBase.Y = 40
    Case 2
        PosBase.Map = 114
        PosBase.x = 14
        PosBase.Y = 60
    Case 3
        PosBase.Map = 114
        PosBase.x = 48
        PosBase.Y = 17
    Case 4
        PosBase.Map = 114
        PosBase.x = 48
        PosBase.Y = 82
End Select
End Function
Private Function NombreEquipo(ByVal Team As Byte) As String
Select Case Team
    Case Equipos.Azul
        NombreEquipo = "Azul"
    Case Equipos.Negro
        NombreEquipo = "Negro"
    Case Equipos.Rojo
        NombreEquipo = "Rojo"
    Case Equipos.Verde
        NombreEquipo = "Verde"
End Select
End Function
Public Sub MuereUser(ByVal userindex As Integer)
UserList(userindex).flags.Muerto = 0
UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
Call DarCuerpoDesnudo(userindex)

'[MaTeO 9]
Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, val(userindex), UserList(userindex).char.Body, UserList(userindex).OrigChar.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
'[/MaTeO 9]
         
Call SendUserStatsBox(userindex)

Dim User_Position As tAsedioPos
User_Position = PosBase(UserList(userindex).Asedio.Team)
With User_Position
    Call WarpUserCharX(userindex, .Map, .x, .Y, True)
End With
End Sub
Public Sub MuereRey(ByVal userindex As Integer)
If userindex > 0 Then
    With UserList(userindex)
        If .Asedio.Team > 0 Then
            Select Case .Asedio.Team
                Case Equipos.Negro
                    Npclist(ReyIndex).char.Body = 527
                Case Equipos.Verde
                    Npclist(ReyIndex).char.Body = 528
                Case Equipos.Azul
                    Npclist(ReyIndex).char.Body = 529
                Case Equipos.Rojo
                    Npclist(ReyIndex).char.Body = 530
            End Select
            Npclist(ReyIndex).Stats.MinHP = Npclist(ReyIndex).Stats.MaxHP
            Call ChangeNPCChar(SendTarget.ToMap, 0, Npclist(ReyIndex).pos.Map, ReyIndex, Npclist(ReyIndex).char.Body, Npclist(ReyIndex).char.Head, Npclist(ReyIndex).char.Heading)
            ReyTeam = .Asedio.Team
            Call LogAsedio("El rey ahora es del equipo " & ReyTeam)
        End If
    End With
    Call EnviarAsedio("||¡El rey ahora es del equipo " & NombreEquipo(ReyTeam) & "!~255~255~255~1~0~")
End If
End Sub
Public Sub DoTimerAsedio()
    If Asedio.Estado <> AStatus.Curso Then Exit Sub
    Asedio.Tiempo = Asedio.Tiempo - 1
    If Asedio.Tiempo = 0 Then
        'If ReyTeam = 0 Then
           ' Call EnviarAsedio("||Se ha finalizado el tiempo del evento, pero al no tener un ganador se agregan 5 minutos." & FONTTYPE_WARNING)
           ' Asedio.Tiempo = Asedio.Tiempo + 5
      '  Else
            Call SendData(SendTarget.toAll, 0, 0, "||¡Ha finalizado el evento y el ganador es el equipo " & NombreEquipo(ReyTeam) & "!" & FONTTYPE_WARNING)
            Dim i As Long
            Dim j As Long
            Dim Participantes As Long
            For i = 1 To Asedio.MaxSlots

                If UserAsedio(i, ReyTeam) > 0 Then
                    If UserList(UserAsedio(i, ReyTeam)).Asedio.Team = ReyTeam Then
                        Participantes = Participantes + 1
                    End If
                End If
            Next i
            Dim PremioxP As Long
            If Participantes <> 0 Then
                PremioxP = Asedio.Premio / Participantes
            End If
            For i = 1 To Asedio.MaxSlots
                For j = 1 To 4
                    If UserAsedio(i, j) > 0 Then
                       ' Call WarpUserCharX(UserAsedio(i, j), 1, 50, 50, False) 'PRUEBALO
                         Call LogAsedio("Damos premio a equipo: " & ReyTeam)
                        If j = ReyTeam Then
                            UserList(UserAsedio(i, j)).Stats.GLD = UserList(UserAsedio(i, j)).Stats.GLD + PremioxP
                            Call SendData(SendTarget.toIndex, UserAsedio(i, j), 0, "||Has ganado " & PremioxP & " de oro " & FONTTYPE_INFO)
                            UserList(UserAsedio(i, j)).Stats.PuntosCanje = UserList(UserAsedio(i, j)).Stats.PuntosCanje + 1
                            Call SendData(SendTarget.toIndex, UserAsedio(i, j), 0, "||Has ganado 1 punto de canje." & FONTTYPE_INFO)
                            Call SendUserStatsBox(UserAsedio(i, j))
                        End If
                        Call ResetFlagsAsedio(UserAsedio(i, j))
                    End If
                Next j
            Next i
            
            If ReyIndex > 0 Then
                Call QuitarNPC(ReyIndex)
                ReyIndex = 0
            End If
            Asedio.Costo = 0
            Asedio.Estado = Finalizada
            Asedio.MaxSlots = 0
            Asedio.Tiempo = 0
            Asedio.Slots = 0
            Asedio.Premio = 0
            ReyTeam = 0
            
       'End If
    End If
    Call EnviarAsedio("SSED" & Asedio.Tiempo)
End Sub
Public Sub EnviarAsedio(ByRef rData As String)
    Call SendData(SendTarget.ToMap, 0, 114, rData)
    Call SendData(SendTarget.ToMap, 0, 115, rData)
    Debug.Print "Envio: " & rData
End Sub
Public Sub ResetFlagsAsedio(ByVal userindex As Integer)
With UserList(userindex).Asedio
    If .Slot <> 0 And .Team <> 0 Then
        UserAsedio(.Slot, .Team) = 0
    End If
    If .Participando Then
        Call WarpUserCharX(userindex, 1, 50, 50, False)
    End If
    .Participando = False
    .Slot = 0
    .Team = 0
    If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
        UserList(userindex).char.Body = ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).Ropaje
        Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, val(userindex), UserList(userindex).char.Body, UserList(userindex).OrigChar.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
    Else
        Call DarCuerpoDesnudo(userindex, False)
    End If
End With

End Sub
Public Sub CancelAsedio()
            Call SendData(SendTarget.toAll, 0, 0, "||El Asedio ha sido cancelado." & FONTTYPE_WARNING)
            Dim i As Long
            Dim j As Long
            For i = 1 To Asedio.MaxSlots
                For j = 1 To 4
                    If UserAsedio(i, j) > 0 Then
                        'Call WarpUserCharX(UserAsedio(i, j), 1, 50, 50, False)
                        UserList(UserAsedio(i, j)).Stats.GLD = UserList(UserAsedio(i, j)).Stats.GLD + Asedio.Costo
                        Call SendUserStatsBox(UserAsedio(i, j))
                        Call ResetFlagsAsedio(UserAsedio(i, j))
                    End If
                Next j
            Next i
            
            For i = 0 To 6
                For j = 1 To 4
                    If Muralla(i, j) > 0 Then
                        Call QuitarNPC(Muralla(i, j))
                        Muralla(i, j) = 0
                    End If
                Next j
            Next i
            
            If ReyIndex > 0 Then
                Call QuitarNPC(ReyIndex)
                ReyIndex = 0
            End If
            
            Asedio.Costo = 0
            Asedio.Estado = Finalizada
            Asedio.MaxSlots = 0
            Asedio.Tiempo = 0
            Asedio.Slots = 0
            Asedio.Premio = 0
End Sub
Public Sub CalcularGrafico(ByVal NpcIndex As Integer)
Dim Vida As Long
Dim TeamNPC As Byte
TeamNPC = Npclist(NpcIndex).MurallaEquipo

If Npclist(NpcIndex).Stats.MinHP <= 0 Then
    MapData(Muralla_Position(TeamNPC).Map, Muralla_Position(TeamNPC).x + 3, Muralla_Position(TeamNPC).Y).OBJInfo.ObjIndex = 0
Else
    Vida = Fix(((Npclist(NpcIndex).Stats.MinHP / 100) / (Npclist(NpcIndex).Stats.MaxHP / 100)) * 100) + 1
    
    MapData(Muralla_Position(TeamNPC).Map, Muralla_Position(TeamNPC).x + 3, Muralla_Position(TeamNPC).Y).OBJInfo.ObjIndex = ItemMuralla + TeamNPC - 1
End If
Select Case Vida
    Case 80 To 101 'Intacta
        ObjData(ItemMuralla + TeamNPC - 1).GrhIndex = Muralla_Max
    Case 35 To 79 'Maso maso
        ObjData(ItemMuralla + TeamNPC - 1).GrhIndex = Muralla_Medio
    Case 1 To 34 'Casi destruida
        ObjData(ItemMuralla + TeamNPC - 1).GrhIndex = Muralla_Min
End Select
If Vida = 0 Then
    Call SendToAreaByPos(Muralla_Position(TeamNPC).Map, Muralla_Position(TeamNPC).x + 3, Muralla_Position(TeamNPC).Y, "BO" & Muralla_Position(TeamNPC).x + 3 & "," & Muralla_Position(TeamNPC).Y)
Else
    Call ModAreas.SendToAreaByPos(Muralla_Position(TeamNPC).Map, Muralla_Position(TeamNPC).x + 3, Muralla_Position(TeamNPC).Y, "HO" & ObjData(ItemMuralla + TeamNPC - 1).GrhIndex & "," & Muralla_Position(TeamNPC).x + 3 & "," & Muralla_Position(TeamNPC).Y)
End If
End Sub


