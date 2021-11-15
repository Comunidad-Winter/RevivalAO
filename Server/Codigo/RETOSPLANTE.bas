Attribute VB_Name = "RETOSPLANTE"
    Public Sub ComensarDueloPlantes(ByVal userindex As Integer, ByVal TIndex As Integer)
    YaHayPlante = True
    UserList(userindex).flags.EstaDueleando1 = True
    UserList(userindex).flags.Oponente1 = TIndex
    UserList(TIndex).flags.EstaDueleando1 = True
    Call WarpUserChar(TIndex, 1, 75, 33)
    UserList(TIndex).flags.Oponente1 = userindex
    Call WarpUserChar(userindex, 1, 76, 33)
    Call SendData(toall, 0, 0, "||Plantes: " & UserList(TIndex).name & " y " & UserList(userindex).name & " van a competir en un Reto de plantes." & FONTTYPE_PLANTE)
    Plante1 = UserList(TIndex).name
    Plante2 = UserList(userindex).name
    End Sub
    Public Sub ResetDueloPlantes(ByVal userindex As Integer, ByVal TIndex As Integer)
    On Error GoTo errrorxaoo
    UserList(userindex).flags.EsperandoDuelo1 = False
    UserList(userindex).flags.Oponente1 = 0
    UserList(userindex).flags.EstaDueleando1 = False
        Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                  FuturePos.Map = 1
                  FuturePos.x = 50: FuturePos.y = 50
                  Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then Call WarpUserChar(userindex, NuevaPos.Map, NuevaPos.x, NuevaPos.y, True)
                    Call ClosestLegalPos(FuturePos, NuevaPos)
    If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then Call WarpUserChar(TIndex, NuevaPos.Map, NuevaPos.x, NuevaPos.y, True)
    UserList(TIndex).flags.EsperandoDuelo1 = False
    UserList(TIndex).flags.Oponente1 = 0
    UserList(TIndex).flags.EstaDueleando1 = False
    YaHayPlante = False
errrorxaoo:
YaHayPlante = False
    End Sub
    Public Sub TerminarDueloPlantes(ByVal Ganador As Integer, ByVal Perdedor As Integer)
    On Error GoTo errorxao
    Call SendData(toall, Ganador, 0, "||Plantes: " & UserList(Ganador).name & " venció a " & UserList(Perdedor).name & " en un reto de plantes." & FONTTYPE_PLANTE)
    If UserList(Perdedor).Stats.GLD >= 500000 Then
    UserList(Perdedor).Stats.GLD = UserList(Perdedor).Stats.GLD - 500000
    End If
    UserList(Ganador).Stats.GLD = UserList(Ganador).Stats.GLD + 500000
    UserList(Ganador).Stats.PuntosPlante = UserList(Ganador).Stats.PuntosPlante + 1
    Call CompruebaPlantes(Ganador)
    Call SendUserStatsBox(Perdedor)
    Call SendUserStatsBox(Ganador)
    Call ResetDueloPlantes(Ganador, Perdedor)
     YaHayPlante = False
errorxao:
YaHayPlante = False
    End Sub
    Public Sub DesconectarDueloPlantes(ByVal Ganador As Integer, ByVal Perdedor As Integer)
    On Error GoTo errorxaoo
    Call SendData(toall, Ganador, 0, "||Plantes: El reto ha sido cancelado por la desconexión de " & UserList(Perdedor).name & "." & FONTTYPE_PLANTE)
    Call ResetDueloPlantes(Ganador, Perdedor)
     YaHayPlante = False
errorxaoo:
    End Sub
