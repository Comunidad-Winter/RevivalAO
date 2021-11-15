Attribute VB_Name = "RETOS"
    Public Sub ComensarDuelo(ByVal userindex As Integer, ByVal TIndex As Integer)
    UserList(userindex).flags.EstaDueleando = True
    UserList(userindex).flags.Oponente = TIndex
    UserList(TIndex).flags.EstaDueleando = True
    Call WarpUserChar(TIndex, 78, 41, 50)
    UserList(TIndex).flags.Oponente = userindex
    Call WarpUserChar(userindex, 78, 60, 50)
    Call SendData(toAll, 0, 0, "||Retos: " & UserList(TIndex).name & " y " & UserList(userindex).name & " van a competir en un Reto." & FONTTYPE_RETOS)
    Retos1 = UserList(TIndex).name
    Retos2 = UserList(userindex).name
    End Sub
    Public Sub ResetDuelo(ByVal userindex As Integer, ByVal TIndex As Integer)
    On Error GoTo errrorxaoo
    UserList(userindex).flags.EsperandoDuelo = False
    UserList(userindex).flags.Oponente = 0
    UserList(userindex).flags.EstaDueleando = False
     Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                  FuturePos.Map = 1
                  FuturePos.x = 50: FuturePos.Y = 50
                  Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(userindex, NuevaPos.Map, NuevaPos.x, NuevaPos.Y, True)

                  Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(TIndex, NuevaPos.Map, NuevaPos.x, NuevaPos.Y, True)
    UserList(TIndex).flags.EsperandoDuelo = False
    UserList(TIndex).flags.Oponente = 0
    UserList(TIndex).flags.EstaDueleando = False
errrorxaoo:
    End Sub
    Public Sub TerminarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer)
    On Error GoTo errorxao
    Call SendData(toAll, Ganador, 0, "||Retos: " & UserList(Ganador).name & " venció a " & UserList(Perdedor).name & " en un reto." & FONTTYPE_RETOS)
 
    If UserList(Perdedor).Stats.GLD >= 1000000 Then
    UserList(Perdedor).Stats.GLD = UserList(Perdedor).Stats.GLD - 1000000
    UserList(Ganador).Stats.GLD = UserList(Ganador).Stats.GLD + 1000000
    End If

    UserList(Ganador).Stats.PuntosRetos = UserList(Ganador).Stats.PuntosRetos + 1
    Call CompruebaRetos(Ganador)
    Call SendUserStatsBox(Perdedor)
    Call SendUserStatsBox(Ganador)
    Call ResetDuelo(Ganador, Perdedor)
errorxao:
    End Sub
    Public Sub DesconectarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer)
    On Error GoTo errorxaoo
    Call SendData(toAll, Ganador, 0, "||Retos: El reto ha sido cancelado por la desconexión de " & UserList(Perdedor).name & "." & FONTTYPE_RETOS)
    Call ResetDuelo(Ganador, Perdedor)
errorxaoo:
    End Sub
