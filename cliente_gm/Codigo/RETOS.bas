Attribute VB_Name = "RETOS"
    Public Sub ComensarDuelo(ByVal UserIndex As Integer, ByVal TIndex As Integer)
    UserList(UserIndex).flags.EstaDueleando = True
    UserList(UserIndex).flags.Oponente = TIndex
    UserList(TIndex).flags.EstaDueleando = True
    Call WarpUserChar(TIndex, 63, 54, 26)
    UserList(TIndex).flags.Oponente = UserIndex
    Call WarpUserChar(UserIndex, 63, 71, 39)
    Call SendData(ToAll, 0, 0, "||Retos> " & UserList(TIndex).name & " y " & UserList(UserIndex).name & " van a competir en un Reto." & FONTTYPE_TALK)
    End Sub
    Public Sub ResetDuelo(ByVal UserIndex As Integer, ByVal TIndex As Integer)
    UserList(UserIndex).flags.EsperandoDuelo = False
    UserList(UserIndex).flags.Oponente = 0
    UserList(UserIndex).flags.EstaDueleando = False
    Call WarpUserChar(UserIndex, 1, 50, 50)
    Call WarpUserChar(TIndex, 1, 51, 51)
    UserList(TIndex).flags.EsperandoDuelo = False
    UserList(TIndex).flags.Oponente = 0
    UserList(TIndex).flags.EstaDueleando = False
    End Sub
    Public Sub TerminarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer)
    Call SendData(ToAll, Ganador, 0, "||Retos> " & UserList(Ganador).name & " venció a " & UserList(Perdedor).name & " en un reto." & FONTTYPE_TALK)
    UserList(Perdedor).Stats.GLD = UserList(Perdedor).Stats.GLD - 200000
    UserList(Ganador).Stats.GLD = UserList(Ganador).Stats.GLD + 200000
    Call SendUserStatsBox(Perdedor)
    Call SendUserStatsBox(Ganador)
    Call ResetDuelo(Ganador, Perdedor)
    End Sub
    Public Sub DesconectarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer)
    Call SendData(ToAll, Ganador, 0, "||Retos> El reto ha sido cancelado por la desconexión de " & UserList(Perdedor).name & "." & FONTTYPE_TALK)
    Call ResetDuelo(Ganador, Perdedor)
    End Sub
