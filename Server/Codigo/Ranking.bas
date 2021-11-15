Attribute VB_Name = "ModRanking"
Public Sub EnviaRank(ByVal UserIndex As Integer)
   SendData SendTarget.toIndex, UserIndex, 0, "BINMODEPT" & _
            Ranking.MaxOro.UserName _
            & "," & Ranking.MaxOro.value & _
            "," & Ranking.MaxTrofeos.UserName _
            & "," & Ranking.MaxTrofeos.value _
            & "," & Ranking.MaxUsuariosMatados.UserName _
            & "," & Ranking.MaxUsuariosMatados.value _
            & "," & Ranking.MaxTorneos.UserName _
            & "," & Ranking.MaxTorneos.value _
            & "," & Ranking.MaxDeaths.UserName _
            & "," & Ranking.MaxDeaths.value _
            & "," & Ranking.MaxRetos.UserName _
            & "," & Ranking.MaxRetos.value _
            & "," & Ranking.MaxDuelos.UserName _
            & "," & Ranking.MaxDuelos.value _
            & "," & Ranking.MaxPlantes.UserName _
            & "," & Ranking.MaxPlantes.value
End Sub
Public Sub EnviaPuntos(ByVal UserIndex As Integer)
 SendData SendTarget.toIndex, UserIndex, 0, "WETA" & _
            UserList(UserIndex).Stats.PuntosDeath _
            & "," & UserList(UserIndex).Stats.PuntosDuelos & _
            "," & UserList(UserIndex).Stats.PuntosPlante _
            & "," & UserList(UserIndex).Stats.PuntosRetos _
            & "," & UserList(UserIndex).Stats.PuntosTorneo _
            & "," & UserList(UserIndex).Stats.PuntosCanje
End Sub
Public Sub CompruebaOro(ByVal UserIndex As Integer)
' actualiza el ranking de oro si el usuario tiene mas oro que el mayor del ranking
If UserList(UserIndex).Stats.GLD > Ranking.MaxOro.value And UserList(UserIndex).flags.Privilegios = PlayerType.User Then
Ranking.MaxOro.value = UserList(UserIndex).Stats.GLD
Ranking.MaxOro.UserName = UserList(UserIndex).name
End If
End Sub
Public Sub CompruebaTrofeos(ByVal UserIndex As Integer)
' actualiza el ranking de trofeos si el usuario tiene mas trofeos que el mayor del ranking
If UserList(UserIndex).Stats.TrofOro > Ranking.MaxTrofeos.value And UserList(UserIndex).flags.Privilegios = PlayerType.User Then
Ranking.MaxTrofeos.value = UserList(UserIndex).Stats.TrofOro
Ranking.MaxTrofeos.UserName = UserList(UserIndex).name
End If
End Sub
Public Sub CompruebaUserDies(ByVal UserIndex As Integer)
' actualiza el ranking de muertes si el usuario tiene mas muertes que el mayor del ranking
If UserList(UserIndex).Stats.UsuariosMatados > Ranking.MaxUsuariosMatados.value And UserList(UserIndex).flags.Privilegios = PlayerType.User Then
Ranking.MaxUsuariosMatados.value = UserList(UserIndex).Stats.UsuariosMatados
Ranking.MaxUsuariosMatados.UserName = UserList(UserIndex).name
End If
End Sub

Public Sub CompruebaDuelos(ByVal UserIndex As Integer)
' actualiza el ranking de duelos si el usuario tiene mas duelos que el mayor del ranking
If UserList(UserIndex).Stats.PuntosDuelos > Ranking.MaxDuelos.value And UserList(UserIndex).flags.Privilegios = PlayerType.User Then
Ranking.MaxDuelos.value = UserList(UserIndex).Stats.PuntosDuelos
Ranking.MaxDuelos.UserName = UserList(UserIndex).name
End If
End Sub
Public Sub CompruebaRetos(ByVal UserIndex As Integer)
' actualiza el ranking de duelos si el usuario tiene mas duelos que el mayor del ranking
If UserList(UserIndex).Stats.PuntosRetos > Ranking.MaxRetos.value And UserList(UserIndex).flags.Privilegios = PlayerType.User Then
Ranking.MaxRetos.value = UserList(UserIndex).Stats.PuntosRetos
Ranking.MaxRetos.UserName = UserList(UserIndex).name
End If
End Sub
Public Sub CompruebaPlantes(ByVal UserIndex As Integer)
' actualiza el ranking de plantes si el usuario tiene mas plantes que el mayor del ranking
If UserList(UserIndex).Stats.PuntosPlante > Ranking.MaxPlantes.value And UserList(UserIndex).flags.Privilegios = PlayerType.User Then
Ranking.MaxPlantes.value = UserList(UserIndex).Stats.PuntosPlante
Ranking.MaxPlantes.UserName = UserList(UserIndex).name
End If
End Sub
Public Sub CompruebaTorneos(ByVal UserIndex As Integer)
' actualiza el ranking de torneos si el usuario tiene mas torneos que el mayor del ranking
If UserList(UserIndex).Stats.PuntosTorneo > Ranking.MaxTorneos.value And UserList(UserIndex).flags.Privilegios = PlayerType.User Then
Ranking.MaxTorneos.value = UserList(UserIndex).Stats.PuntosTorneo
Ranking.MaxTorneos.UserName = UserList(UserIndex).name
End If
End Sub
Public Sub CompruebaDeaths(ByVal UserIndex As Integer)
' actualiza el ranking de deaths si el usuario tiene mas deaths que el mayor del ranking
If UserList(UserIndex).Stats.PuntosDeath > Ranking.MaxDeaths.value And UserList(UserIndex).flags.Privilegios = PlayerType.User Then
Ranking.MaxDeaths.value = UserList(UserIndex).Stats.PuntosDeath
Ranking.MaxDeaths.UserName = UserList(UserIndex).name
End If
End Sub
