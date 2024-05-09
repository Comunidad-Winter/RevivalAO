Attribute VB_Name = "ModRanking"
Public Sub EnviaRank(ByVal userindex As Integer)
   SendData SendTarget.toIndex, userindex, 0, "BINMODEPT" & _
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
Public Sub EnviaPuntos(ByVal userindex As Integer)
 SendData SendTarget.toIndex, userindex, 0, "WETA" & _
            UserList(userindex).Stats.PuntosDeath _
            & "," & UserList(userindex).Stats.PuntosDuelos & _
            "," & UserList(userindex).Stats.PuntosPlante _
            & "," & UserList(userindex).Stats.PuntosRetos _
            & "," & UserList(userindex).Stats.PuntosTorneo _
            & "," & UserList(userindex).Stats.PuntosCanje
End Sub
Public Sub CompruebaOro(ByVal userindex As Integer)
' actualiza el ranking de oro si el usuario tiene mas oro que el mayor del ranking
If UserList(userindex).Stats.GLD > Ranking.MaxOro.value And UserList(userindex).flags.Privilegios = PlayerType.User Then
Ranking.MaxOro.value = UserList(userindex).Stats.GLD
Ranking.MaxOro.UserName = UserList(userindex).name
End If
End Sub
Public Sub CompruebaTrofeos(ByVal userindex As Integer)
' actualiza el ranking de trofeos si el usuario tiene mas trofeos que el mayor del ranking
If UserList(userindex).Stats.TrofOro > Ranking.MaxTrofeos.value And UserList(userindex).flags.Privilegios = PlayerType.User Then
Ranking.MaxTrofeos.value = UserList(userindex).Stats.TrofOro
Ranking.MaxTrofeos.UserName = UserList(userindex).name
End If
End Sub
Public Sub CompruebaUserDies(ByVal userindex As Integer)
' actualiza el ranking de muertes si el usuario tiene mas muertes que el mayor del ranking
If UserList(userindex).Stats.UsuariosMatados > Ranking.MaxUsuariosMatados.value And UserList(userindex).flags.Privilegios = PlayerType.User Then
Ranking.MaxUsuariosMatados.value = UserList(userindex).Stats.UsuariosMatados
Ranking.MaxUsuariosMatados.UserName = UserList(userindex).name
End If
End Sub

Public Sub CompruebaDuelos(ByVal userindex As Integer)
' actualiza el ranking de duelos si el usuario tiene mas duelos que el mayor del ranking
If UserList(userindex).Stats.PuntosDuelos > Ranking.MaxDuelos.value And UserList(userindex).flags.Privilegios = PlayerType.User Then
Ranking.MaxDuelos.value = UserList(userindex).Stats.PuntosDuelos
Ranking.MaxDuelos.UserName = UserList(userindex).name
End If
End Sub
Public Sub CompruebaRetos(ByVal userindex As Integer)
' actualiza el ranking de duelos si el usuario tiene mas duelos que el mayor del ranking
If UserList(userindex).Stats.PuntosRetos > Ranking.MaxRetos.value And UserList(userindex).flags.Privilegios = PlayerType.User Then
Ranking.MaxRetos.value = UserList(userindex).Stats.PuntosRetos
Ranking.MaxRetos.UserName = UserList(userindex).name
End If
End Sub
Public Sub CompruebaPlantes(ByVal userindex As Integer)
' actualiza el ranking de plantes si el usuario tiene mas plantes que el mayor del ranking
If UserList(userindex).Stats.PuntosPlante > Ranking.MaxPlantes.value And UserList(userindex).flags.Privilegios = PlayerType.User Then
Ranking.MaxPlantes.value = UserList(userindex).Stats.PuntosPlante
Ranking.MaxPlantes.UserName = UserList(userindex).name
End If
End Sub
Public Sub CompruebaTorneos(ByVal userindex As Integer)
' actualiza el ranking de torneos si el usuario tiene mas torneos que el mayor del ranking
If UserList(userindex).Stats.PuntosTorneo > Ranking.MaxTorneos.value And UserList(userindex).flags.Privilegios = PlayerType.User Then
Ranking.MaxTorneos.value = UserList(userindex).Stats.PuntosTorneo
Ranking.MaxTorneos.UserName = UserList(userindex).name
End If
End Sub
Public Sub CompruebaDeaths(ByVal userindex As Integer)
' actualiza el ranking de deaths si el usuario tiene mas deaths que el mayor del ranking
If UserList(userindex).Stats.PuntosDeath > Ranking.MaxDeaths.value And UserList(userindex).flags.Privilegios = PlayerType.User Then
Ranking.MaxDeaths.value = UserList(userindex).Stats.PuntosDeath
Ranking.MaxDeaths.UserName = UserList(userindex).name
End If
End Sub
