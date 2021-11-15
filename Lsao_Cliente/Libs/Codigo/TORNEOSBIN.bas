Attribute VB_Name = "TORNEOSBIN"
Option Explicit
Private Torneo_Activo As Boolean
Private Torneo_Esperando As Boolean
Private Torneo_Rondas As Integer
Private Torneo_Combate As Integer
Private Torneo_Luchadores() As Integer

Private Const mapatorneo As Integer = 62
' esquinas superior isquierda del ring
Private Const esquina1x As Integer = 41
Private Const esquina1y As Integer = 46
' esquina inferior derecha del ring
Private Const esquina2x As Integer = 58
Private Const esquina2y As Integer = 57
' Donde esperan los hijos de puta
Private Const esperax As Integer = 32
Private Const esperay As Integer = 43
' Mapa desconecta
Private Const mapa_fuera As Integer = 1
Private Const fueraesperay As Integer = 50
Private Const fueraesperax As Integer = 50

Sub Rondas_UsuarioMuere(ByVal userindex As Integer, Optional Real As Boolean = True, Optional CambioMapa As Boolean = False)
        Dim i As Integer, pos As Integer, j As Integer
        Dim combate As Integer, LI1 As Integer, LI2 As Integer
        Dim UI1 As Integer, UI2 As Integer
If (Not (Torneo_Activo And (Not Torneo_Esperando))) Then
                Exit Sub
            ElseIf (Torneo_Activo And Torneo_Esperando) Then
                For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                    If (Torneo_Luchadores(i) = userindex) Then
                        Torneo_Luchadores(i) = -1
                        Exit Sub
                    End If
                Next i
            End If

        For pos = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                If (Torneo_Luchadores(pos) = userindex) Then Exit For
        Next pos

        ' si no lo ha encontrado
        If (Torneo_Luchadores(pos) <> userindex) Then Exit Sub

        combate = 1 + (pos - 1) / 2

        'ponemos li1 y li2 (luchador index) de los que combatian
        LI1 = 2 * (combate - 1)
        LI2 = LI1 + 1

        'se informa a la gente
        If (Real) Then
                Call SendData(SendTarget.toall, 0, 0, "||Torneo: " & UserList(userindex).name & " pierde el combate!" & FONTTYPE_TALK)
        Else
                Call SendData(SendTarget.toall, 0, 0, "||Torneo: " & UserList(userindex).name & " se fue del combate!" & FONTTYPE_TALK)
        End If

        'se le teleporta fuera si murio
        If (Real) Then
                Call WarpUserChar(Torneo_Luchadores(i), mapa_fuera, fueraesperax, fueraesperay, True)
        ElseIf (Not CambioMapa) Then
                'haz la mierda para ke se le guarde otro sitio en la ficha
                 Call WarpUserChar(Torneo_Luchadores(i), mapa_fuera, fueraesperax, fueraesperay, True)
        End If

        'se le borra de la lista y se mueve el segundo a li1
        If (Torneo_Luchadores(LI1) = userindex) Then
                Torneo_Luchadores(LI1) = Torneo_Luchadores(LI2) 'cambiamos slot
                Torneo_Luchadores(LI2) = -1
        Else
                Torneo_Luchadores(LI2) = -1
        End If

    'si es la ultima ronda
    If (Torneo_Rondas = 1) Then
        Call WarpUserChar(Torneo_Luchadores(LI1), mapa_fuera, fueraesperax, fueraesperay, True)
        Call SendData(SendTarget.toall, 0, 0, "||GANADOR DEL TORNEO: " & UserList(Torneo_Luchadores(LI1)).name & FONTTYPE_GUILD)
        Torneo_Activo = False
        Exit Sub
    Else
        'a su compañero se le teleporta dentro, condicional por seguridad
        Call WarpUserChar(Torneo_Luchadores(LI1), mapatorneo, esperax, esperay, True)
    End If

                
        'si es el ultimo combate de la ronda
        If (2 ^ Torneo_Combate = 2 * combate) Then

                Call SendData(SendTarget.toall, 0, 0, "||Torneo: Siguiente ronda!" & FONTTYPE_GUILD)
                Torneo_Rondas = Torneo_Rondas - 1

        'antes de llamar a la proxima ronda hay q copiar a los putos xD
        For i = 1 To 2 ^ Torneo_Rondas
                UI1 = Torneo_Luchadores(2 * (i - 1) + 1)
                UI2 = Torneo_Luchadores(2 * i)
                If (UI1 = -1) Then UI1 = UI2
                Torneo_Luchadores(i) = UI1
        Next i

        Call Rondas_Combate(1)
        Exit Sub
        End If

        'vamos al siguiente combate
        Call Rondas_Combate(combate + 1)
End Sub



Sub Rondas_UsuarioDesconecta(ByVal userindex As Integer)
        Call Rondas_UsuarioMuere(userindex, False, False)
End Sub



Sub Rondas_UsuarioCambiamapa(ByVal userindex As Integer)
        Call Rondas_UsuarioMuere(userindex, False, True)
End Sub



Sub Torneos_Inicia(ByVal userindex As Integer, ByVal Rondas As Integer)
        If (Torneo_Activo) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Ya hay un torneo!." & FONTTYPE_INFO)
                Exit Sub
        End If
        Call SendData(SendTarget.toall, 0, 0, "||Torneo: Esta empezando un nuevo torneo 1v1 de " & val(2 ^ Rondas) & " participantes!! para participar pon /PARTICIPAR" & FONTTYPE_GUILD)
        
        Torneo_Rondas = Rondas
        Torneo_Activo = True
        Torneo_Esperando = True

        ReDim Torneo_Luchadores(1 To 2 ^ Rondas) As Integer
        Dim i As Integer
        For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                Torneo_Luchadores(i) = -1
        Next i
End Sub



Sub Torneos_Entra(ByVal userindex As Integer)
        Dim i As Integer
        
        If (Not Torneo_Activo) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No hay ningun torneo!." & FONTTYPE_INFO)
                Exit Sub
        End If
        
        If (Not Torneo_Esperando) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||El torneo ya ha empezado, te quedaste fuera!." & FONTTYPE_INFO)
                Exit Sub
        End If
        
        For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                If (i = userindex) Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas dentro!" & FONTTYPE_WARNING)
                        Exit Sub
                End If
        Next i

        For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
        If (Torneo_Luchadores(i) = -1) Then
                Torneo_Luchadores(i) = userindex
                Call SendData(SendTarget.toindex, userindex, 0, "||Estas dentro del torneo!" & FONTTYPE_INFO)
                Call WarpUserChar(Torneo_Luchadores(i), mapatorneo, esperax, esperay, True)
                Call SendData(SendTarget.toall, 0, 0, "||Torneo: Entra el participante " & UserList(userindex).name & FONTTYPE_INFO)
                If (i = UBound(Torneo_Luchadores)) Then
                Call SendData(SendTarget.toall, 0, 0, "||Torneo: Empieza el torneo!" & FONTTYPE_GUILD)
                Torneo_Esperando = False
                Call Rondas_Combate(1)
        
                End If
        End If
        Next i
End Sub


Sub Rondas_Combate(combate As Integer)
Dim UI1 As Integer, UI2 As Integer
    UI1 = Torneo_Luchadores(2 * (combate - 1) + 1)
    UI2 = Torneo_Luchadores(2 * combate)
    
    If (UI2 = -1) Then
        UI2 = Torneo_Luchadores(2 * (combate - 1) + 1)
        UI1 = Torneo_Luchadores(2 * combate)
    End If
    
    If (UI1 = -1) Then
        Call SendData(SendTarget.toall, 0, 0, "||Torneo: Combate anulado porque un participante involucrado se desconecto" & FONTTYPE_TALK)
        If (Torneo_Rondas = 1) Then
            If (UI2 <> -1) Then
                Call SendData(SendTarget.toall, 0, 0, "||Torneo: Torneo terminado. Ganador del torneo por eliminacion: " & UserList(UI2).name & FONTTYPE_GUILD)
                ' dale_recompensa()
                Torneo_Activo = False
                Exit Sub
            End If
            Call SendData(SendTarget.toall, 0, 0, "||Torneo: Torneo terminado. No hay ganador porque todos se fueron :(" & FONTTYPE_GUILD)
            Exit Sub
        End If
        If (UI2 <> -1) Then _
            Call SendData(SendTarget.toall, 0, 0, "||Torneo: " & UserList(UI2).name & " pasa a la siguiente ronda!" & FONTTYPE_TALK)
    
        If (2 ^ Torneo_Rondas = 2 * combate) Then
            Call SendData(SendTarget.toall, 0, 0, "||Torneo: Siguiente ronda!" & FONTTYPE_GUILD)
            Torneo_Rondas = Torneo_Rondas - 1
            'antes de llamar a la proxima ronda hay q copiar a los putos xD
            Dim i As Integer, j As Integer
            For i = 1 To 2 ^ Torneo_Rondas
                UI1 = Torneo_Luchadores(2 * (i - 1) + 1)
                UI2 = Torneo_Luchadores(2 * i)
                If (UI1 = -1) Then UI1 = UI2
                Torneo_Luchadores(i) = UI1
            Next i
            Call Rondas_Combate(1)
            Exit Sub
        End If
        Call Rondas_Combate(combate + 1)
        Exit Sub
    End If

    Call SendData(SendTarget.toall, 0, 0, "||Torneo: " & UserList(UI1).name & " versus " & UserList(UI2).name & " A las esquinas!! Peleen!" & FONTTYPE_GUILD)

    Call WarpUserChar(UI1, mapatorneo, esquina1x, esquina1y, True)
    Call WarpUserChar(UI2, mapatorneo, esquina2x, esquina2y, True)
End Sub


