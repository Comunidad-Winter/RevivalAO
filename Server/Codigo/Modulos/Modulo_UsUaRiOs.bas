Attribute VB_Name = "UsUaRiOs"

Option Explicit

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Sub ActStats(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)

Dim DaExp As Integer

DaExp = CInt(UserList(VictimIndex).Stats.ELV * 2)

UserList(AttackerIndex).Stats.Exp = UserList(AttackerIndex).Stats.Exp + DaExp
If UserList(AttackerIndex).Stats.Exp > MAXEXP Then _
    UserList(AttackerIndex).Stats.Exp = MAXEXP

'Lo mata
Call SendData(SendTarget.toIndex, AttackerIndex, 0, "||Has matado a " & UserList(VictimIndex).name & "!" & FONTTYPE_FIGHT)
Call SendData(SendTarget.toIndex, AttackerIndex, 0, "||Has ganado " & DaExp & " puntos de experiencia." & FONTTYPE_FIGHT)
      
Call SendData(SendTarget.toIndex, VictimIndex, 0, "||" & UserList(AttackerIndex).name & " te ha matado!" & FONTTYPE_FIGHT)

If TriggerZonaPelea(VictimIndex, AttackerIndex) <> TRIGGER6_PERMITE Then
    If (Not Criminal(VictimIndex)) Then
         UserList(AttackerIndex).Reputacion.AsesinoRep = UserList(AttackerIndex).Reputacion.AsesinoRep + vlASESINO * 2
         If UserList(AttackerIndex).Reputacion.AsesinoRep > MAXREP Then _
            UserList(AttackerIndex).Reputacion.AsesinoRep = MAXREP
         UserList(AttackerIndex).Reputacion.BurguesRep = 0
         UserList(AttackerIndex).Reputacion.NobleRep = 0
         UserList(AttackerIndex).Reputacion.PlebeRep = 0
    Else
         UserList(AttackerIndex).Reputacion.NobleRep = UserList(AttackerIndex).Reputacion.NobleRep + vlNoble
         If UserList(AttackerIndex).Reputacion.NobleRep > MAXREP Then _
            UserList(AttackerIndex).Reputacion.NobleRep = MAXREP
    End If
End If

Call UserDie(VictimIndex)

If UserList(AttackerIndex).Stats.UsuariosMatados < 32000 Then _
    UserList(AttackerIndex).Stats.UsuariosMatados = UserList(AttackerIndex).Stats.UsuariosMatados + 1
Call CompruebaUserDies(AttackerIndex)
'Log
Call LogAsesinato(UserList(AttackerIndex).name & " asesino a " & UserList(VictimIndex).name)


End Sub


Sub RevivirUsuario(ByVal userindex As Integer)

UserList(userindex).flags.Muerto = 0
UserList(userindex).Stats.MinHP = 35

UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MaxSta

'No puede estar empollando
UserList(userindex).flags.EstaEmpo = 0
UserList(userindex).EmpoCont = 0

If UserList(userindex).Stats.MinHP > UserList(userindex).Stats.MaxHP Then
    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
End If

Call DarCuerpoDesnudo(userindex)
'[MaTeO 9]
Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).OrigChar.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
'[/MaTeO 9]
Call EnviarHP(userindex)
Call EnviarSta(userindex)
If UserList(userindex).flags.bandas = True Then
Call Transforma(userindex)
End If
End Sub


'[MaTeO 9]
Sub ChangeUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal userindex As Integer, _
                    ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, _
                    ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer, ByVal Alas As Integer)
'[/MaTeO 9]

    '[MaTeO ASEDIO]
    If UserList(userindex).Asedio.Participando Then
        If UserList(userindex).Raza = "Humano" Or _
           UserList(userindex).Raza = "Elfo" Or _
           UserList(userindex).Raza = "Elfo Oscuro" Then
                If UCase$(UserList(userindex).Clase) = "MAGO" Then
                    Select Case UserList(userindex).Asedio.Team
                        Case Equipos.Azul
                             Body = 516
                        Case Equipos.Negro
                             Body = 508
                        Case Equipos.Rojo
                             Body = 520
                        Case Equipos.Verde
                             Body = 512
                    End Select
                Else
                    Select Case UserList(userindex).Asedio.Team
                        Case Equipos.Azul
                             Body = 514
                        Case Equipos.Negro
                             Body = 506
                        Case Equipos.Rojo
                             Body = 518
                        Case Equipos.Verde
                             Body = 510
                    End Select
                End If
        Else
                If UCase$(UserList(userindex).Clase) = "MAGO" Then
                    Select Case UserList(userindex).Asedio.Team
                        Case Equipos.Azul
                             Body = 515
                        Case Equipos.Negro
                             Body = 507
                        Case Equipos.Rojo
                             Body = 519
                        Case Equipos.Verde
                             Body = 511
                    End Select
                Else
                    Select Case UserList(userindex).Asedio.Team
                        Case Equipos.Azul
                             Body = 517
                        Case Equipos.Negro
                             Body = 509
                        Case Equipos.Rojo
                             Body = 521
                        Case Equipos.Verde
                             Body = 513
                    End Select
                End If
        End If
    End If
    '[/MaTeO ASEDIO]

    UserList(userindex).char.Body = Body
    UserList(userindex).char.Head = Head
    UserList(userindex).char.Heading = Heading
    UserList(userindex).char.WeaponAnim = Arma
    UserList(userindex).char.ShieldAnim = Escudo
    UserList(userindex).char.CascoAnim = Casco
    Debug.Print "ChangeUserChar: " & UserList(userindex).Asedio.Participando & Rnd

    '[MaTeO 9]
    UserList(userindex).char.Alas = Alas
    '[/MaTeO 9]
       
    If sndRoute = SendTarget.ToMap Then
        '[MaTeO 9]
        Call SendToUserArea(userindex, "CP" & UserList(userindex).char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(userindex).char.FX & "," & UserList(userindex).char.loops & "," & Casco & "," & Alas)
        '[/MaTeO 9]
    Else
        '[MaTeO 9]
        Call SendData(sndRoute, sndIndex, sndMap, "CP" & UserList(userindex).char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(userindex).char.FX & "," & UserList(userindex).char.loops & "," & Casco & "," & Alas)
        '[/MaTeO 9]
    End If
End Sub

Sub EnviarSubirNivel(ByVal userindex As Integer, ByVal Puntos As Integer)
    Call SendData(SendTarget.toIndex, userindex, 0, "SUNI" & Puntos)
End Sub

Sub EnviarSkills(ByVal userindex As Integer)
    Dim i As Integer
    Dim cad As String
    
    For i = 1 To NUMSKILLS
       cad = cad & UserList(userindex).Stats.UserSkills(i) & ","
    Next i
    
    SendData SendTarget.toIndex, userindex, 0, "SKILLS" & cad$
End Sub

Sub EnviarFama(ByVal userindex As Integer)
    Dim cad As String
    
    cad = cad & UserList(userindex).Reputacion.AsesinoRep & ","
    cad = cad & UserList(userindex).Reputacion.BandidoRep & ","
    cad = cad & UserList(userindex).Reputacion.BurguesRep & ","
    cad = cad & UserList(userindex).Reputacion.LadronesRep & ","
    cad = cad & UserList(userindex).Reputacion.NobleRep & ","
    cad = cad & UserList(userindex).Reputacion.PlebeRep & ","
    
    Dim L As Long
    
    L = (-UserList(userindex).Reputacion.AsesinoRep) + _
        (-UserList(userindex).Reputacion.BandidoRep) + _
        UserList(userindex).Reputacion.BurguesRep + _
        (-UserList(userindex).Reputacion.LadronesRep) + _
        UserList(userindex).Reputacion.NobleRep + _
        UserList(userindex).Reputacion.PlebeRep
    L = L / 6
    
    UserList(userindex).Reputacion.Promedio = L
    
    cad = cad & UserList(userindex).Reputacion.Promedio
    
    SendData SendTarget.toIndex, userindex, 0, "FAMA" & cad
End Sub

Sub EnviarAtrib(ByVal userindex As Integer)
Dim i As Integer
Dim cad As String
For i = 1 To NUMATRIBUTOS
  cad = cad & UserList(userindex).Stats.UserAtributos(i) & ","
Next
Call SendData(SendTarget.toIndex, userindex, 0, "ATR" & cad)
End Sub

Public Sub EnviarMiniEstadisticas(ByVal userindex As Integer)
With UserList(userindex)
    Call SendData(SendTarget.toIndex, userindex, 0, "MEST" & .Faccion.CiudadanosMatados & "," & _
                .Faccion.CriminalesMatados & "," & .Stats.UsuariosMatados & "," & _
                .Stats.NPCsMuertos & "," & .Clase & "," & .Counters.Pena)
End With

End Sub

Sub EraseUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, userindex As Integer)

On Error GoTo ErrorHandler
   
    CharList(UserList(userindex).char.CharIndex) = 0
    
    If UserList(userindex).char.CharIndex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar <= 1 Then Exit Do
        Loop
    End If
    
Dim code As String
    code = str(UserList(userindex).char.CharIndex)
    'Le mandamos el mensaje para que borre el personaje a los clientes que estén en el mismo mapa
    If sndRoute = SendTarget.ToMap Then
        Call SendToUserArea(userindex, "BP" & code)
        Call QuitarUser(userindex, UserList(userindex).pos.Map)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "BP" & code)
    End If
    
    MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex = 0
    UserList(userindex).char.CharIndex = 0
    
    NumChars = NumChars - 1
    
    Exit Sub
    
ErrorHandler:
        Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.Description)

End Sub

Sub MakeUserChar(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal userindex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer)
On Local Error GoTo hayerror
    Dim CharIndex As Integer

    If InMapBounds(Map, x, Y) Then
        'If needed make a new character in list
        If UserList(userindex).char.CharIndex = 0 Then
            CharIndex = NextOpenCharIndex
            UserList(userindex).char.CharIndex = CharIndex
            CharList(CharIndex) = userindex
        End If
        
        'Place character on map
        MapData(Map, x, Y).userindex = userindex
        
        'Send make character command to clients
        Dim klan As String
        Dim Title As String
        If UserList(userindex).GuildIndex > 0 Then
            klan = Guilds(UserList(userindex).GuildIndex).GuildName
        End If
        Title = IIf(Len(UserList(userindex).Titulo), "[" & UserList(userindex).Titulo & "]", "")
        
        Dim bCr As Byte
        Dim SendPrivilegios As Byte
       
        bCr = Criminal(userindex)

        If klan <> "" Then
            If sndRoute = SendTarget.toIndex Then
#If SeguridadAlkon Then
                If EncriptarProtocolosCriticos Then
                    If UserList(userindex).flags.Privilegios > PlayerType.User Then
                        If UserList(userindex).showName Then
                            '[MaTeO 9]
                            Call SendCryptedData(SendTarget.toIndex, sndIndex, sndMap, "CC" & UserList(userindex).char.Body & "," & UserList(userindex).char.Head & "," & UserList(userindex).char.Heading & "," & UserList(userindex).char.CharIndex & "," & x & "," & Y & "," & UserList(userindex).char.WeaponAnim & "," & UserList(userindex).char.ShieldAnim & "," & UserList(userindex).char.FX & "," & 999 & "," & UserList(userindex).char.CascoAnim & "," & UserList(userindex).name & " <" & klan & ">" & Title & "," & bCr & "," & IIf(UserList(userindex).flags.EsRolesMaster, 5, UserList(userindex).flags.Privilegios) & "," & UserList(userindex).char.Alas)
                            '[/MaTeO 9]
                        Else
                            'Hide the name and clan
                            '[MaTeO 9]
                            Call SendCryptedData(SendTarget.toIndex, sndIndex, sndMap, "CC" & UserList(userindex).char.Body & "," & UserList(userindex).char.Head & "," & UserList(userindex).char.Heading & "," & UserList(userindex).char.CharIndex & "," & x & "," & Y & "," & UserList(userindex).char.WeaponAnim & "," & UserList(userindex).char.ShieldAnim & "," & UserList(userindex).char.FX & "," & 999 & "," & UserList(userindex).char.CascoAnim & ",," & bCr & "," & IIf(UserList(userindex).flags.EsRolesMaster, 5, UserList(userindex).flags.Privilegios) & "," & UserList(userindex).char.Alas)
                            '[/MaTeO 9]
                        End If
                    Else
                        '[MaTeO 9]
                        Call SendCryptedData(SendTarget.toIndex, sndIndex, sndMap, "CC" & UserList(userindex).char.Body & "," & UserList(userindex).char.Head & "," & UserList(userindex).char.Heading & "," & UserList(userindex).char.CharIndex & "," & x & "," & Y & "," & UserList(userindex).char.WeaponAnim & "," & UserList(userindex).char.ShieldAnim & "," & UserList(userindex).char.FX & "," & 999 & "," & UserList(userindex).char.CascoAnim & "," & UserList(userindex).name & " <" & klan & ">" & Title & "," & bCr & "," & IIf(UserList(userindex).flags.PertAlCons = 1, 4, IIf(UserList(userindex).flags.PertAlConsCaos = 1, 6, 0)) & "," & UserList(userindex).char.Alas)
                        '[/MaTeO 9]
                    End If
                Else
#End If
Dim code As String

                    If UserList(userindex).flags.Privilegios > PlayerType.User Then
                        If UserList(userindex).showName Then
                            '[MaTeO 9]
                            code = UserList(userindex).char.Body & "," & UserList(userindex).char.Head & "," & UserList(userindex).char.Heading & "," & UserList(userindex).char.CharIndex & "," & x & "," & Y & "," & UserList(userindex).char.WeaponAnim & "," & UserList(userindex).char.ShieldAnim & "," & UserList(userindex).char.FX & "," & 999 & "," & UserList(userindex).char.CascoAnim & "," & UserList(userindex).name & " <" & klan & ">" & Title & "," & bCr & "," & IIf(UserList(userindex).flags.EsRolesMaster, 5, UserList(userindex).flags.Privilegios) & "," & UserList(userindex).char.Alas
                            '[/MaTeO 9]
                            Call SendData(sndRoute, sndIndex, sndMap, "CC" & code) 'mandamos el CC encriptado
                        Else
                            'Hide the name and clan
                            '[MaTeO 9]
                            code = UserList(userindex).char.Body & "," & UserList(userindex).char.Head & "," & UserList(userindex).char.Heading & "," & UserList(userindex).char.CharIndex & "," & x & "," & Y & "," & UserList(userindex).char.WeaponAnim & "," & UserList(userindex).char.ShieldAnim & "," & UserList(userindex).char.FX & "," & 999 & "," & UserList(userindex).char.CascoAnim & ",," & bCr & "," & IIf(UserList(userindex).flags.EsRolesMaster, 5, UserList(userindex).flags.Privilegios) & "," & UserList(userindex).char.Alas
                            '[/MaTeO 9]
                            Call SendData(sndRoute, sndIndex, sndMap, "CC" & code)
                        End If
                    Else
                        '[MaTeO 9]
                        code = UserList(userindex).char.Body & "," & UserList(userindex).char.Head & "," & UserList(userindex).char.Heading & "," & UserList(userindex).char.CharIndex & "," & x & "," & Y & "," & UserList(userindex).char.WeaponAnim & "," & UserList(userindex).char.ShieldAnim & "," & UserList(userindex).char.FX & "," & 999 & "," & UserList(userindex).char.CascoAnim & "," & UserList(userindex).name & " <" & klan & ">" & Title & "," & bCr & "," & IIf(UserList(userindex).flags.PertAlCons = 1, 4, IIf(UserList(userindex).flags.PertAlConsCaos = 1, 6, 0)) & "," & UserList(userindex).char.Alas
                        '[/MaTeO 9]
                        Call SendData(sndRoute, sndIndex, sndMap, "CC" & code)
                    End If
#If SeguridadAlkon Then
                End If
#End If
            ElseIf sndRoute = SendTarget.ToMap Then
                Call AgregarUser(userindex, UserList(userindex).pos.Map)
                Call CheckUpdateNeededUser(userindex, USER_NUEVO)
            End If
        Else 'if tiene clan
            If sndRoute = SendTarget.toIndex Then
#If SeguridadAlkon Then
                If EncriptarProtocolosCriticos Then
                    If UserList(userindex).flags.Privilegios > PlayerType.User Then
                        If UserList(userindex).showName Then
                            Call SendCryptedData(SendTarget.toIndex, sndIndex, sndMap, "BC" & UserList(userindex).char.Body & "," & UserList(userindex).char.Head & "," & UserList(userindex).char.Heading & "," & UserList(userindex).char.CharIndex & "," & x & "," & Y & "," & UserList(userindex).char.WeaponAnim & "," & UserList(userindex).char.ShieldAnim & "," & UserList(userindex).char.FX & "," & 999 & "," & UserList(userindex).char.CascoAnim & "," & UserList(userindex).name & Title & "," & bCr & "," & IIf(UserList(userindex).flags.EsRolesMaster, 5, UserList(userindex).flags.Privilegios) & "," & UserList(userindex).char.Alas)
                        Else
                            'Hide the name
                            Call SendCryptedData(SendTarget.toIndex, sndIndex, sndMap, "BC" & UserList(userindex).char.Body & "," & UserList(userindex).char.Head & "," & UserList(userindex).char.Heading & "," & UserList(userindex).char.CharIndex & "," & x & "," & Y & "," & UserList(userindex).char.WeaponAnim & "," & UserList(userindex).char.ShieldAnim & "," & UserList(userindex).char.FX & "," & 999 & "," & UserList(userindex).char.CascoAnim & ",," & bCr & "," & IIf(UserList(userindex).flags.EsRolesMaster, 5, UserList(userindex).flags.Privilegios) & "," & UserList(userindex).char.Alas)
                        End If
                    Else
                        Call SendCryptedData(SendTarget.toIndex, sndIndex, sndMap, "BC" & UserList(userindex).char.Body & "," & UserList(userindex).char.Head & "," & UserList(userindex).char.Heading & "," & UserList(userindex).char.CharIndex & "," & x & "," & Y & "," & UserList(userindex).char.WeaponAnim & "," & UserList(userindex).char.ShieldAnim & "," & UserList(userindex).char.FX & "," & 999 & "," & UserList(userindex).char.CascoAnim & "," & UserList(userindex).name & Title & "," & bCr & "," & IIf(UserList(userindex).flags.PertAlCons = 1, 4, IIf(UserList(userindex).flags.PertAlConsCaos = 1, 6, 0)) & "," & UserList(userindex).char.Alas)
                    End If
                Else
#End If
                    If UserList(userindex).flags.Privilegios > PlayerType.User Then
                        If UserList(userindex).showName Then
                            Call SendData(SendTarget.toIndex, sndIndex, sndMap, "BC" & UserList(userindex).char.Body & "," & UserList(userindex).char.Head & "," & UserList(userindex).char.Heading & "," & UserList(userindex).char.CharIndex & "," & x & "," & Y & "," & UserList(userindex).char.WeaponAnim & "," & UserList(userindex).char.ShieldAnim & "," & UserList(userindex).char.FX & "," & 999 & "," & UserList(userindex).char.CascoAnim & "," & UserList(userindex).name & Title & "," & bCr & "," & IIf(UserList(userindex).flags.EsRolesMaster, 5, UserList(userindex).flags.Privilegios) & "," & UserList(userindex).char.Alas)
                        Else
                            Call SendData(SendTarget.toIndex, sndIndex, sndMap, "BC" & UserList(userindex).char.Body & "," & UserList(userindex).char.Head & "," & UserList(userindex).char.Heading & "," & UserList(userindex).char.CharIndex & "," & x & "," & Y & "," & UserList(userindex).char.WeaponAnim & "," & UserList(userindex).char.ShieldAnim & "," & UserList(userindex).char.FX & "," & 999 & "," & UserList(userindex).char.CascoAnim & ",," & bCr & "," & IIf(UserList(userindex).flags.EsRolesMaster, 5, UserList(userindex).flags.Privilegios) & "," & UserList(userindex).char.Alas)
                        End If
                    Else
                        Call SendData(SendTarget.toIndex, sndIndex, sndMap, "BC" & UserList(userindex).char.Body & "," & UserList(userindex).char.Head & "," & UserList(userindex).char.Heading & "," & UserList(userindex).char.CharIndex & "," & x & "," & Y & "," & UserList(userindex).char.WeaponAnim & "," & UserList(userindex).char.ShieldAnim & "," & UserList(userindex).char.FX & "," & 999 & "," & UserList(userindex).char.CascoAnim & "," & UserList(userindex).name & Title & "," & bCr & "," & IIf(UserList(userindex).flags.PertAlCons = 1, 4, IIf(UserList(userindex).flags.PertAlConsCaos = 1, 6, 0)) & "," & UserList(userindex).char.Alas)
                    End If
#If SeguridadAlkon Then
                End If
#End If
            ElseIf sndRoute = SendTarget.ToMap Then
                Call AgregarUser(userindex, UserList(userindex).pos.Map)
                Call CheckUpdateNeededUser(userindex, USER_NUEVO)
            End If
       End If   'if clan
    End If
Exit Sub

hayerror:
    LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.Description)
    'Resume Next
    Call CloseSocket(userindex)
End Sub

Sub CheckUserLevel(ByVal userindex As Integer)

On Error GoTo errhandler

Dim Pts As Integer
Dim AumentoHIT As Integer
Dim AumentoMANA As Integer
Dim AumentoSTA As Integer
Dim WasNewbie As Boolean

'¿Alcanzo el maximo nivel?
If UserList(userindex).Stats.ELV >= STAT_MAXELV Then
    UserList(userindex).Stats.Exp = 0
    UserList(userindex).Stats.ELU = 0
    Exit Sub
End If

WasNewbie = EsNewbie(userindex)

'Si exp >= then Exp para subir de nivel entonce subimos el nivel
'If UserList(UserIndex).Stats.Exp >= UserList(UserIndex).Stats.ELU Then
Do While UserList(userindex).Stats.Exp >= UserList(userindex).Stats.ELU
    
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_NIVEL)
    Call SendData(SendTarget.toIndex, userindex, 0, "||¡Has subido de nivel!" & FONTTYPE_INFO)
    
    If UserList(userindex).Stats.ELV = 1 Then
        Pts = 10
    Else
        Pts = 6
    End If
    
    UserList(userindex).Stats.SkillPts = UserList(userindex).Stats.SkillPts + Pts
    
    Call SendData(SendTarget.toIndex, userindex, 0, "||Has ganado " & Pts & " skillpoints." & FONTTYPE_INFO)
     ' rodra , no avisa total no hay =)
    UserList(userindex).Stats.ELV = UserList(userindex).Stats.ELV + 1
    
    UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp - UserList(userindex).Stats.ELU
    
    If Not EsNewbie(userindex) And WasNewbie Then
        Call QuitarNewbieObj(userindex)
        If UCase$(MapInfo(UserList(userindex).pos.Map).Restringir) = "SI" Then
            Call WarpUserChar(userindex, 1, 62, 42, True)
            Call SendData(SendTarget.toIndex, userindex, 0, "||Debes abandonar el Dungeon Newbie." & FONTTYPE_WARNING)
        End If
    End If

    If UserList(userindex).Stats.ELV < 11 Then
        UserList(userindex).Stats.ELU = UserList(userindex).Stats.ELU * 1.5
    ElseIf UserList(userindex).Stats.ELV < 25 Then
        UserList(userindex).Stats.ELU = UserList(userindex).Stats.ELU * 1.3
    Else
        UserList(userindex).Stats.ELU = UserList(userindex).Stats.ELU * 1.2
    End If

    Dim AumentoHP As Integer
    Select Case UCase$(UserList(userindex).Clase)
        Case "GUERRERO"
            Select Case UserList(userindex).Stats.UserAtributos(Constitucion)
                Case 21
                    AumentoHP = RandomNumber(10, 13)
                Case 20
                    AumentoHP = RandomNumber(8, 13)
                Case 19
                    AumentoHP = RandomNumber(8, 12)
                Case 18
                    AumentoHP = RandomNumber(8, 12)
                Case Else
                    AumentoHP = RandomNumber(6, UserList(userindex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
            End Select
            
            AumentoHIT = IIf(UserList(userindex).Stats.ELV > 35, 2, 3)
            AumentoSTA = AumentoSTDef
        
        Case "CAZADOR"
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(8, 13)
                Case 20
                    AumentoHP = RandomNumber(7, 12)
                Case 19
                    AumentoHP = RandomNumber(8, 11)
                Case 18
                    AumentoHP = RandomNumber(6, 11)
                Case Else
                    AumentoHP = RandomNumber(6, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2)
            End Select

            AumentoHIT = IIf(UserList(userindex).Stats.ELV > 35, 2, 3)
            AumentoSTA = AumentoSTDef
        
        Case "PALADIN"
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(8, 13)
                Case 20
                    AumentoHP = RandomNumber(7, 12)
                Case 19
                    AumentoHP = RandomNumber(7, 11)
                Case 18
                    AumentoHP = RandomNumber(7, 11)
                Case Else
                    AumentoHP = RandomNumber(6, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2)
            End Select
            
            AumentoHIT = IIf(UserList(userindex).Stats.ELV > 35, 1, 3)
            AumentoMANA = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
            
        Case "MAGO"
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(6, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 9)
                Case 19
                    AumentoHP = RandomNumber(5, 8)
                Case 18
                    AumentoHP = RandomNumber(5, 8)
                Case Else
                    AumentoHP = RandomNumber(5, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
            End Select
            If AumentoHP < 1 Then AumentoHP = 4
            
            AumentoHIT = 1
            AumentoMANA = 3 * UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTMago
        
        Case "CLERIGO"
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(7, 11)
                Case 20
                    AumentoHP = RandomNumber(6, 11)
                Case 19
                    AumentoHP = RandomNumber(6, 10)
                Case 18
                    AumentoHP = RandomNumber(5, 10)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(userindex).Stats.UserAtributos(Constitucion) \ 2)
            End Select
            
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case "ASESINO"
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(7, 11)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19
                    AumentoHP = RandomNumber(5, 9)
                Case 18
                    AumentoHP = RandomNumber(6, 9)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(userindex).Stats.UserAtributos(Constitucion) \ 2)
            End Select
            
            AumentoHIT = IIf(UserList(userindex).Stats.ELV > 35, 1, 3)
            AumentoMANA = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case "BARDO"
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(7, 11)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19
                    AumentoHP = RandomNumber(6, 9)
                Case 18
                    AumentoHP = RandomNumber(6, 9)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(userindex).Stats.UserAtributos(Constitucion) \ 2)
            End Select
            
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case Else
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(6, 9)
                Case 20
                    AumentoHP = RandomNumber(5, 9)
                Case 19, 18
                    AumentoHP = RandomNumber(4, 8)
                Case Else
                    AumentoHP = RandomNumber(5, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
            End Select

            AumentoHIT = 2
            AumentoSTA = AumentoSTDef
    End Select
    
    'Actualizamos HitPoints
    UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP + AumentoHP
    If UserList(userindex).Stats.MaxHP > STAT_MAXHP Then _
        UserList(userindex).Stats.MaxHP = STAT_MAXHP
    'Actualizamos Stamina
    UserList(userindex).Stats.MaxSta = UserList(userindex).Stats.MaxSta + AumentoSTA
    If UserList(userindex).Stats.MaxSta > STAT_MAXSTA Then _
        UserList(userindex).Stats.MaxSta = STAT_MAXSTA
    'Actualizamos Mana
    UserList(userindex).Stats.MaxMAN = UserList(userindex).Stats.MaxMAN + AumentoMANA
    If UserList(userindex).Stats.ELV < 36 Then
        If UserList(userindex).Stats.MaxMAN > STAT_MAXMAN Then _
            UserList(userindex).Stats.MaxMAN = STAT_MAXMAN
    Else
        If UserList(userindex).Stats.MaxMAN > 9999 Then _
            UserList(userindex).Stats.MaxMAN = 9999
    End If
    
    'Actualizamos Golpe Máximo
    UserList(userindex).Stats.MaxHIT = UserList(userindex).Stats.MaxHIT + AumentoHIT
    If UserList(userindex).Stats.ELV < 36 Then
        If UserList(userindex).Stats.MaxHIT > STAT_MAXHIT_UNDER36 Then _
            UserList(userindex).Stats.MaxHIT = STAT_MAXHIT_UNDER36
    Else
        If UserList(userindex).Stats.MaxHIT > STAT_MAXHIT_OVER36 Then _
            UserList(userindex).Stats.MaxHIT = STAT_MAXHIT_OVER36
    End If
    
    'Actualizamos Golpe Mínimo
    UserList(userindex).Stats.MinHIT = UserList(userindex).Stats.MinHIT + AumentoHIT
    If UserList(userindex).Stats.ELV < 36 Then
        If UserList(userindex).Stats.MinHIT > STAT_MAXHIT_UNDER36 Then _
            UserList(userindex).Stats.MinHIT = STAT_MAXHIT_UNDER36
    Else
        If UserList(userindex).Stats.MinHIT > STAT_MAXHIT_OVER36 Then _
            UserList(userindex).Stats.MinHIT = STAT_MAXHIT_OVER36
    End If
    
    'Notificamos al user
    If AumentoHP > 0 Then SendData SendTarget.toIndex, userindex, 0, "||Has ganado " & AumentoHP & " puntos de vida." & FONTTYPE_INFO
    If AumentoSTA > 0 Then SendData SendTarget.toIndex, userindex, 0, "||Has ganado " & AumentoSTA & " puntos de vitalidad." & FONTTYPE_INFO
    If AumentoMANA > 0 Then SendData SendTarget.toIndex, userindex, 0, "||Has ganado " & AumentoMANA & " puntos de magia." & FONTTYPE_INFO
    If AumentoHIT > 0 Then
        SendData SendTarget.toIndex, userindex, 0, "||Tu golpe maximo aumento en " & AumentoHIT & " puntos." & FONTTYPE_INFO
        SendData SendTarget.toIndex, userindex, 0, "||Tu golpe minimo aumento en " & AumentoHIT & " puntos." & FONTTYPE_INFO
    End If
    
    'Promedio CHOTS
    If UserList(userindex).Stats.ELV > 13 Then
    Dim Expromedio
    Dim Promedio
    Expromedio = Round((UserList(userindex).Stats.MaxHP - AumentoHP) / (UserList(userindex).Stats.ELV - 1), 2)
    Promedio = Round(UserList(userindex).Stats.MaxHP / UserList(userindex).Stats.ELV, 2)
    Call SendData(SendTarget.toIndex, userindex, 0, "||El Promedio de vida de tu Personaje era de " & Expromedio & FONTTYPE_ORO)
    Call SendData(SendTarget.toIndex, userindex, 0, "||Ahora el Promedio es de " & Promedio & FONTTYPE_ORO)
    End If
    'Promedio CHOTS
    
    Call LogDesarrollo(Date & " " & UserList(userindex).name & " paso a nivel " & UserList(userindex).Stats.ELV & " gano HP: " & AumentoHP)
    
    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
    
    Call EnviarSkills(userindex)
    Call EnviarSubirNivel(userindex, Pts)
    Call SendUserStatsBox(userindex)
    
Loop
'End If

If UserList(userindex).Stats.ELV = STAT_MAXELV Then
Exit Sub
Else
If UserList(userindex).Stats.ELU = 0 Then
Dim ind As String
    ind = UserList(userindex).char.CharIndex
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbCyan & "°" & "Has pasado al nivel " & UserList(userindex).Stats.ELV & "°" & ind)
End If
End If

Exit Sub

errhandler:
    LogError ("Error en la subrutina CheckUserLevel")
End Sub

Function PuedeAtravesarAgua(ByVal userindex As Integer) As Boolean

PuedeAtravesarAgua = _
  UserList(userindex).flags.Navegando = 1
End Function

Sub MoveUserChar(ByVal userindex As Integer, ByVal nHeading As Byte)

Dim nPos As WorldPos
    
    nPos = UserList(userindex).pos
    Call HeadtoPos(nHeading, nPos)
    
    If LegalPos(UserList(userindex).pos.Map, nPos.x, nPos.Y, PuedeAtravesarAgua(userindex)) Then
        If MapInfo(UserList(userindex).pos.Map).NumUsers > 1 Then
            'si no estoy solo en el mapa...
#If SeguridadAlkon Then
            Call SendCryptedMoveChar(nPos.Map, userindex, nPos.x, nPos.Y)
#Else
            Call SendToUserAreaButindex(userindex, "+" & UserList(userindex).char.CharIndex & "," & nPos.x & "," & nPos.Y)
#End If
        End If
        
        'Update map and user pos
        MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex = 0
        UserList(userindex).pos = nPos
        UserList(userindex).char.Heading = nHeading
        MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex = userindex
        If ZonaCura(userindex) Then Call AutoCuraUser(userindex)
        'Actualizamos las áreas de ser necesario
        Call ModAreas.CheckUpdateNeededUser(userindex, nHeading)
    Else
        Call SendData(SendTarget.toIndex, userindex, 0, "PU" & UserList(userindex).pos.x & "," & UserList(userindex).pos.Y)
    End If
    
   

    
End Sub
Sub AutoCuraUser(ByVal userindex As Integer)
If UserList(userindex).flags.Muerto = 1 Then
Call RevivirUsuario(userindex)
UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
Call SendData(toIndex, userindex, 0, "||El sacerdote te ha resucitado y curado." & FONTTYPE_INFO)
Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & 64 & "," & 2)
Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "TW121")
Call SendUserStatsBox(userindex)
End If

If UserList(userindex).Stats.MinHP < UserList(userindex).Stats.MaxHP Then
UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
Call SendData(toIndex, userindex, 0, "||El sacerdote te ha curado." & FONTTYPE_INFO)
Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & 64 & "," & 2)
Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "TW121")
Call SendUserStatsBox(userindex)
End If

If UserList(userindex).flags.Envenenado = 1 Then UserList(userindex).flags.Envenenado = 0


End Sub
Sub ChangeUserInv(userindex As Integer, Slot As Byte, Object As UserOBJ)



    If Object.ObjIndex > 0 Then
    ' cambiamos precio divido en 2 si es cheke de oro
    If ObjData(Object.ObjIndex).OBJType = eOBJType.otPLATA Then
    PrecioQl = 2
    Else
    PrecioQl = 3
    End If
        Call SendData(SendTarget.toIndex, userindex, 0, "CSI" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).name & "," & Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
        & ObjData(Object.ObjIndex).OBJType & "," _
        & ObjData(Object.ObjIndex).MaxHIT & "," _
        & ObjData(Object.ObjIndex).MinHIT & "," _
        & ObjData(Object.ObjIndex).MaxDef & "," _
        & ObjData(Object.ObjIndex).Valor \ PrecioQl)
    Else
        Call SendData(SendTarget.toIndex, userindex, 0, "CSI" & Slot & "," & "0" & "," & "(Vacío)" & "," & "0" & "," & "0")
    End If

End Sub


Function NextOpenCharIndex() As Integer
'Modificada por el oso para codificar los MP1234,2,1 en 2 bytes
'para lograrlo, el charindex no puede tener su bit numero 6 (desde 0) en 1
'y tampoco puede ser un charindex que tenga el bit 0 en 1.

On Local Error GoTo hayerror

Dim LoopC As Integer
    
    LoopC = 1
    
    While LoopC < MAXCHARS
        If CharList(LoopC) = 0 And Not ((LoopC And &HFFC0&) = 64) Then
            NextOpenCharIndex = LoopC
            NumChars = NumChars + 1
            If LoopC > LastChar Then LastChar = LoopC
            Exit Function
        Else
            LoopC = LoopC + 1
        End If
    Wend

Exit Function
hayerror:
LogError ("NextOpenCharIndex: num: " & Err.Number & " desc: " & Err.Description)

End Function

Function NextOpenUser() As Integer
        Dim LoopC As Long
       
        For LoopC = 1 To MaxUsers + 1
            If LoopC > MaxUsers Then Exit For
            If (UserList(LoopC).ConnID = -1 And UserList(LoopC).flags.UserLogged = False) Then Exit For
        Next LoopC
       
        NextOpenUser = LoopC
    End Function
Sub SendUserHitBox(ByVal userindex As Integer)
Dim lagaminarma As Integer
Dim lagamaxarma As Integer
Dim lagaminarmor As Integer
Dim lagamaxarmor As Integer
Dim lagaminescu As Integer
Dim lagamaxescu As Integer
Dim lagamincasc As Integer
Dim lagamaxcasc As Integer
If UserList(userindex).Invent.WeaponEqpSlot > 0 Then
lagaminarma = ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).MinHIT
lagamaxarma = ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).MaxHIT
Else
lagaminarma = "0"
lagamaxarma = "0"
End If
If UserList(userindex).Invent.ArmourEqpSlot > 0 Then
lagaminarmor = ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).MinDef
lagamaxarmor = ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).MaxDef
Else
lagaminarmor = "0"
lagamaxarmor = "0"
End If
If UserList(userindex).Invent.EscudoEqpSlot > 0 Then
lagaminescu = ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).MinDef
lagamaxescu = ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).MaxDef
Else
lagaminescu = "0"
lagamaxescu = "0"
End If
If UserList(userindex).Invent.CascoEqpSlot > 0 Then
lagamincasc = ObjData(UserList(userindex).Invent.CascoEqpObjIndex).MinDef
lagamaxcasc = ObjData(UserList(userindex).Invent.CascoEqpObjIndex).MaxDef
Else
lagamincasc = "0"
lagamaxcasc = "0"
End If
Call SendData(toIndex, userindex, 0, "ARM" & lagaminarma & "," & lagamaxarma & "," & lagaminarmor & "," & lagamaxarmor & "," & lagaminescu & "," & lagamaxescu & "," & lagamincasc & "," & lagamaxcasc)
End Sub
Sub EnviarDopa(ByVal userindex As Integer)
Dim Amarilla As Byte
Dim Verde As Byte
Verde = val(UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza))
Amarilla = val(UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad))
Call SendData(SendTarget.toIndex, userindex, 0, "DRG" & Amarilla & "," & Verde)
End Sub
Sub SendUserStatsBox(ByVal userindex As Integer)
Call CompruebaOro(userindex)

Call SendData(SendTarget.toIndex, userindex, 0, "EST" & UserList(userindex).Stats.MaxHP & "," & UserList(userindex).Stats.MinHP & "," & UserList(userindex).Stats.MaxMAN & "," & UserList(userindex).Stats.MinMAN & "," & UserList(userindex).Stats.MaxSta & "," & UserList(userindex).Stats.MinSta & "," & UserList(userindex).Stats.GLD & "," & UserList(userindex).Stats.ELV & "," & UserList(userindex).Stats.ELU & "," & UserList(userindex).Stats.Exp)

End Sub
Sub EnviarHP(ByVal userindex As Integer)
Call SendData(SendTarget.toIndex, userindex, 0, "VID" & UserList(userindex).Stats.MinHP)
End Sub
Sub EnviarMn(ByVal userindex As Integer)
Call SendData(SendTarget.toIndex, userindex, 0, "MN" & UserList(userindex).Stats.MinMAN)
End Sub
Sub EnviarSta(ByVal userindex As Integer)
Call SendData(SendTarget.toIndex, userindex, 0, "STA" & UserList(userindex).Stats.MinSta)
End Sub
Sub EnviarOro(ByVal userindex As Integer)
Call CompruebaOro(userindex)

Call SendData(SendTarget.toIndex, userindex, 0, "ORO" & UserList(userindex).Stats.GLD)

End Sub
Sub EnviarExp(ByVal userindex As Integer)
Call SendData(SendTarget.toIndex, userindex, 0, "EXP" & UserList(userindex).Stats.Exp)
End Sub
Sub EnviarHambreYsed(ByVal userindex As Integer)
Call SendData(SendTarget.toIndex, userindex, 0, "EHYS" & UserList(userindex).Stats.MinAGU & "," & UserList(userindex).Stats.MinHam)
End Sub

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
Dim GuildI As Integer


    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Estadisticas de: " & UserList(userindex).name & FONTTYPE_INFO)
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Nivel: " & UserList(userindex).Stats.ELV & "  EXP: " & UserList(userindex).Stats.Exp & "/" & UserList(userindex).Stats.ELU & FONTTYPE_INFO)
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Salud: " & UserList(userindex).Stats.MinHP & "/" & UserList(userindex).Stats.MaxHP & "  Mana: " & UserList(userindex).Stats.MinMAN & "/" & UserList(userindex).Stats.MaxMAN & "  Vitalidad: " & UserList(userindex).Stats.MinSta & "/" & UserList(userindex).Stats.MaxSta & FONTTYPE_INFO)
    
    If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
        Call SendData(SendTarget.toIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(userindex).Stats.MinHIT & "/" & UserList(userindex).Stats.MaxHIT & " (" & ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).MaxHIT & ")" & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.toIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(userindex).Stats.MinHIT & "/" & UserList(userindex).Stats.MaxHIT & FONTTYPE_INFO)
    End If
    
    If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
        Call SendData(SendTarget.toIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: " & ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).MaxDef & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.toIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: 0" & FONTTYPE_INFO)
    End If
    
    If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then
        Call SendData(SendTarget.toIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: " & ObjData(UserList(userindex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(userindex).Invent.CascoEqpObjIndex).MaxDef & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.toIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: 0" & FONTTYPE_INFO)
    End If
    
    GuildI = UserList(userindex).GuildIndex
    If GuildI > 0 Then
        Call SendData(SendTarget.toIndex, sendIndex, 0, "||Clan: " & Guilds(GuildI).GuildName & FONTTYPE_INFO)
        If UCase$(Guilds(GuildI).GetLeader) = UCase$(UserList(sendIndex).name) Then
            Call SendData(SendTarget.toIndex, sendIndex, 0, "||Status: Lider" & FONTTYPE_INFO)
        End If
        'guildpts no tienen objeto
        'Call SendData(SendTarget.ToIndex, sendIndex, 0, "||User GuildPoints: " & UserList(UserIndex).GuildInfo.GuildPoints & FONTTYPE_INFO)
    End If
    
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Oro: " & UserList(userindex).Stats.GLD & "  Posicion: " & UserList(userindex).pos.x & "," & UserList(userindex).pos.Y & " en mapa " & UserList(userindex).pos.Map & FONTTYPE_INFO)
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Dados: " & UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) & ", " & UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) & ", " & UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) & ", " & UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) & ", " & UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) & FONTTYPE_INFO)
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Trofeos de Oro: " & UserList(userindex).Stats.TrofOro & "~255~255~6~0~0~")
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Trofeos de Plata: " & UserList(userindex).Stats.TrofPlata & "~255~255~251~0~0~")
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Trofeos de Bronce: " & UserList(userindex).Stats.TrofBronce & "~187~0~0~0~0~")
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Amuletos de Madera: " & UserList(userindex).Stats.TrofMadera & "~237~207~139~0~0~")
    

End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
With UserList(userindex)
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Pj: " & .name & FONTTYPE_INFO)
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||CiudadanosMatados: " & .Faccion.CiudadanosMatados & " CriminalesMatados: " & .Faccion.CriminalesMatados & " UsuariosMatados: " & .Stats.UsuariosMatados & FONTTYPE_INFO)
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||NPCsMuertos: " & .Stats.NPCsMuertos & FONTTYPE_INFO)
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Clase: " & .Clase & FONTTYPE_INFO)
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Pena: " & .Counters.Pena & FONTTYPE_INFO)
End With

End Sub

Sub SendUserMiniStatsTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
Dim CharFile As String
Dim Ban As String
Dim BanDetailPath As String

    BanDetailPath = App.Path & "\logs\" & "BanDetail.dat"
    CharFile = CharPath & CharName & ".chr"
    
    If FileExist(CharFile) Then
        Call SendData(SendTarget.toIndex, sendIndex, 0, "||Pj: " & CharName & FONTTYPE_INFO)
        ' 3 en uno :p
        Call SendData(SendTarget.toIndex, sendIndex, 0, "||CiudadanosMatados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " UsuariosMatados: " & GetVar(CharFile, "MUERTES", "UserMuertes") & FONTTYPE_INFO)
        Call SendData(SendTarget.toIndex, sendIndex, 0, "||NPCsMuertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes") & FONTTYPE_INFO)
        Call SendData(SendTarget.toIndex, sendIndex, 0, "||Clase: " & GetVar(CharFile, "INIT", "Clase") & FONTTYPE_INFO)
        Call SendData(SendTarget.toIndex, sendIndex, 0, "||Pena: " & GetVar(CharFile, "COUNTERS", "PENA") & FONTTYPE_INFO)
        Ban = GetVar(CharFile, "FLAGS", "Ban")
        Call SendData(SendTarget.toIndex, sendIndex, 0, "||Ban: " & Ban & FONTTYPE_INFO)
        If Ban = "1" Then
            Call SendData(SendTarget.toIndex, sendIndex, 0, "||Baneado por: " & GetVar(CharFile, CharName, "BannedBy") & " El Motivo Fue: " & GetVar(BanDetailPath, CharName, "Reason") & FONTTYPE_INFO)
        End If
    Else
        Call SendData(SendTarget.toIndex, sendIndex, 0, "||El pj no existe: " & CharName & FONTTYPE_INFO)
    End If
    
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
On Error Resume Next

    Dim j As Long
    
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||" & UserList(userindex).name & FONTTYPE_INFO)
    Call SendData(SendTarget.toIndex, sendIndex, 0, "|| Tiene " & UserList(userindex).Invent.NroItems & " objetos." & FONTTYPE_INFO)
    
    For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(userindex).Invent.Object(j).ObjIndex > 0 Then
            Call SendData(SendTarget.toIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(UserList(userindex).Invent.Object(j).ObjIndex).name & " Cantidad:" & UserList(userindex).Invent.Object(j).Amount & FONTTYPE_INFO)
        End If
    Next j
End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
On Error Resume Next

    Dim j As Long
    Dim CharFile As String, Tmp As String
    Dim ObjInd As Long, ObjCant As Long
    
    CharFile = CharPath & CharName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call SendData(SendTarget.toIndex, sendIndex, 0, "||" & CharName & FONTTYPE_INFO)
        Call SendData(SendTarget.toIndex, sendIndex, 0, "|| Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos." & FONTTYPE_INFO)
        
        For j = 1 To MAX_INVENTORY_SLOTS
            Tmp = GetVar(CharFile, "Inventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))
            If ObjInd > 0 Then
                Call SendData(SendTarget.toIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(ObjInd).name & " Cantidad:" & ObjCant & FONTTYPE_INFO)
            End If
        Next j
    Else
        Call SendData(SendTarget.toIndex, sendIndex, 0, "||Usuario inexistente: " & CharName & FONTTYPE_INFO)
    End If
    
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
On Error Resume Next
Dim j As Integer
Call SendData(SendTarget.toIndex, sendIndex, 0, "||" & UserList(userindex).name & FONTTYPE_INFO)
For j = 1 To NUMSKILLS
    Call SendData(SendTarget.toIndex, sendIndex, 0, "|| " & SkillsNames(j) & " = " & UserList(userindex).Stats.UserSkills(j) & FONTTYPE_INFO)
Next
Call SendData(SendTarget.toIndex, sendIndex, 0, "|| SkillLibres:" & UserList(userindex).Stats.SkillPts & FONTTYPE_INFO)
End Sub

Function DameUserindex(SocketId As Integer) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
Do Until UserList(LoopC).ConnID = SocketId

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        DameUserindex = 0
        Exit Function
    End If
    
Loop
  
DameUserindex = LoopC

End Function

Function DameUserIndexConNombre(ByVal nombre As String) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
nombre = UCase$(nombre)

Do Until UCase$(UserList(LoopC).name) = nombre

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        DameUserIndexConNombre = 0
        Exit Function
    End If
    
Loop
  
DameUserIndexConNombre = LoopC

End Function


Function EsMascotaCiudadano(ByVal NpcIndex As Integer, ByVal userindex As Integer) As Boolean

If Npclist(NpcIndex).MaestroUser > 0 Then
        EsMascotaCiudadano = Not Criminal(Npclist(NpcIndex).MaestroUser)
        If EsMascotaCiudadano Then Call SendData(SendTarget.toIndex, Npclist(NpcIndex).MaestroUser, 0, "||¡¡" & UserList(userindex).name & " esta atacando tu mascota!!" & FONTTYPE_FIGHT)
End If

End Function

Sub NpcAtacado(ByVal NpcIndex As Integer, ByVal userindex As Integer)


'Guardamos el usuario que ataco el npc
Npclist(NpcIndex).flags.AttackedBy = UserList(userindex).name

If Npclist(NpcIndex).MaestroUser > 0 Then Call AllMascotasAtacanUser(userindex, Npclist(NpcIndex).MaestroUser)

If EsMascotaCiudadano(NpcIndex, userindex) Then
            Call VolverCriminal(userindex)
            Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
            Npclist(NpcIndex).Hostile = 1
Else
    'Reputacion
    If Npclist(NpcIndex).Stats.Alineacion = 0 Then
       If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            UserList(userindex).Reputacion.NobleRep = 0
            UserList(userindex).Reputacion.PlebeRep = 0
            UserList(userindex).Reputacion.AsesinoRep = UserList(userindex).Reputacion.AsesinoRep + 200
            If UserList(userindex).Reputacion.AsesinoRep > MAXREP Then _
                UserList(userindex).Reputacion.AsesinoRep = MAXREP
       Else
            If Not Npclist(NpcIndex).MaestroUser > 0 Then   'mascotas nooo!
                UserList(userindex).Reputacion.BandidoRep = UserList(userindex).Reputacion.BandidoRep + vlASALTO
                If UserList(userindex).Reputacion.BandidoRep > MAXREP Then _
                    UserList(userindex).Reputacion.BandidoRep = MAXREP
            End If
       End If
    ElseIf Npclist(NpcIndex).Stats.Alineacion = 1 Then
       UserList(userindex).Reputacion.PlebeRep = UserList(userindex).Reputacion.PlebeRep + vlCAZADOR / 2
       If UserList(userindex).Reputacion.PlebeRep > MAXREP Then _
        UserList(userindex).Reputacion.PlebeRep = MAXREP
    End If
    
    'hacemos que el npc se defienda
    Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
    Npclist(NpcIndex).Hostile = 1
    
End If

End Sub

Function PuedeApuñalar(ByVal userindex As Integer) As Boolean

If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
 PuedeApuñalar = _
 ((UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR) _
 And (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).Apuñala = 1)) _
 Or _
  ((UCase$(UserList(userindex).Clase) = "ASESINO") And _
  (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).Apuñala = 1))
Else
 PuedeApuñalar = False
End If
End Function
Sub SubirSkill(ByVal userindex As Integer, ByVal Skill As Integer)
Dim Aumenta As Integer
Aumenta = RandomNumber(1, 2)
If UserList(userindex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Or UserList(userindex).Stats.UserSkills(Skill) = 1000 Then Exit Sub
If Aumenta = 2 Then
UserList(userindex).Stats.UserSkills(Skill) = UserList(userindex).Stats.UserSkills(Skill) + 1
Call SendData(toIndex, userindex, 0, "||¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(userindex).Stats.UserSkills(Skill) & " pts." & FONTTYPE_INFO)
UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + 100
Call SendData(toIndex, userindex, 0, "Z25")
Call CheckUserLevel(userindex)
Call EnviarExp(userindex)
End If
End Sub

Sub UserDie(ByVal userindex As Integer)
On Error GoTo ErrorHandler

    'Sonido
    If UCase$(UserList(userindex).Genero) = "MUJER" Then
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, e_SoundIndex.MUERTE_MUJER)
    Else
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, e_SoundIndex.MUERTE_HOMBRE)
    End If

    'Quitar el dialogo del user muerto
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "QDL" & UserList(userindex).char.CharIndex)
    
    UserList(userindex).Stats.MinHP = 0
    UserList(userindex).Stats.MinSta = 0
    UserList(userindex).flags.AtacadoPorNpc = 0
    UserList(userindex).flags.AtacadoPorUser = 0
    UserList(userindex).flags.Envenenado = 0
    UserList(userindex).flags.Muerto = 1
    
    
    Dim aN As Integer
    
    aN = UserList(userindex).flags.AtacadoPorNpc
    
    If aN > 0 Then
        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
        Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
        Npclist(aN).flags.AttackedBy = ""
    End If
    
    '<<<< Paralisis >>>>
    If UserList(userindex).flags.Paralizado = 1 Then
        UserList(userindex).flags.Paralizado = 0
        Call SendData(SendTarget.toIndex, userindex, 0, "PARADOW")
        
    End If
    
    '<<< Estupidez >>>
    If UserList(userindex).flags.Estupidez = 1 Then
        UserList(userindex).flags.Estupidez = 0
        Call SendData(SendTarget.toIndex, userindex, 0, "NESTUP")
    End If
    
    '<<<< Descansando >>>>
    If UserList(userindex).flags.Descansar Then
        UserList(userindex).flags.Descansar = False
        Call SendData(SendTarget.toIndex, userindex, 0, "DOK")
    End If
    
    '<<<< Meditando >>>>
    If UserList(userindex).flags.Meditando Then
        UserList(userindex).flags.Meditando = False
        Call SendData(SendTarget.toIndex, userindex, 0, "MEDOK")
    End If
    
    '<<<< Invisible >>>>
    If UserList(userindex).flags.Invisible = 1 Or UserList(userindex).flags.Oculto = 1 Then
        UserList(userindex).flags.Oculto = 0
        UserList(userindex).flags.Invisible = 0
        'no hace falta encriptar este NOVER
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).char.CharIndex & ",0")
    End If
    
    If TriggerZonaPelea(userindex, userindex) <> TRIGGER6_PERMITE Then
    If Not UserList(userindex).pos.Map = 1 Then
        ' << Si es newbie no pierde el inventario >>
        If Not EsNewbie(userindex) Or Criminal(userindex) Then
            Call TirarTodo(userindex)
        Else
            If EsNewbie(userindex) Then Call TirarTodosLosItemsNoNewbies(userindex)
        End If
        End If
    End If
    
    ' DESEQUIPA TODOS LOS OBJETOS
    'desequipar armadura
    If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.ArmourEqpSlot)
    End If
    'desequipar arma
    If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.WeaponEqpSlot)
    End If
    If UserList(userindex).Invent.AlaEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.AlaEqpSlot)
    End If
    'desequipar casco
    If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.CascoEqpSlot)
    End If
    'desequipar herramienta
    If UserList(userindex).Invent.HerramientaEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.HerramientaEqpSlot)
    End If
    'desequipar municiones
    If UserList(userindex).Invent.MunicionEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.MunicionEqpSlot)
    End If
    'desequipar escudo
    If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.EscudoEqpSlot)
    End If
       
    ' << Reseteamos los posibles FX sobre el personaje >>
    If UserList(userindex).char.loops = LoopAdEternum Then
        UserList(userindex).char.FX = 0
        UserList(userindex).char.loops = 0
    End If
'////////////////////////////////////////////////////////////////////////////////////////////////////////////
 If UserList(userindex).flags.automatico = True Then
Call Rondas_UsuarioMuere(userindex)
End If
If UserList(userindex).flags.death = True Then
Call death_muere(userindex)
End If
If UserList(userindex).flags.bandas = True Then
Call Ban_Muere(userindex)
End If
If UserList(userindex).flags.Montado = True Then
UserList(userindex).flags.NumeroMont = 0
UserList(userindex).flags.Montado = False
End If



    ' <<Si pierde el duelo se va>>
If UserList(userindex).pos.Map = 61 And userindex = duelosespera Then
Call WarpUserChar(userindex, 1, 50, 50, True)
Call SendData(SendTarget.toIndex, userindex, 0, "||Pierdes el duelo." & FONTTYPE_WARNING)
Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Duelos: " & UserList(duelosespera).name & " ha Perdido el duelo." & FONTTYPE_TALK)
Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Duelos: " & UserList(duelosreta).name & " ha ganado el duelo y espera otro rival." & FONTTYPE_TALK)
duelosespera = duelosreta
numduelos = 0
End If
If UserList(userindex).pos.Map = 117 Then
Call WarpUserCharX(userindex, 1, 50, 50, True)
UserList(userindex).Counters.maparql = 0
End If
If UserList(userindex).pos.Map = 61 And userindex = duelosreta Then
Call WarpUserChar(userindex, 1, 50, 50, True)
numduelos = numduelos + 1
UserList(duelosespera).Stats.PuntosDuelos = UserList(duelosespera).Stats.PuntosDuelos + 1
Call CompruebaDuelos(duelosespera)
Call SendData(SendTarget.toIndex, userindex, 0, "||Pierdes el duelo." & FONTTYPE_WARNING)
Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Duelos: " & UserList(duelosreta).name & " ha Perdido el duelo." & FONTTYPE_TALK)
If (numduelos <= 50 And numduelos Mod 5 = 0) Or (numduelos Mod 10 = 0) Then
Call SendData(SendTarget.toAll, 0, 0, "TW123")
Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Duelos: " & UserList(duelosespera).name & " ha ganado " & numduelos & " consecutivos!" & FONTTYPE_TALK)

End If
Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Duelos: " & UserList(duelosespera).name & " ha ganado el duelo y espera otro rival." & FONTTYPE_TALK)
End If
    ' << Restauramos el mimetismo
    If UserList(userindex).flags.Mimetizado = 1 Then
        UserList(userindex).char.Body = UserList(userindex).CharMimetizado.Body
        UserList(userindex).char.Head = UserList(userindex).CharMimetizado.Head
        UserList(userindex).char.CascoAnim = UserList(userindex).CharMimetizado.CascoAnim
      
        UserList(userindex).char.ShieldAnim = UserList(userindex).CharMimetizado.ShieldAnim
        UserList(userindex).char.WeaponAnim = UserList(userindex).CharMimetizado.WeaponAnim
        UserList(userindex).Counters.Mimetismo = 0
        UserList(userindex).flags.Mimetizado = 0
    End If
    
    '<< Cambiamos la apariencia del char >>
    If UserList(userindex).flags.Navegando = 0 Then
        UserList(userindex).char.Body = iCuerpoMuerto
        UserList(userindex).char.Head = iCabezaMuerto
        UserList(userindex).char.ShieldAnim = NingunEscudo
        UserList(userindex).char.WeaponAnim = NingunArma
        UserList(userindex).char.CascoAnim = NingunCasco
       
    Else
        UserList(userindex).char.Body = iFragataFantasmal ';)
    End If
   
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        
        If UserList(userindex).MascotasIndex(i) > 0 Then
               If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                    Call MuereNpc(UserList(userindex).MascotasIndex(i), 0)
               Else
                    Npclist(UserList(userindex).MascotasIndex(i)).MaestroUser = 0
                    Npclist(UserList(userindex).MascotasIndex(i)).Movement = Npclist(UserList(userindex).MascotasIndex(i)).flags.OldMovement
                    Npclist(UserList(userindex).MascotasIndex(i)).Hostile = Npclist(UserList(userindex).MascotasIndex(i)).flags.OldHostil
                    UserList(userindex).MascotasIndex(i) = 0
                    UserList(userindex).MascotasType(i) = 0
               End If
        End If
        
    Next i
    
    UserList(userindex).NroMacotas = 0
    
    If Criminal(userindex) Then
    Call SendData(SendTarget.toIndex, userindex, 0, "Z33")
    Else
    Call SendData(SendTarget.toIndex, userindex, 0, "Z34")
    End If
    
    'If MapInfo(UserList(UserIndex).Pos.Map).Pk Then
    '        Dim MiObj As Obj
    '        Dim nPos As WorldPos
    '        MiObj.ObjIndex = RandomNumber(554, 555)
    '        MiObj.Amount = 1
    '        nPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    '        Dim ManchaSangre As New cGarbage
    '        ManchaSangre.Map = nPos.Map
    '        ManchaSangre.X = nPos.X
    '        ManchaSangre.Y = nPos.Y
    '        Call TrashCollector.Add(ManchaSangre)
    'End If
     
    '<< Actualizamos clientes >>
       '[MaTeO 9]
    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, val(userindex), UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, NingunArma, NingunEscudo, NingunCasco, NingunAlas)
    '[/MaTeO 9]
  
    Call SendUserStatsBox(userindex)
    Call SendUserHitBox(userindex)
    Call EnviarDopa(userindex)
    
     
    '<<Castigos por party>>
    If UserList(userindex).PartyIndex > 0 Then
        Call mdParty.ObtenerExito(userindex, UserList(userindex).Stats.ELV * -10 * mdParty.CantMiembros(userindex), UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y)
    End If
Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & 57 & "," & 1)
If UserList(userindex).flags.EstaDueleando1 = True Then
    Call TerminarDueloPlantes(UserList(userindex).flags.Oponente1, userindex)
    End If
    If UserList(userindex).flags.EnDosVDos = True Then
    Call VerificarDosVDos(userindex)
End If
If UserList(userindex).flags.EstaDueleando = True Then
    Call TerminarDuelo(UserList(userindex).flags.Oponente, userindex)
    End If
    
    '[MaTeO ASEDIO]
    If UserList(userindex).Asedio.Participando Then
        Call modAsedio.MuereUser(userindex)
    End If
    '[/MaTeO ASEDIO]
Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripción: " & Err.Description)
End Sub


Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)
On Error GoTo ErrorHandler
    If EsNewbie(Muerto) Then Exit Sub
    
    If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
    
    If Criminal(Muerto) Then
        If UserList(Atacante).flags.LastCrimMatado <> UserList(Muerto).name Then
            UserList(Atacante).flags.LastCrimMatado = UserList(Muerto).name
            If UserList(Atacante).Faccion.CriminalesMatados < 65000 Then _
                UserList(Atacante).Faccion.CriminalesMatados = UserList(Atacante).Faccion.CriminalesMatados + 1
        End If
        
        If UserList(Atacante).Faccion.CriminalesMatados > MAXUSERMATADOS Then
            UserList(Atacante).Faccion.CriminalesMatados = 0
            UserList(Atacante).Faccion.RecompensasReal = 0
        End If
        
        If UserList(Atacante).Faccion.RecibioExpInicialCaos = 1 And UserList(Muerto).Faccion.FuerzasCaos = 1 Then
            UserList(Atacante).Faccion.Reenlistadas = 200  'jaja que trucho
            
            'con esto evitamos que se vuelva a reenlistar
        End If
    Else
        If UserList(Atacante).flags.LastCiudMatado <> UserList(Muerto).name Then
            UserList(Atacante).flags.LastCiudMatado = UserList(Muerto).name
            If UserList(Atacante).Faccion.CiudadanosMatados < 65000 Then _
                UserList(Atacante).Faccion.CiudadanosMatados = UserList(Atacante).Faccion.CiudadanosMatados + 1
        End If
        
        If UserList(Atacante).Faccion.CiudadanosMatados > MAXUSERMATADOS Then
            UserList(Atacante).Faccion.CiudadanosMatados = 0
            UserList(Atacante).Faccion.RecompensasCaos = 0
        End If
    End If
ErrorHandler:
  '  Call LogError("Error en SUB CONTARMUERTE. Error: " & Err.Number & " Descripción: " & Err.Description)

End Sub

Sub Tilelibre(ByRef pos As WorldPos, ByRef nPos As WorldPos, ByRef Obj As Obj)
'Call LogTarea("Sub Tilelibre")

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer
Dim hayobj As Boolean
    hayobj = False
    nPos.Map = pos.Map
    
    Do While Not LegalPos(pos.Map, nPos.x, nPos.Y) Or hayobj
        
        If LoopC > 15 Then
            Notfound = True
            Exit Do
        End If
        
        For tY = pos.Y - LoopC To pos.Y + LoopC
            For tX = pos.x - LoopC To pos.x + LoopC
            
                If LegalPos(nPos.Map, tX, tY) Then
                    'We continue if: a - the item is different from 0 and the dropped item or b - the amount dropped + amount in map exceeds MAX_INVENTORY_OBJS
                    hayobj = (MapData(nPos.Map, tX, tY).OBJInfo.ObjIndex > 0 And MapData(nPos.Map, tX, tY).OBJInfo.ObjIndex <> Obj.ObjIndex)
                    If Not hayobj Then _
                        hayobj = (MapData(nPos.Map, tX, tY).OBJInfo.Amount + Obj.Amount > MAX_INVENTORY_OBJS)
                    If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                        nPos.x = tX
                        nPos.Y = tY
                        tX = pos.x + LoopC
                        tY = pos.Y + LoopC
                    End If
                End If
            
            Next tX
        Next tY
        
        LoopC = LoopC + 1
        
    Loop
    
    If Notfound = True Then
        nPos.x = 0
        nPos.Y = 0
    End If

End Sub

Sub WarpUserChar(ByVal userindex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)

Dim OldMap As Integer
Dim OldX As Integer
Dim OldY As Integer

    'Quitar el dialogo
    Call SendToUserArea(userindex, "QDL" & UserList(userindex).char.CharIndex)
    Call SendData(SendTarget.toIndex, userindex, UserList(userindex).pos.Map, "QTDL")
    
    OldMap = UserList(userindex).pos.Map
    OldX = UserList(userindex).pos.x
    OldY = UserList(userindex).pos.Y
    
    Call EraseUserChar(SendTarget.ToMap, 0, OldMap, userindex)
        
    If OldMap <> Map Then
        Call SendData(SendTarget.toIndex, userindex, 0, "CM" & Map & "," & MapInfo(UserList(userindex).pos.Map).MapVersion)
        Call SendData(SendTarget.toIndex, userindex, 0, "TM" & MapInfo(Map).Music)
        Call SendData(SendTarget.toIndex, userindex, 0, "N~" & MapInfo(Map).name)
        
        'Update new Map Users
        MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
    
        'Update old Map Users
        MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
        If MapInfo(OldMap).NumUsers < 0 Then
            MapInfo(OldMap).NumUsers = 0
        End If
    End If
    
    UserList(userindex).pos.x = x
    UserList(userindex).pos.Y = Y
    UserList(userindex).pos.Map = Map
    
    Call MakeUserChar(SendTarget.ToMap, 0, Map, userindex, Map, x, Y)
    Call SendData(SendTarget.toIndex, userindex, 0, "IP" & UserList(userindex).char.CharIndex)
    
    'Seguis invisible al pasar de mapa
    If (UserList(userindex).flags.Invisible = 1 Or UserList(userindex).flags.Oculto = 1) And (Not UserList(userindex).flags.AdminInvisible = 1) Then
        Call SendToUserArea(userindex, "NOVER" & UserList(userindex).char.CharIndex & ",1", EncriptarProtocolosCriticos)
    End If
    
    If FX And UserList(userindex).flags.AdminInvisible = 0 Then 'FX
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_WARP)
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & FXIDs.FXWARP & ",0")
    End If
    
    Call WarpMascotas(userindex)
End Sub

Sub UpdateUserMap(ByVal userindex As Integer)

Dim Map As Integer
Dim x As Integer
Dim Y As Integer

'EnviarNoche UserIndex

On Error GoTo 0

Map = UserList(userindex).pos.Map

For Y = YMinMapSize To YMaxMapSize
    For x = XMinMapSize To XMaxMapSize
        If MapData(Map, x, Y).userindex > 0 And userindex <> MapData(Map, x, Y).userindex Then
            Call MakeUserChar(SendTarget.toIndex, userindex, 0, MapData(Map, x, Y).userindex, Map, x, Y)
#If SeguridadAlkon Then
            If EncriptarProtocolosCriticos Then
                If UserList(MapData(Map, x, Y).userindex).flags.Invisible = 1 Or UserList(MapData(Map, x, Y).userindex).flags.Oculto = 1 Then Call SendCryptedData(SendTarget.toIndex, userindex, 0, "NOVER" & UserList(MapData(Map, x, Y).userindex).char.CharIndex & ",1")
            Else
#End If
                If UserList(MapData(Map, x, Y).userindex).flags.Invisible = 1 Or UserList(MapData(Map, x, Y).userindex).flags.Oculto = 1 Then Call SendData(SendTarget.toIndex, userindex, 0, "NOVER" & UserList(MapData(Map, x, Y).userindex).char.CharIndex & ",1")
#If SeguridadAlkon Then
            End If
#End If
        End If

        If MapData(Map, x, Y).NpcIndex > 0 Then
            Call MakeNPCChar(SendTarget.toIndex, userindex, 0, MapData(Map, x, Y).NpcIndex, Map, x, Y)
        End If

        If MapData(Map, x, Y).OBJInfo.ObjIndex > 0 Then
            If ObjData(MapData(Map, x, Y).OBJInfo.ObjIndex).OBJType <> eOBJType.otArboles Then
                Call MakeObj(SendTarget.toIndex, userindex, 0, MapData(Map, x, Y).OBJInfo, Map, x, Y)
                If ObjData(MapData(Map, x, Y).OBJInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                          Call Bloquear(SendTarget.toIndex, userindex, 0, Map, x, Y, MapData(Map, x, Y).Blocked)
                          Call Bloquear(SendTarget.toIndex, userindex, 0, Map, x - 1, Y, MapData(Map, x - 1, Y).Blocked)
                End If
            End If
        End If
        
    Next x
Next Y

End Sub


Sub WarpMascotas(ByVal userindex As Integer)
Dim i As Integer

Dim UMascRespawn  As Boolean
Dim miflag As Byte, MascotasReales As Integer
Dim prevMacotaType As Integer

Dim PetTypes(1 To MAXMASCOTAS) As Integer
Dim PetRespawn(1 To MAXMASCOTAS) As Boolean
Dim PetTiempoDeVida(1 To MAXMASCOTAS) As Integer

Dim NroPets As Integer, InvocadosMatados As Integer


NroPets = UserList(userindex).NroMacotas
InvocadosMatados = 0

    'Matamos los invocados
    '[Alejo 18-03-2004]
    For i = 1 To MAXMASCOTAS
        If UserList(userindex).MascotasIndex(i) > 0 Then
            ' si la mascota tiene tiempo de vida > 0 significa q fue invocada.
            If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                Call QuitarNPC(UserList(userindex).MascotasIndex(i))
                UserList(userindex).MascotasIndex(i) = 0
                InvocadosMatados = InvocadosMatados + 1
                NroPets = NroPets - 1
            End If
        End If
    Next i
    
    If InvocadosMatados > 0 Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||Pierdes el control de tus mascotas." & FONTTYPE_INFO)
    End If
    
    For i = 1 To MAXMASCOTAS
        If UserList(userindex).MascotasIndex(i) > 0 Then
            PetRespawn(i) = Npclist(UserList(userindex).MascotasIndex(i)).flags.Respawn = 0
            PetTypes(i) = UserList(userindex).MascotasType(i)
            PetTiempoDeVida(i) = Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia
            Call QuitarNPC(UserList(userindex).MascotasIndex(i))
        End If
    Next i
    
    For i = 1 To MAXMASCOTAS
        If PetTypes(i) > 0 Then
            UserList(userindex).MascotasIndex(i) = SpawnNpc(PetTypes(i), UserList(userindex).pos, False, PetRespawn(i))
            UserList(userindex).MascotasType(i) = PetTypes(i)
            'Controlamos que se sumoneo OK
            If UserList(userindex).MascotasIndex(i) = 0 Then
                UserList(userindex).MascotasIndex(i) = 0
                UserList(userindex).MascotasType(i) = 0
                If UserList(userindex).NroMacotas > 0 Then UserList(userindex).NroMacotas = UserList(userindex).NroMacotas - 1
                Exit Sub
            End If
            Npclist(UserList(userindex).MascotasIndex(i)).MaestroUser = userindex
            Npclist(UserList(userindex).MascotasIndex(i)).Movement = TipoAI.SigueAmo
            Npclist(UserList(userindex).MascotasIndex(i)).Target = 0
            Npclist(UserList(userindex).MascotasIndex(i)).TargetNPC = 0
            Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia = PetTiempoDeVida(i)
            Call FollowAmo(UserList(userindex).MascotasIndex(i))
        End If
    Next i
    
    UserList(userindex).NroMacotas = NroPets

End Sub


Sub RepararMascotas(ByVal userindex As Integer)
Dim i As Integer
Dim MascotasReales As Integer

    For i = 1 To MAXMASCOTAS
      If UserList(userindex).MascotasType(i) > 0 Then MascotasReales = MascotasReales + 1
    Next i
    
    If MascotasReales <> UserList(userindex).NroMacotas Then UserList(userindex).NroMacotas = 0

End Sub

Sub Cerrar_Usuario(ByVal userindex As Integer, Optional ByVal Tiempo As Integer = -1)
    If Tiempo = -1 Then Tiempo = IntervaloCerrarConexion
    
    If UserList(userindex).flags.UserLogged And Not UserList(userindex).Counters.Saliendo Then
        UserList(userindex).Counters.Saliendo = True
        UserList(userindex).Counters.Salir = IIf(UserList(userindex).flags.Privilegios > PlayerType.User Or Not MapInfo(UserList(userindex).pos.Map).Pk, 4, Tiempo)
        
    
        Call SendData(SendTarget.toIndex, userindex, 0, "||Cerrando...Se cerrará el juego en " & UserList(userindex).Counters.Salir & " segundos..." & FONTTYPE_INFO)
    End If
End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal userindex As Integer, ByVal UserIndexDestino As Integer, ByVal NuevoNick As String)
Dim ViejoNick As String
Dim ViejoCharBackup As String

If UserList(UserIndexDestino).flags.UserLogged = False Then Exit Sub
ViejoNick = UserList(UserIndexDestino).name

If FileExist(CharPath & ViejoNick & ".chr", vbNormal) Then
    'hace un backup del char
    ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
    Name CharPath & ViejoNick & ".chr" As ViejoCharBackup
End If

End Sub

Public Sub Empollando(ByVal userindex As Integer)
If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).OBJInfo.ObjIndex > 0 Then
    UserList(userindex).flags.EstaEmpo = 1
Else
    UserList(userindex).flags.EstaEmpo = 0
    UserList(userindex).EmpoCont = 0
End If

End Sub

Sub SendUserStatsTxtOFF(ByVal sendIndex As Integer, ByVal nombre As String)

If FileExist(CharPath & nombre & ".chr", vbArchive) = False Then
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Pj Inexistente" & FONTTYPE_INFO)
Else
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Estadisticas de: " & nombre & FONTTYPE_INFO)
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Nivel: " & GetVar(CharPath & nombre & ".chr", "stats", "elv") & "  EXP: " & GetVar(CharPath & nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & nombre & ".chr", "stats", "elu") & FONTTYPE_INFO)
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Vitalidad: " & GetVar(CharPath & nombre & ".chr", "stats", "minsta") & "/" & GetVar(CharPath & nombre & ".chr", "stats", "maxSta") & FONTTYPE_INFO)
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Salud: " & GetVar(CharPath & nombre & ".chr", "stats", "MinHP") & "/" & GetVar(CharPath & nombre & ".chr", "Stats", "MaxHP") & "  Mana: " & GetVar(CharPath & nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & nombre & ".chr", "Stats", "MaxMAN") & FONTTYPE_INFO)
    
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & GetVar(CharPath & nombre & ".chr", "stats", "MaxHIT") & FONTTYPE_INFO)
    
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Oro: " & GetVar(CharPath & nombre & ".chr", "stats", "GLD") & FONTTYPE_INFO)
    
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Trofeos de Oro: " & GetVar(CharPath & nombre & ".chr", "stats", "TrofOro") & "~255~255~6~0~0~")
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Trofeos de Plata: " & GetVar(CharPath & nombre & ".chr", "stats", "TrofPlata") & "~255~255~251~0~0~")
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Trofeos de Bronce: " & GetVar(CharPath & nombre & ".chr", "stats", "TrofBronce") & "~187~0~0~0~0~")
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Amuletos de Madera: " & GetVar(CharPath & nombre & ".chr", "stats", "TrofMadera") & "~237~207~139~0~0~")
End If
Exit Sub

End Sub
Sub SendUserOROTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
On Error Resume Next
Dim j As Integer
Dim CharFile As String, Tmp As String
Dim ObjInd As Long, ObjCant As Long

CharFile = CharPath & CharName & ".chr"

If FileExist(CharFile, vbNormal) Then
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||" & CharName & FONTTYPE_INFO)
    Call SendData(SendTarget.toIndex, sendIndex, 0, "|| Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco." & FONTTYPE_INFO)
    Else
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Usuario inexistente: " & CharName & FONTTYPE_INFO)
End If

End Sub


