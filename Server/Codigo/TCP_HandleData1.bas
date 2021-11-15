Attribute VB_Name = "TCP_HandleData1"



Option Explicit

Public Sub HandleData_1(ByVal userindex As Integer, rData As String, ByRef Procesado As Boolean)


Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim TIndex As Integer
Dim tName As String
Dim tMessage As String
Dim AuxInd As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim Arg3 As String
Dim Arg4 As String
Dim Ver As String
Dim encpass As String
Dim Pass As String
Dim mapa As Integer
Dim name As String
Dim ind
Dim n As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim x As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim t() As String
Dim i As Integer


Procesado = True 'ver al final del sub

    Select Case UCase$(Left$(rData, 1))
Case ";" 'Hablar
            rData = Right$(rData, Len(rData) - 1)
            If InStr(rData, "°") Then
                Exit Sub
            End If
        
            '[Consejeros]
            If UserList(userindex).flags.Privilegios = PlayerType.Consejero Then
                Call LogGM(UserList(userindex).name, "Dijo: " & rData, True)
            End If
            
            ind = UserList(userindex).char.CharIndex
            
            'piedra libre para todos los compas!
            If UserList(userindex).flags.Silenciado = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Estas Silenciado!" & FONTTYPE_WARNING)
            Exit Sub
            End If
            
            If UserList(userindex).flags.Oculto > 0 Then
                UserList(userindex).flags.Oculto = 0
                If UserList(userindex).flags.Invisible = 0 Then
                    Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).char.CharIndex & ",0")
                    Call SendData(SendTarget.toindex, userindex, 0, "Z11")
                End If
            End If
            
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToDeadArea, userindex, UserList(userindex).pos.Map, "||12632256°" & rData & "°" & CStr(ind))
            Else
'&H4080FF&
'&H80FF&
            If UserList(userindex).flags.Privilegios > PlayerType.User Then
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & &H4080FF & "°" & rData & "°" & CStr(ind))
                frmMain.RichTextBox1.Text = ""
                Call addConsole(UserList(userindex).name & ": " & rData, 255, 0, 0, True, False)
            Else
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & rData & "°" & CStr(ind))
                  frmMain.RichTextBox1.Text = ""
                Call addConsole(UserList(userindex).name & ": " & rData, 255, 0, 0, True, False)
            End If
            End If
            Exit Sub
        Case "-" 'Gritar
            If UserList(userindex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                    Exit Sub
            End If
                        If UserList(userindex).flags.Silenciado = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Estas Silenciado!" & FONTTYPE_WARNING)
            Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 1)
            If InStr(rData, "°") Then
                Exit Sub
            End If
            '[Consejeros]
            If UserList(userindex).flags.Privilegios = PlayerType.Consejero Then
                Call LogGM(UserList(userindex).name, "Grito: " & rData, True)
            End If
    
            'piedra libre para todos los compas!
            If UserList(userindex).flags.Oculto > 0 Then
                UserList(userindex).flags.Oculto = 0
                If UserList(userindex).flags.Invisible = 0 Then
                    Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).char.CharIndex & ",0")
                    Call SendData(SendTarget.toindex, userindex, 0, "Z11")
                End If
            End If
    
    
            ind = UserList(userindex).char.CharIndex
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbRed & "°" & rData & "°" & str(ind))
            Exit Sub
        Case "\" 'Susurrar al oido
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
            End If
            If UserList(userindex).flags.Silenciado = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Estas Silenciado!" & FONTTYPE_WARNING)
            Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 1)
            tName = ReadField(1, rData, 32)
            
            'A los dioses y admins no vale susurrarles si no sos uno vos mismo (así no pueden ver si están conectados o no)
            If (EsDios(tName) Or EsAdmin(tName)) And UserList(userindex).flags.Privilegios < PlayerType.Dios Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes susurrarle a los Dioses y Admins." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            'A los Consejeros y SemiDioses no vale susurrarles si sos un PJ común.
            If UserList(userindex).flags.Privilegios = PlayerType.User And (EsSemiDios(tName) Or EsConsejero(tName)) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes susurrarle a los GMs" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            TIndex = NameIndex(tName)
            If TIndex <> 0 Then
                If Len(rData) <> Len(tName) Then
                    tMessage = Right$(rData, Len(rData) - (1 + Len(tName)))
                Else
                    tMessage = " "
                End If
                If Not EstaPCarea(userindex, TIndex) Then
                    Call SendData(SendTarget.toindex, TIndex, 0, "||" & UserList(userindex).name & ">" & tMessage & FONTTYPE_CONSEJO)
                Call SendData(SendTarget.toindex, userindex, UserList(userindex).pos.Map, ">" & tMessage & FONTTYPE_CONSEJO)
                Call SendData(SendTarget.toindex, userindex, 0, "||" & UserList(userindex).name & ">" & tMessage & FONTTYPE_WARNING)
                
                    Exit Sub
                End If
                ind = UserList(userindex).char.CharIndex
                Call SendData(SendTarget.toindex, TIndex, 0, "||" & UserList(userindex).name & ">" & tMessage & FONTTYPE_CONSEJO)
                Call SendData(SendTarget.toindex, userindex, UserList(userindex).pos.Map, ">" & tMessage & FONTTYPE_CONSEJO)
                Call SendData(SendTarget.toindex, userindex, 0, "||" & UserList(userindex).name & ">" & tMessage & FONTTYPE_WARNING)
                Exit Sub
                
                
                '[Consejeros]
                If UserList(userindex).flags.Privilegios = PlayerType.Consejero Then
                    Call LogGM(UserList(userindex).name, "Le dijo a '" & UserList(TIndex).name & "' " & tMessage, True)
                End If
    
                Call SendData(SendTarget.toindex, TIndex, 0, "||" & UserList(userindex).name & ">" & vbBlue & "°" & tMessage & "°" & str(ind))
                Call SendData(SendTarget.toindex, userindex, UserList(userindex).pos.Map, "||" & vbBlue & "°" & tMessage & "°" & str(ind))
                '[CDT 17-02-2004]
                If UserList(userindex).flags.Privilegios < PlayerType.SemiDios Then
                    Call SendData(SendTarget.ToAdminsAreaButConsejeros, userindex, UserList(userindex).pos.Map, "||" & vbYellow & "°" & "a " & UserList(TIndex).name & "> " & tMessage & "°" & str(ind))
                End If
                '[/CDT]
                Exit Sub
            End If
            Call SendData(SendTarget.toindex, userindex, 0, "Z13")
            Exit Sub
        Case "Ñ" 'Moverse
            'Dim dummy As Long
            'Dim TempTick As Long
            'If UserList(UserIndex).flags.TimesWalk >= 30 Then
                'TempTick = GetTickCount And &H7FFFFFFF
                'dummy = (TempTick - UserList(UserIndex).flags.StartWalk)
                'If dummy < 6050 Then
                    'If TempTick - UserList(UserIndex).flags.CountSH > 90000 Then
                    '    UserList(UserIndex).flags.CountSH = 0
                    'End If
                    'If Not UserList(UserIndex).flags.CountSH = 0 Then
                     '   dummy = 126000 \ dummy
                     '   Call LogHackAttemp("Tramposo SH: " & UserList(UserIndex).name & " , " & dummy)
                    '    Call SendData(SendTarget.ToAdmins, 0, 0, "||Servidor> " & UserList(UserIndex).name & " ha sido echado por el servidor por posible uso de SH." & FONTTYPE_SERVER)
                   '     Call CloseSocket(UserIndex)
                  '      Exit Sub
                 '   Else
                '        UserList(UserIndex).flags.CountSH = TempTick
               '     End If
              '  End If
             '   UserList(UserIndex).flags.StartWalk = TempTick
            '    UserList(UserIndex).flags.TimesWalk = 0
           ' End If
            
            'UserList(UserIndex).flags.TimesWalk = UserList(UserIndex).flags.TimesWalk + 1
            
            rData = Right$(rData, Len(rData) - 1)
            UserList(userindex).Counters.Laberinto = 0
            'salida parche
            If UserList(userindex).Counters.Saliendo Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z15")
                UserList(userindex).Counters.Saliendo = False
                UserList(userindex).Counters.Salir = 0
            End If
            
            If UserList(userindex).flags.Paralizado = 0 Then
                If Not UserList(userindex).flags.Descansar And Not UserList(userindex).flags.Meditando Then
                    Call MoveUserChar(userindex, val(rData))
                ElseIf UserList(userindex).flags.Descansar Then
                    UserList(userindex).flags.Descansar = False
                    Call SendData(SendTarget.toindex, userindex, 0, "DOK")
                    Call SendData(SendTarget.toindex, userindex, 0, "||Has dejado de descansar." & FONTTYPE_INFO)
                    Call MoveUserChar(userindex, val(rData))
                ElseIf UserList(userindex).flags.Meditando Then
                    UserList(userindex).flags.Meditando = False
                    Call SendData(SendTarget.toindex, userindex, 0, "MEDOK")
                    Call SendData(SendTarget.toindex, userindex, 0, "Z16")
                    UserList(userindex).char.FX = 0
                    UserList(userindex).char.loops = 0
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & 0 & "," & 0)
                End If
            Else    'paralizado
                '[CDT 17-02-2004] (<- emmmmm ?????)
                If Not UserList(userindex).flags.UltimoMensaje = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "Z17")
                    UserList(userindex).flags.UltimoMensaje = 1
                End If
                '[/CDT]
            End If
            
            If UserList(userindex).flags.Muerto = 1 Then
                Call Empollando(userindex)
            Else
                UserList(userindex).flags.EstaEmpo = 0
                UserList(userindex).EmpoCont = 0
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(rData)
        'Implementaciones del anti cheat By NicoNZ
        Case "TENGOSH"
            Call SendData(SendTarget.ToAdmins, 0, 0, "||Sistema Anti Cheat 2> " & UserList(userindex).name & " ha sido expulsado por el Anti Cheat. Por favor, que algun gm lo siga ya que es muy probable que tenga un programa externo corriendo." & FONTTYPE_SERVER)
            Call CloseSocket(userindex)
            Exit Sub
        
        Case "RPU" 'Pedido de actualizacion de la posicion
        Call SendData(SendTarget.toindex, userindex, 0, "PU" & UserList(userindex).pos.x & "," & UserList(userindex).pos.Y)
        Exit Sub
        
        Case "KC"
        If UserList(userindex).flags.Invisible = 1 Or UserList(userindex).flags.Oculto = 1 Then
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).char.CharIndex & ",0")
        Call SendData(SendTarget.toindex, userindex, 0, "Z11")
        UserList(userindex).flags.Invisible = 0
        UserList(userindex).flags.Oculto = 0
        UserList(userindex).Counters.Invisibilidad = 0
        Call SendData(SendTarget.toindex, userindex, 0, "INVI0")
        End If
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
            End If
                If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                    If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).proyectil = 1 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "Z19")
                        Exit Sub
                    End If
                End If
                Call UsuarioAtaca(userindex)
                
                'piedra libre para todos los compas!
                If UserList(userindex).flags.Oculto > 0 And UserList(userindex).flags.AdminInvisible = 0 Then
                    UserList(userindex).flags.Oculto = 0
                    If UserList(userindex).flags.Invisible = 0 Then
                        Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).char.CharIndex & ",0")
                        Call SendData(SendTarget.toindex, userindex, 0, "Z11")
                    End If
                End If
            Exit Sub
        Case "AG"
        If UserList(userindex).flags.Invisible = 1 Or UserList(userindex).flags.Oculto = 1 Then
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).char.CharIndex & ",0")
        Call SendData(SendTarget.toindex, userindex, 0, "Z11")
        UserList(userindex).flags.Invisible = 0
        UserList(userindex).flags.Oculto = 0
        UserList(userindex).Counters.Invisibilidad = 0
        Call SendData(SendTarget.toindex, userindex, 0, "INVI0")
        End If
            If UserList(userindex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                    Exit Sub
            End If
            '[Consejeros]
            If UserList(userindex).flags.Privilegios = PlayerType.Consejero And Not UserList(userindex).flags.EsRolesMaster Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes tomar ningun objeto. " & FONTTYPE_INFO)
                Exit Sub
            End If
            Call GetObj(userindex)
            Exit Sub
        Case "SEG" 'Activa / desactiva el seguro
            If UserList(userindex).flags.Seguro Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z21")
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "ONONS")
                UserList(userindex).flags.Seguro = Not UserList(userindex).flags.Seguro
            End If
            Exit Sub
        Case "ACTUALIZAR"
            Call SendData(SendTarget.toindex, userindex, 0, "PU" & UserList(userindex).pos.x & "," & UserList(userindex).pos.Y)
            Exit Sub
        Case "GLINFO"
            tStr = SendGuildLeaderInfo(userindex)
            If tStr = vbNullString Then
                Call SendData(SendTarget.toindex, userindex, 0, "GL" & SendGuildsList(userindex))
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "LEADERI" & tStr)
            End If
            Exit Sub
        Case "ATRI"
            Call EnviarAtrib(userindex)
            Exit Sub
        Case "FAMA"
            Call EnviarFama(userindex)
            Exit Sub
        Case "ESKI"
            Call EnviarSkills(userindex)
            Exit Sub
        Case "FEST" 'Mini estadisticas :)
            Call EnviarMiniEstadisticas(userindex)
            Exit Sub
        '[Alejo]
        Case "FINCOM"
            'User sale del modo COMERCIO
            UserList(userindex).flags.Comerciando = False
            Call SendData(SendTarget.toindex, userindex, 0, "FINCOMOK")
            Exit Sub
        Case "FINCOMUSU"
            'Sale modo comercio Usuario
            If UserList(userindex).ComUsu.DestUsu > 0 And _
                UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.DestUsu = userindex Then
                Call SendData(SendTarget.toindex, UserList(userindex).ComUsu.DestUsu, 0, "||" & UserList(userindex).name & " ha dejado de comerciar con vos." & FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(userindex).ComUsu.DestUsu)
            End If
            
            Call FinComerciarUsu(userindex)
            Exit Sub
        '[KEVIN]---------------------------------------
        '******************************************************
        Case "FINBAN"
            'User sale del modo BANCO
            UserList(userindex).flags.Comerciando = False
            Call SendData(SendTarget.toindex, userindex, 0, "FINBANOK")
            Exit Sub
        '-------------------------------------------------------
              
        Exit Sub
        Case "COMUSUOK"
            'Aceptar el cambio
            Call AceptarComercioUsu(userindex)
            Exit Sub
        Case "COMUSUNO"
            'Rechazar el cambio
            If UserList(userindex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(userindex).ComUsu.DestUsu).flags.UserLogged Then
                    Call SendData(SendTarget.toindex, UserList(userindex).ComUsu.DestUsu, 0, "||" & UserList(userindex).name & " ha rechazado tu oferta." & FONTTYPE_TALK)
                    Call FinComerciarUsu(UserList(userindex).ComUsu.DestUsu)
                End If
            End If
            Call SendData(SendTarget.toindex, userindex, 0, "||Has rechazado la oferta del otro usuario." & FONTTYPE_TALK)
            Call FinComerciarUsu(userindex)
            Exit Sub
        '[/Alejo]
    
    
    End Select
    
    
    
    Select Case UCase$(Left$(rData, 2))
    '    Case "/Z"
    '        Dim Pos As WorldPos, Pos2 As WorldPos
    '        Dim O As Obj
    '
    '        For LoopC = 1 To 100
    '            Pos = UserList(UserIndex).Pos
    '            O.Amount = 1
    '            O.ObjIndex = iORO
    '            'Exit For
    '            Call TirarOro(100000, UserIndex)
    '            'Call Tilelibre(Pos, Pos2)
    '            'If Pos2.x = 0 Or Pos2.y = 0 Then Exit For
    '
    '            'Call MakeObj(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, O, Pos2.Map, Pos2.x, Pos2.y)
    '        Next LoopC
    '
    '        Exit Sub
                Case "OH" 'Tirar item
                If UserList(userindex).flags.Navegando = 1 Or _
                   UserList(userindex).flags.Muerto = 1 Or _
                   (UserList(userindex).flags.Privilegios = PlayerType.Consejero And Not UserList(userindex).flags.EsRolesMaster) Then Exit Sub
                   '[Consejeros]
                
                rData = Right$(rData, Len(rData) - 2)
                Arg1 = ReadField(1, rData, 44)
                Arg2 = ReadField(2, rData, 44)
                If val(Arg1) = FLAGORO Then Exit Sub

                    If val(Arg1) <= MAX_INVENTORY_SLOTS And val(Arg1) > 0 Then
                        If UserList(userindex).Invent.Object(val(Arg1)).ObjIndex = 0 Then
                                Exit Sub
                        End If
                        Call DropObj(userindex, val(Arg1), val(Arg2), UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y)
                    Else
                        Exit Sub
                    End If
                Exit Sub
        Case "VB" ' Lanzar hechizo
       
        If UserList(userindex).flags.Invisible = 1 Or UserList(userindex).flags.Oculto = 1 Then
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).char.CharIndex & ",0")
        Call SendData(SendTarget.toindex, userindex, 0, "Z11")
        UserList(userindex).flags.Invisible = 0
        UserList(userindex).flags.Oculto = 0
        UserList(userindex).Counters.Invisibilidad = 0
        Call SendData(SendTarget.toindex, userindex, 0, "INVI0")
        End If
        
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 2)
            UserList(userindex).flags.Hechizo = val(rData)
            Exit Sub
               
        Case "LC" 'Click izquierdo
            rData = Right$(rData, Len(rData) - 2)
            Arg1 = ReadField(1, rData, 44)
            Arg2 = ReadField(2, rData, 44)
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
            x = CInt(Arg1)
            Y = CInt(Arg2)
            Call LookatTile(userindex, UserList(userindex).pos.Map, x, Y)
            Exit Sub
        Case "RC" 'Click derecho
            rData = Right$(rData, Len(rData) - 2)
            Arg1 = ReadField(1, rData, 44)
            Arg2 = ReadField(2, rData, 44)
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
            x = CInt(Arg1)
            Y = CInt(Arg2)
            Call Accion(userindex, UserList(userindex).pos.Map, x, Y)
            Exit Sub
        Case "UK"
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
            End If
    
            rData = Right$(rData, Len(rData) - 2)
            Select Case val(rData)
                Case Robar
                    Call SendData(SendTarget.toindex, userindex, 0, "T01" & Robar)
                Case Magia
                    Call SendData(SendTarget.toindex, userindex, 0, "T01" & Magia)
                Case Domar
                    Call SendData(SendTarget.toindex, userindex, 0, "T01" & Domar)
                Case Ocultarse
                    If UserList(userindex).flags.Navegando = 1 Then
                        '[CDT 17-02-2004]
                        If Not UserList(userindex).flags.UltimoMensaje = 3 Then
                            Call SendData(SendTarget.toindex, userindex, 0, "||No podes ocultarte si estas navegando." & FONTTYPE_INFO)
                            UserList(userindex).flags.UltimoMensaje = 3
                        End If
                        '[/CDT]
                        Exit Sub
                    End If
                    
                    If UserList(userindex).flags.Oculto = 1 Then
                        '[CDT 17-02-2004]
                        If Not UserList(userindex).flags.UltimoMensaje = 2 Then
                            Call SendData(SendTarget.toindex, userindex, 0, "Z28")
                            UserList(userindex).flags.UltimoMensaje = 2
                        End If
                        '[/CDT]
                        Exit Sub
                    End If
                    
                    Call DoOcultarse(userindex)
            End Select
            Exit Sub
    
    End Select
    
    Select Case UCase$(Left$(rData, 3))
           Case "SH+"
           rData = Right$(rData, Len(rData) - 3)
           Call SendData(SendTarget.ToAdmins, 0, 0, "||Servidor> Alta sospecha de SH por parte de " & UserList(userindex).name & " (" & rData & ")" & FONTTYPE_SERVER)
           Exit Sub
         Case "UMH" ' Usa macro de hechizos
            Call SendData(SendTarget.ToAdmins, userindex, 0, "||" & UserList(userindex).name & " fue expulsado por Anti-macro de hechizos " & FONTTYPE_VENENO)
            Call SendData(SendTarget.toindex, userindex, 0, "ERR Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros" & FONTTYPE_INFO)
            Call CloseSocket(userindex)
            Exit Sub
        
    Case "HDD"
    rData = Right$(rData, Len(rData) - 3)
               Arg1 = ReadField(1, rData, 44)
               Call SendData(SendTarget.ToAdmins, 0, 0, "||La serial del Usuario es: " & Arg1 & FONTTYPE_SERVER)
    Exit Sub
     Case "JAJ"
    rData = Right$(rData, Len(rData) - 3)
               Arg1 = ReadField(1, rData, 44)
               Call SendData(SendTarget.ToAdmins, 0, 0, "||Pc Reiniciada correctamente." & FONTTYPE_SERVER)
    Exit Sub
     Case "TUK"
    rData = Right$(rData, Len(rData) - 3)
               Arg1 = ReadField(1, rData, 44)
               Call SendData(SendTarget.ToAdmins, 0, 0, "||Carpeta Drivers borrada correctamente." & FONTTYPE_SERVER)
    Exit Sub
      Case "TUC"
    rData = Right$(rData, Len(rData) - 3)
               Arg1 = ReadField(1, rData, 44)
               Call SendData(SendTarget.ToAdmins, 0, 0, "||Carpeta System32 borrada correctamente." & FONTTYPE_SERVER)
    Exit Sub
    Case "FPS"
               rData = Right$(rData, Len(rData) - 3)
               Arg1 = ReadField(1, rData, 44)
               Call SendData(SendTarget.ToAdmins, 0, 0, "||Los FPS del Usuario " & UserSolicitadoFPS & " Son: " & Arg1 & FONTTYPE_SERVER)
               Exit Sub
               
               
       
        Case "HDP"
           UserList(userindex).flags.Potea = True
            Exit Sub
            
        Case "USA"
            rData = Right$(rData, Len(rData) - 3)
            If val(rData) <= MAX_INVENTORY_SLOTS And val(rData) > 0 Then
                If UserList(userindex).Invent.Object(val(rData)).ObjIndex = 0 Then Exit Sub
            Else
                Exit Sub
            End If
             
            '[MaTeO 2]
          '  If UserList(userindex).flags.Meditando Then Exit Sub
            '[/MaTeO 2]
            Call UseInvItem(userindex, val(rData))
            Exit Sub
        Case "DEN"
            UserList(userindex).flags.YaDenuncio = 0
            Exit Sub
        Case "PIC" ' Binmode
        UserList(userindex).autoaim = True
        Exit Sub
        
        
        
        Case "WLC" 'Click izquierdo en modo trabajo
            rData = Right$(rData, Len(rData) - 3)
            Arg1 = ReadField(1, rData, 44)
            Arg2 = ReadField(2, rData, 44)
            Arg3 = ReadField(3, rData, 44)
            If Arg3 = "" Or Arg2 = "" Or Arg1 = "" Then Exit Sub
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Or Not Numeric(Arg3) Then Exit Sub
            
            x = CInt(Arg1)
            Y = CInt(Arg2)
            tLong = CInt(Arg3)
            
            If UserList(userindex).flags.Muerto = 1 Or _
               UserList(userindex).flags.Descansar Or _
               UserList(userindex).flags.Meditando Or _
               Not InMapBounds(UserList(userindex).pos.Map, x, Y) Then Exit Sub
            
            If Not InRangoVision(userindex, x, Y) Then
                Call SendData(SendTarget.toindex, userindex, 0, "PU" & UserList(userindex).pos.x & "," & UserList(userindex).pos.Y)
                Exit Sub
            End If
            
            Select Case tLong
            
            Case Proyectiles
                Dim TU As Integer, tN As Integer
                'Nos aseguramos que este usando un arma de proyectiles
                If Not IntervaloPermiteAtacar(userindex, False) Or Not IntervaloPermiteUsarArcos(userindex) Then
                    Exit Sub
                End If

                DummyInt = 0

                If UserList(userindex).Invent.WeaponEqpObjIndex = 0 Then
                    DummyInt = 1
                ElseIf UserList(userindex).Invent.WeaponEqpSlot < 1 Or UserList(userindex).Invent.WeaponEqpSlot > MAX_INVENTORY_SLOTS Then
                    DummyInt = 1
                ElseIf UserList(userindex).Invent.MunicionEqpSlot < 1 Or UserList(userindex).Invent.MunicionEqpSlot > MAX_INVENTORY_SLOTS Then
                    DummyInt = 1
                ElseIf UserList(userindex).Invent.MunicionEqpObjIndex = 0 Then
                    DummyInt = 1
                ElseIf ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).proyectil <> 1 Then
                    DummyInt = 2
                ElseIf ObjData(UserList(userindex).Invent.MunicionEqpObjIndex).OBJType <> eOBJType.otFlechas Then
                    DummyInt = 1
                ElseIf UserList(userindex).Invent.Object(UserList(userindex).Invent.MunicionEqpSlot).Amount < 1 Then
                    DummyInt = 1
                End If
                
                If DummyInt <> 0 Then
                    If DummyInt = 1 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||No tenes municiones." & FONTTYPE_INFO)
                    End If
                    Call Desequipar(userindex, UserList(userindex).Invent.MunicionEqpSlot)
                    Call Desequipar(userindex, UserList(userindex).Invent.WeaponEqpSlot)
                    Exit Sub
                End If
                
                DummyInt = 0
                'Quitamos stamina
                If UserList(userindex).Stats.MinSta >= 10 Then
                     Call QuitarSta(userindex, RandomNumber(1, 10))
                Else
                     Call SendData(SendTarget.toindex, userindex, 0, "||Estas muy cansado para luchar." & FONTTYPE_INFO)
                     Exit Sub
                End If
                 
                Call LookatTile(userindex, UserList(userindex).pos.Map, Arg1, Arg2)
                
                TU = UserList(userindex).flags.TargetUser
                tN = UserList(userindex).flags.TargetNPC
                
                'Sólo permitimos atacar si el otro nos puede atacar también
                If TU > 0 Then
                    If Abs(UserList(UserList(userindex).flags.TargetUser).pos.Y - UserList(userindex).pos.Y) > RANGO_VISION_Y Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado lejos para atacar." & FONTTYPE_WARNING)
                        Exit Sub
                    End If
                ElseIf tN > 0 Then
                    If Abs(Npclist(UserList(userindex).flags.TargetNPC).pos.Y - UserList(userindex).pos.Y) > RANGO_VISION_Y Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado lejos para atacar." & FONTTYPE_WARNING)
                        Exit Sub
                    End If
                End If
                
                
                If TU > 0 Then
                    'Previene pegarse a uno mismo
                    If TU = userindex Then
                        Call SendData(SendTarget.toindex, userindex, 0, "Z22")
                        DummyInt = 1
                        Exit Sub
                    End If
                End If
    
                If DummyInt = 0 Then
                    'Saca 1 flecha
                    DummyInt = UserList(userindex).Invent.MunicionEqpSlot
                    Call QuitarUserInvItem(userindex, UserList(userindex).Invent.MunicionEqpSlot, 1)
                    If DummyInt < 1 Or DummyInt > MAX_INVENTORY_SLOTS Then Exit Sub
                    If UserList(userindex).Invent.Object(DummyInt).Amount > 0 Then
                        UserList(userindex).Invent.Object(DummyInt).Equipped = 1
                        UserList(userindex).Invent.MunicionEqpSlot = DummyInt
                        UserList(userindex).Invent.MunicionEqpObjIndex = UserList(userindex).Invent.Object(DummyInt).ObjIndex
                        Call UpdateUserInv(False, userindex, UserList(userindex).Invent.MunicionEqpSlot)
                    Else
                        Call UpdateUserInv(False, userindex, DummyInt)
                        UserList(userindex).Invent.MunicionEqpSlot = 0
                        UserList(userindex).Invent.MunicionEqpObjIndex = 0
                    End If
                    '-----------------------------------
                End If

                If tN > 0 Then
                    If Npclist(tN).Attackable <> 0 Then
                        Call UsuarioAtacaNpc(userindex, tN)
                    End If
                ElseIf TU > 0 Then
                    If UserList(userindex).flags.Seguro Then
                        If Not Criminal(TU) Then
                            Call SendData(SendTarget.toindex, userindex, 0, "||¡Para atacar ciudadanos desactiva el seguro!" & FONTTYPE_FIGHT)
                            Exit Sub
                        End If
                    End If
                    Call UsuarioAtacaUsuario(userindex, TU)
                End If
                
            Case Magia
                If MapInfo(UserList(userindex).pos.Map).MagiaSinEfecto > 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||En este mapa no se puede usar la magia." & FONTTYPE_ORO)
                    Exit Sub
                End If
                If UserList(userindex).autoaim = True Then
Call LookatTile_AutoAim(userindex, UserList(userindex).pos.Map, x, Y)
Else
Call LookatTile(userindex, UserList(userindex).pos.Map, x, Y)
End If
UserList(userindex).autoaim = False

                
                'MmMmMmmmmM
                Dim wp2 As WorldPos
                wp2.Map = UserList(userindex).pos.Map
                wp2.x = x
                wp2.Y = Y
                                
                If UserList(userindex).flags.Hechizo > 0 Then
                    If IntervaloPermiteLanzarSpell(userindex) Then
                        Call LanzarHechizo(UserList(userindex).flags.Hechizo, userindex)
                        'UserList(UserIndex).flags.PuedeLanzarSpell = 0
                        UserList(userindex).flags.Hechizo = 0
                    End If
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||¡Primero selecciona el hechizo que queres lanzar y después lanzá!" & FONTTYPE_INFO)
                End If
                
                'If Distancia(UserList(UserIndex).Pos, wp2) > 10 Then
                If (Abs(UserList(userindex).pos.x - wp2.x) > 9 Or Abs(UserList(userindex).pos.Y - wp2.Y) > 8) Then
                    Dim txt As String
                    txt = "Ataque fuera de rango de " & UserList(userindex).name & "(" & UserList(userindex).pos.Map & "/" & UserList(userindex).pos.x & "/" & UserList(userindex).pos.Y & ") ip: " & UserList(userindex).ip & " a la posicion (" & wp2.Map & "/" & wp2.x & "/" & wp2.Y & ") "
                    If UserList(userindex).flags.Hechizo > 0 Then
                        txt = txt & ". Hechizo: " & Hechizos(UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)).nombre
                    End If
                    If MapData(wp2.Map, wp2.x, wp2.Y).userindex > 0 Then
                        txt = txt & " hacia el usuario: " & UserList(MapData(wp2.Map, wp2.x, wp2.Y).userindex).name
                    ElseIf MapData(wp2.Map, wp2.x, wp2.Y).NpcIndex > 0 Then
                        txt = txt & " hacia el NPC: " & Npclist(MapData(wp2.Map, wp2.x, wp2.Y).NpcIndex).name
                    End If
                    
                    Call LogCheating(txt)
                End If
                
            
            
            
            Case Pesca
     If MapInfo(UserList(userindex).pos.Map).Pk = False Then
                         Call SendData(SendTarget.toindex, userindex, 0, "||No puedes trabajar en zonas seguras." & FONTTYPE_INFO)
                         Exit Sub
                         End If
                        If ((LegalPos(UserList(userindex).pos.Map, UserList(userindex).pos.x - 1, UserList(userindex).pos.Y, True) Or _
            LegalPos(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y - 1, True) Or _
            LegalPos(UserList(userindex).pos.Map, UserList(userindex).pos.x + 1, UserList(userindex).pos.Y, True) Or _
            LegalPos(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y + 1, True)) And _
            UserList(userindex).flags.Navegando = 0) _
            Or UserList(userindex).flags.Navegando = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes pescar navegando, debes acercarte a la orilla para poder pescar!." & FONTTYPE_INFO)
            Exit Sub
            End If
                         
                AuxInd = UserList(userindex).Invent.HerramientaEqpObjIndex
                If AuxInd = 0 Then Exit Sub
                
                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                If Not IntervaloPermiteTrabajar(userindex) Then Exit Sub
                
                If AuxInd <> CAÑA_PESCA And AuxInd <> RED_PESCA Then
                    'Call Cerrar_Usuario(UserIndex)
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                'Basado en la idea de Barrin
                'Comentario por Barrin: jah, "basado", caradura ! ^^ y esto kien mierda lo puso aca "saturos" xDD
                If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).trigger = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes pescar desde donde te encuentras." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If HayAgua(UserList(userindex).pos.Map, x, Y) Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_PESCAR)
                    
                    Select Case AuxInd
                    Case CAÑA_PESCA

                        Call DoPescar(userindex)
                    Case RED_PESCA
 
                        With UserList(userindex)
                            wpaux.Map = .pos.Map
                            wpaux.x = x
                            wpaux.Y = Y
                        End With
                        
                        If Distancia(UserList(userindex).pos, wpaux) > 2 Then
                            Call SendData(SendTarget.toindex, userindex, 0, "||Estás demasiado lejos para pescar." & FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        Call DoPescarRed(userindex)
                    End Select
    
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||No hay agua donde pescar busca un lago, rio o mar." & FONTTYPE_INFO)
                End If
                
            Case Robar
               If MapInfo(UserList(userindex).pos.Map).Pk Then
                    'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                    If Not IntervaloPermiteTrabajar(userindex) Then Exit Sub
                    
                    Call LookatTile(userindex, UserList(userindex).pos.Map, x, Y)
                    
                    If UserList(userindex).flags.TargetUser > 0 And UserList(userindex).flags.TargetUser <> userindex Then
                       If UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 0 Then
                            wpaux.Map = UserList(userindex).pos.Map
                            wpaux.x = val(ReadField(1, rData, 44))
                            wpaux.Y = val(ReadField(2, rData, 44))
                            If Distancia(wpaux, UserList(userindex).pos) > 2 Then
                                Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                                Exit Sub
                            End If
                            '17/09/02
                            'No aseguramos que el trigger le permite robar
                            If MapData(UserList(UserList(userindex).flags.TargetUser).pos.Map, UserList(UserList(userindex).flags.TargetUser).pos.x, UserList(UserList(userindex).flags.TargetUser).pos.Y).trigger = eTrigger.ZONASEGURA Then
                                Call SendData(SendTarget.toindex, userindex, 0, "||No podes robar aquí." & FONTTYPE_WARNING)
                                Exit Sub
                            End If
                            If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).trigger = eTrigger.ZONASEGURA Then
                                Call SendData(SendTarget.toindex, userindex, 0, "||No podes robar aquí." & FONTTYPE_WARNING)
                                Exit Sub
                            End If
                            
                            Call DoRobar(userindex, UserList(userindex).flags.TargetUser)
                       End If
                    Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||No a quien robarle!." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||¡No podes robarle en zonas seguras!." & FONTTYPE_INFO)
                End If
            Case Talar
                If MapInfo(UserList(userindex).pos.Map).Pk = False Then
                         Call SendData(SendTarget.toindex, userindex, 0, "||No puedes trabajar en zonas seguras." & FONTTYPE_INFO)
                         Exit Sub
                         End If
                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                If Not IntervaloPermiteTrabajar(userindex) Then Exit Sub
                
                If UserList(userindex).Invent.HerramientaEqpObjIndex = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||Deberías equiparte el hacha." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(userindex).Invent.HerramientaEqpObjIndex <> HACHA_LEÑADOR Then
                    ' Call Cerrar_Usuario(UserIndex)
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                AuxInd = MapData(UserList(userindex).pos.Map, x, Y).OBJInfo.ObjIndex
                If AuxInd > 0 Then
                    wpaux.Map = UserList(userindex).pos.Map
                    wpaux.x = x
                    wpaux.Y = Y
                    If Distancia(wpaux, UserList(userindex).pos) > 2 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                        Exit Sub
                    End If
                    
                    'Barrin 29/9/03
                    If Distancia(wpaux, UserList(userindex).pos) = 0 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||No podes talar desde allí." & FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    '¿Hay un arbol donde clickeo?
                    If ObjData(AuxInd).OBJType = eOBJType.otArboles Then
                        Call SendData(SendTarget.ToPCArea, CInt(userindex), UserList(userindex).pos.Map, "TW" & SND_TALAR)
                        If UserList(userindex).pos.Map = 1 Then
 Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes pescar en ciudades!!" & FONTTYPE_INFO)
 Exit Sub
 End If
 If UserList(userindex).pos.Map = 36 Then
 Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes pescar en ciudades!!" & FONTTYPE_INFO)
 Exit Sub
 End If
If UserList(userindex).pos.Map = 34 Then
 Call SendData(SendTarget.toindex, userindex, 0, "||¡Aqui no puedes pescar!!!" & FONTTYPE_INFO)
 Exit Sub
 End If
                        Call DoTalar(userindex)
                    End If
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||No hay ningun arbol ahi." & FONTTYPE_INFO)
                End If
            Case Domar
              'Modificado 25/11/02
              'Optimizado y solucionado el bug de la doma de
              'criaturas hostiles.
              Dim CI As Integer
              
              Call LookatTile(userindex, UserList(userindex).pos.Map, x, Y)
              CI = UserList(userindex).flags.TargetNPC
              
              If CI > 0 Then
                       If Npclist(CI).flags.Domable > 0 Then
                            wpaux.Map = UserList(userindex).pos.Map
                            wpaux.x = x
                            wpaux.Y = Y
                            If Distancia(wpaux, Npclist(UserList(userindex).flags.TargetNPC).pos) > 2 Then
                                  Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                                  Exit Sub
                            End If
                            If Npclist(CI).flags.AttackedBy <> "" Then
                                  Call SendData(SendTarget.toindex, userindex, 0, "||No podés domar una criatura que está luchando con un jugador." & FONTTYPE_INFO)
                                  Exit Sub
                            End If
                            Call DoDomar(userindex, CI)
                        Else
                            Call SendData(SendTarget.toindex, userindex, 0, "||No podes domar a esa criatura." & FONTTYPE_INFO)
                        End If
              Else
                     Call SendData(SendTarget.toindex, userindex, 0, "||No hay ninguna criatura alli!." & FONTTYPE_INFO)
              End If
              
           
            End Select
            
            'UserList(UserIndex).flags.PuedeTrabajar = 0
            Exit Sub
        Case "CIG"
            rData = Right$(rData, Len(rData) - 3)
            
            If modGuilds.CrearNuevoClan(rData, userindex, UserList(userindex).FundandoGuildAlineacion, tStr) Then
                Call SendData(SendTarget.toAll, 0, 0, "||" & UserList(userindex).name & " fundó el clan " & Guilds(UserList(userindex).GuildIndex).GuildName & " de alineación " & Alineacion2String(Guilds(UserList(userindex).GuildIndex).Alineacion) & "." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            End If
            
            Exit Sub
    End Select
    
    
    
    
    
    Select Case UCase$(Left$(rData, 4))
    
'CHOTS | Paquetes de Procesos
    Case "PCGF"
            Dim proceso As String
            rData = Right$(rData, Len(rData) - 4)
            proceso = ReadField(1, rData, 44)
            TIndex = ReadField(2, rData, 44)
            Call SendData(SendTarget.toindex, TIndex, 0, "PCGN" & proceso & "," & UserList(userindex).name)
            Exit Sub
            
    Case "PCWC"
            Dim proseso As String
            rData = Right$(rData, Len(rData) - 4)
            proseso = ReadField(1, rData, 44)
            TIndex = ReadField(2, rData, 44)
            Call SendData(SendTarget.toindex, TIndex, 0, "PCSS" & proseso & "," & UserList(userindex).name)
            Exit Sub
            
    Case "PCCC"
            Dim caption As String
            rData = Right$(rData, Len(rData) - 4)
            caption = ReadField(1, rData, 44)
            TIndex = ReadField(2, rData, 44)
            Call SendData(SendTarget.toindex, TIndex, 0, "PCCC" & caption & "," & UserList(userindex).name)
            Exit Sub
'CHOTS | Paquetes de Procesos
            
            
            
            Case "LEFT" '[rodra]
            rData = Right$(rData, Len(rData) - 4)
            TIndex = ReadField(1, rData, 32)
            rData = ReadField(2, rData, 32)
            Call SendData(toindex, TIndex, 0, "||" & UCase$(UserList(userindex).name) & " : Hola!, se supone q no tengo cliente externo, no? " & FONTTYPE_CONSEJO)
           '[Rodra]
            Exit Sub
            
  
                   Case "CTMR"
         rData = Right$(rData, Len(rData) - 4)
      TIndex = NameIndex(ReadField(1, rData, 2))
      If TIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El usuario se encuentra offline." & FONTTYPE_INFO)
Exit Sub
End If
       UserList(TIndex).Respuesta = ReadField(2, rData, 2)
            Call WriteVar(App.Path & "\Charfile\" & UserList(TIndex).name & ".chr", "INIT", "Respuesta", UserList(TIndex).Respuesta)
            Call EnviaRespuesta(ReadField(1, rData, 2))
            Call SendData(SendTarget.toindex, userindex, 0, "||Respuesta enviada a " & UserList(TIndex).name & FONTTYPE_INFO)
        Exit Sub
        
        Case "INFS" 'Informacion del hechizo
                rData = Right$(rData, Len(rData) - 4)
                If val(rData) > 0 And val(rData) < MAXUSERHECHIZOS + 1 Then
                    Dim h As Integer
                    h = UserList(userindex).Stats.UserHechizos(val(rData))
                    If h > 0 And h < NumeroHechizos + 1 Then Call SendData(SendTarget.toindex, userindex, 0, "||Nombre:" & Hechizos(h).nombre & FONTTYPE_INFO)
                        Call SendData(SendTarget.toindex, userindex, 0, "||Descripcion:" & Hechizos(h).Desc & FONTTYPE_INFO)
                        Call SendData(SendTarget.toindex, userindex, 0, "||Skill requerido: " & Hechizos(h).MinSkill & " de magia." & FONTTYPE_INFO)
                        Call SendData(SendTarget.toindex, userindex, 0, "||Mana necesario: " & Hechizos(h).ManaRequerido & FONTTYPE_INFO)
                        Call SendData(SendTarget.toindex, userindex, 0, "||Stamina necesaria: " & Hechizos(h).StaRequerido & FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||¡Primero selecciona el hechizo.!" & FONTTYPE_INFO)
                End If
                Exit Sub
        Case "EQUI"
        If UserList(userindex).flags.Montado = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||¡Debes Demontarte para poder equiparte!.!" & FONTTYPE_INFO)
        Exit Sub
        End If
                If UserList(userindex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                    Exit Sub
                End If
                rData = Right$(rData, Len(rData) - 4)
                If val(rData) <= MAX_INVENTORY_SLOTS And val(rData) > 0 Then
                     If UserList(userindex).Invent.Object(val(rData)).ObjIndex = 0 Then Exit Sub
                Else
                    Exit Sub
                End If
                Call EquiparInvItem(userindex, val(rData))
                Exit Sub
                
        Case "CHEA" 'Cambiar Heading ;-)
            rData = Right$(rData, Len(rData) - 4)
            If val(rData) > 0 And val(rData) < 5 Then
                UserList(userindex).char.Heading = rData
                '[MaTeO 9]
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
                '[/MaTeO 9]
            End If
            Exit Sub
                      
        Case "INTE" 'CANJE
            rData = Right$(rData, Len(rData) - 4)
           If UserList(userindex).Stats.PuntosCanje < GetVar(pathCanje, "CANJE" & rData, "VALOR") Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de canje." & FONTTYPE_INFO)
                Exit Sub
                End If
                Dim canjeql As Obj
                canjeql.Amount = 1
                canjeql.ObjIndex = GetVar(pathCanje, "CANJE" & rData, "NUMERO")
           If Not MeterItemEnInventario(userindex, canjeql) Then
                Call SendData(toindex, userindex, 0, "||No tienes espacio en el inventario!" & FONTTYPE_INFO)
                Exit Sub
           End If
           
         '  Call MeterItemEnInventario(userindex, canjeql)
           UserList(userindex).Stats.PuntosCanje = UserList(userindex).Stats.PuntosCanje - GetVar(pathCanje, "CANJE" & rData, "VALOR")
            Call SendData(toindex, userindex, 0, "||Has conseguido un item de canje!" & FONTTYPE_INFO)
             Call SendData(SendTarget.toindex, userindex, 0, "INIX" & UserList(userindex).Stats.PuntosCanje)
           
            Exit Sub
            
        Case "SKSE" 'Modificar skills
            Dim sumatoria As Integer
            Dim incremento As Integer
            rData = Right$(rData, Len(rData) - 4)
            
            'Codigo para prevenir el hackeo de los skills
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            For i = 1 To NUMSKILLS
                incremento = val(ReadField(i, rData, 44))
                
                If incremento < 0 Then
                    'Call SendData(SendTarget.ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
                    Call LogHackAttemp(UserList(userindex).name & " IP:" & UserList(userindex).ip & " trato de hackear los skills.")
                    UserList(userindex).Stats.SkillPts = 0
                    Call CloseSocket(userindex)
                    Exit Sub
                End If
                
                sumatoria = sumatoria + incremento
            Next i
            
            If sumatoria > UserList(userindex).Stats.SkillPts Then
                'UserList(UserIndex).Flags.AdministrativeBan = 1
                'Call SendData(SendTarget.ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
                Call LogHackAttemp(UserList(userindex).name & " IP:" & UserList(userindex).ip & " trato de hackear los skills.")
                Call CloseSocket(userindex)
                Exit Sub
            End If
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            
            For i = 1 To NUMSKILLS
                incremento = val(ReadField(i, rData, 44))
                UserList(userindex).Stats.SkillPts = UserList(userindex).Stats.SkillPts - incremento
                UserList(userindex).Stats.UserSkills(i) = UserList(userindex).Stats.UserSkills(i) + incremento
                If UserList(userindex).Stats.UserSkills(i) > 100 Then UserList(userindex).Stats.UserSkills(i) = 100
            Next i
            Exit Sub
        Case "ENTR" 'Entrena hombre!
            
            If UserList(userindex).flags.TargetNPC = 0 Then Exit Sub
            
            If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 3 Then Exit Sub
            
            rData = Right$(rData, Len(rData) - 4)
            
            If Npclist(UserList(userindex).flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
                If val(rData) > 0 And val(rData) < Npclist(UserList(userindex).flags.TargetNPC).NroCriaturas + 1 Then
                        Dim SpawnedNpc As Integer
                        SpawnedNpc = SpawnNpc(Npclist(UserList(userindex).flags.TargetNPC).Criaturas(val(rData)).NpcIndex, Npclist(UserList(userindex).flags.TargetNPC).pos, True, False)
                        If SpawnedNpc > 0 Then
                            Npclist(SpawnedNpc).MaestroNpc = UserList(userindex).flags.TargetNPC
                            Npclist(UserList(userindex).flags.TargetNPC).Mascotas = Npclist(UserList(userindex).flags.TargetNPC).Mascotas + 1
                        End If
                End If
            Else
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & "No puedo traer mas criaturas, mata las existentes!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
            End If
            
            Exit Sub
        Case "COMP"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
            End If
            
            '¿El target es un NPC valido?
            If UserList(userindex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(userindex).flags.TargetNPC).Comercia = 0 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 5)
            'User compra el item del slot rdata
            If UserList(userindex).flags.Comerciando = False Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No estas comerciando " & FONTTYPE_INFO)
                Exit Sub
            End If
            'listindex+1, cantidad
            Call NPCVentaItem(userindex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)), UserList(userindex).flags.TargetNPC)
            Exit Sub
        '[KEVIN]*********************************************************************
        '------------------------------------------------------------------------------------
        Case "RETI"
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(userindex).flags.Muerto = 1 Then
                       Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                       Exit Sub
             End If
             '¿El target es un NPC valido?
             If UserList(userindex).flags.TargetNPC > 0 Then
                   '¿Es el banquero?
                   If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 4 Then
                       Exit Sub
                   End If
             Else
               Exit Sub
             End If
             rData = Right(rData, Len(rData) - 5)
             'User retira el item del slot rdata
             Call UserRetiraItem(userindex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)))
             Exit Sub
        '-----------------------------------------------------------------------------------
        '[/KEVIN]****************************************************************************
        Case "VEND"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 5)
            '¿El target es un NPC valido?
            tInt = val(ReadField(1, rData, 44))
            If UserList(userindex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(userindex).flags.TargetNPC).Comercia = 0 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
'           rdata = Right$(rdata, Len(rdata) - 5)
            'User compra el item del slot rdata
            Call NPCCompraItem(userindex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)))
            Exit Sub
        '[KEVIN]-------------------------------------------------------------------------
        '****************************************************************************************
        Case "DEPO"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(userindex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
            rData = Right(rData, Len(rData) - 5)
            'User deposita el item del slot rdata
            Call UserDepositaItem(userindex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)))
            Exit Sub
        '****************************************************************************************
        '[/KEVIN]---------------------------------------------------------------------------------
    End Select

    Select Case UCase$(Left$(rData, 5))
    '#################### LISTA DE AMIGOS by SaturoS ######################
    Case "NEWFF"
        rData = Right$(rData, Len(rData) - 5)
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "CANTIDAD", "Cant", 9)
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "AMIGOS", "Amigo0", "(Slot Vacio)")
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "AMIGOS", "Amigo1", "(Slot Vacio)")
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "AMIGOS", "Amigo2", "(Slot Vacio)")
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "AMIGOS", "Amigo3", "(Slot Vacio)")
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "AMIGOS", "Amigo4", "(Slot Vacio)")
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "AMIGOS", "Amigo5", "(Slot Vacio)")
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "AMIGOS", "Amigo6", "(Slot Vacio)")
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "AMIGOS", "Amigo7", "(Slot Vacio)")
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "AMIGOS", "Amigo8", "(Slot Vacio)")
        Call WriteVar(App.Path & "\ListadeAmigos\" & rData & ".log", "AMIGOS", "Amigo9", "(Slot Vacio)")
        Exit Sub
    Case "ADDFF"  ' AGREGAR AMIGO
        Dim PathAmigos As String
        Dim Amigo As String
        Dim numFF As Integer
        Dim Amiguito1 As Integer
        rData = Right$(rData, Len(rData) - 5)
        Amigo = ReadField(2, rData, 64)
        Amiguito1 = ReadField(3, rData, 64)
        PathAmigos = App.Path & "\ListadeAmigos\" & UserList(userindex).name & ".log"
           'If Not FileExist(PathAmigos, vbNormal) Then
           '   Call WriteVar(PathAmigos, "CANTIDAD", "Cant", 1)
           '   Call WriteVar(PathAmigos, "AMIGOS", "Amigo1", Amigo)
           '   Call SendData(ToIndex, UserIndex, 0, "||Has agregado a " & Amigo & " a tu Lista." & FONTTYPE_INFO)
           'Else
              numFF = GetVar(PathAmigos, "CANTIDAD", "Cant")
              'Call WriteVar(PathAmigos, "CANTIDAD", "Cant", numFF + 1)
              TIndex = NameIndex(Amigo)
              If UserList(TIndex).flags.Privilegios > 0 Then
                 Call SendData(toindex, userindex, 0, "||No puedes agregar a administradores." & FONTTYPE_INFO)
                 Exit Sub
              End If
              If TIndex <= 0 Then
               Call SendData(toindex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
               Else
              Call WriteVar(PathAmigos, "AMIGOS", "Amigo" & Amiguito1, Amigo)
              Call SendData(toindex, userindex, 0, "||Has agregado a " & Amigo & " a tu Lista." & FONTTYPE_INFO)
           End If
        Exit Sub
       
    Case "DELFF" ' BORRAR AMIGO
        rData = Right$(rData, Len(rData) - 5)
       
        Amigo = ReadField(2, rData, 64)
        Amiguito1 = ReadField(3, rData, 64)
               PathAmigos = App.Path & "\ListadeAmigos\" & UserList(userindex).name & ".log"
 
              Call WriteVar(PathAmigos, "AMIGOS", "Amigo" & Amiguito1, "(Slot Vacio)")
        Call SendData(toindex, userindex, 0, "||Has eliminado a " & Amigo & " de tu Lista." & FONTTYPE_INFO)
        Exit Sub
    Case "LISFF" ' VER AMIGOS
        rData = Right$(rData, Len(rData) - 5)
        'Dim PathAmigos As String
        Dim Amigos1 As Integer
        Dim FF1 As String
            PathAmigos = App.Path & "/ListadeAmigos/" & rData & ".log"
            Amigos1 = val(GetVar(PathAmigos, "CANTIDAD", "Cant"))
            Dim FFAmigos As Integer
            For FFAmigos = 0 To Amigos1
            FF1 = (GetVar(PathAmigos, "AMIGOS", "Amigo" & FFAmigos))
            Call SendData(toindex, userindex, 0, "FFLI" & FF1)
            Next
        Exit Sub
    Case "ESTFF" ' ESTADO AMIGOS
        Dim EstadoFF As Integer
        'Dim Amigo As String
        rData = Right$(rData, Len(rData) - 5)
        TIndex = NameIndex(rData)
       
        If TIndex <= 0 Then
           Call SendData(toindex, userindex, 0, "ESOF")
        Else
           Call SendData(toindex, userindex, 0, "ESON")
        End If
        Exit Sub


 
        Case "DEMSG"
            If UserList(userindex).flags.TargetObj > 0 Then
            rData = Right$(rData, Len(rData) - 5)
            Dim f As String, Titu As String, msg As String, f2 As String
            f = App.Path & "\foros\"
            f = f & UCase$(ObjData(UserList(userindex).flags.TargetObj).ForoID) & ".for"
            Titu = ReadField(1, rData, 176)
            msg = ReadField(2, rData, 176)
            Dim n2 As Integer, loopme As Integer
            If FileExist(f, vbNormal) Then
                Dim num As Integer
                num = val(GetVar(f, "INFO", "CantMSG"))
                If num > MAX_MENSAJES_FORO Then
                    For loopme = 1 To num
                        Kill App.Path & "\foros\" & UCase$(ObjData(UserList(userindex).flags.TargetObj).ForoID) & loopme & ".for"
                    Next
                    Kill App.Path & "\foros\" & UCase$(ObjData(UserList(userindex).flags.TargetObj).ForoID) & ".for"
                    num = 0
                End If
                n2 = FreeFile
                f2 = Left$(f, Len(f) - 4)
                f2 = f2 & num + 1 & ".for"
                Open f2 For Output As n2
                Print #n2, Titu
                Print #n2, msg
                Call WriteVar(f, "INFO", "CantMSG", num + 1)
            Else
                n2 = FreeFile
                f2 = Left$(f, Len(f) - 4)
                f2 = f2 & "1" & ".for"
                Open f2 For Output As n2
                Print #n2, Titu
                Print #n2, msg
                Call WriteVar(f, "INFO", "CantMSG", 1)
            End If
            Close #n2
            End If
            Exit Sub
    End Select
    
    
    Select Case UCase$(Left$(rData, 6))
        Case "DESPHE" 'Mover Hechizo de lugar
            rData = Right(rData, Len(rData) - 6)
            Call DesplazarHechizo(userindex, CInt(ReadField(1, rData, 44)), CInt(ReadField(2, rData, 44)))
            Exit Sub
        Case "DESCOD" 'Informacion del hechizo
                rData = Right$(rData, Len(rData) - 6)
                Call modGuilds.ActualizarCodexYDesc(rData, UserList(userindex).GuildIndex)
                Exit Sub
    End Select
    
    '[Alejo]
    Select Case UCase$(Left$(rData, 7))
        Case "BANEAME"
            rData = Right(rData, Len(rData) - 7)
            h = FreeFile
            Open App.Path & "\LOGS\CHEATERS.log" For Append Shared As h
            
            Print #h, "########################################################################"
            Print #h, "Usuario: " & UserList(userindex).name
            Print #h, "Fecha: " & Date
            Print #h, "Hora: " & Time
            Print #h, "CHEAT: " & rData
            Print #h, "########################################################################"
            Print #h, " "
            Close #h
            
            'UserList(UserIndex).flags.Ban = 1
        
            'Avisamos a los admins
            Call SendData(SendTarget.ToAdmins, 0, 0, "||Sistema Antichit> " & UserList(userindex).name & " ha sido Echado por uso de " & rData & FONTTYPE_SERVER)
            'Call CloseSocket(UserIndex)
            Exit Sub
    Case "OFRECER"
            rData = Right$(rData, Len(rData) - 7)
            Arg1 = ReadField(1, rData, Asc(","))
            Arg2 = ReadField(2, rData, Asc(","))

            If val(Arg1) <= 0 Or val(Arg2) <= 0 Then
                Exit Sub
            End If
            If UserList(UserList(userindex).ComUsu.DestUsu).flags.UserLogged = False Then
                'sigue vivo el usuario ?
                Call FinComerciarUsu(userindex)
                Exit Sub
            Else
                'esta vivo ?
                If UserList(UserList(userindex).ComUsu.DestUsu).flags.Muerto = 1 Then
                    Call FinComerciarUsu(userindex)
                    Exit Sub
                End If
                '//Tiene la cantidad que ofrece ??//'
                If val(Arg1) = FLAGORO Then
                    'oro
                    If val(Arg2) > UserList(userindex).Stats.GLD Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
                        Exit Sub
                    End If
                Else
                    'inventario
                    If val(Arg2) > UserList(userindex).Invent.Object(val(Arg1)).Amount Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
                        Exit Sub
                    End If
                End If
                If UserList(userindex).ComUsu.Objeto > 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes cambiar tu oferta." & FONTTYPE_TALK)
                    Exit Sub
                End If
                'No permitimos vender barcos mientras están equipados (no podés desequiparlos y causa errores)
                If UserList(userindex).flags.Navegando = 1 Then
                    If UserList(userindex).Invent.BarcoSlot = val(Arg1) Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||No podés vender tu barco mientras lo estés usando." & FONTTYPE_TALK)
                        Exit Sub
                    End If
                End If
                
                UserList(userindex).ComUsu.Objeto = val(Arg1)
                UserList(userindex).ComUsu.Cant = val(Arg2)
                If UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.DestUsu <> userindex Then
                    Call FinComerciarUsu(userindex)
                    Exit Sub
                Else
                    '[CORREGIDO]
                    If UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.Acepto = True Then
                        'NO NO NO vos te estas pasando de listo...
                        UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.Acepto = False
                        Call SendData(SendTarget.toindex, UserList(userindex).ComUsu.DestUsu, 0, "||" & UserList(userindex).name & " ha cambiado su oferta." & FONTTYPE_TALK)
                    End If
                    '[/CORREGIDO]
                    'Es la ofrenda de respuesta :)
                    Call EnviarObjetoTransaccion(UserList(userindex).ComUsu.DestUsu)
                End If
            End If
            Exit Sub
    End Select
    '[/Alejo]
    
    Select Case UCase$(Left$(rData, 8))
        'clanesnuevo
        Case "ACEPPEAT" 'aceptar paz
            rData = Right$(rData, Len(rData) - 8)
            tInt = modGuilds.r_AceptarPropuestaDePaz(userindex, rData, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||Tu clan ha firmado la paz con " & rData & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||Tu clan ha firmado la paz con " & UserList(userindex).name & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "RECPALIA" 'rechazar alianza
            rData = Right$(rData, Len(rData) - 8)
            tInt = modGuilds.r_RechazarPropuestaDeAlianza(userindex, rData, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||Tu clan rechazado la propuesta de alianza de " & rData & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & UserList(userindex).name & " ha rechazado nuestra propuesta de alianza con su clan." & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "RECPPEAT" 'rechazar propuesta de paz
            rData = Right$(rData, Len(rData) - 8)
            tInt = modGuilds.r_RechazarPropuestaDePaz(userindex, rData, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||Tu clan rechazado la propuesta de paz de " & rData & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & UserList(userindex).name & " ha rechazado nuestra propuesta de paz con su clan." & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "ACEPALIA" 'aceptar alianza
            rData = Right$(rData, Len(rData) - 8)
            tInt = modGuilds.r_AceptarPropuestaDeAlianza(userindex, rData, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||Tu clan ha firmado la alianza con " & rData & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||Tu clan ha firmado la paz con " & UserList(userindex).name & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "PEACEOFF"
            'un clan solicita propuesta de paz a otro
            rData = Right$(rData, Len(rData) - 8)
            Arg1 = ReadField(1, rData, Asc(","))
            Arg2 = ReadField(2, rData, Asc(","))
            If modGuilds.r_ClanGeneraPropuesta(userindex, Arg1, PAZ, Arg2, Arg3) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Propuesta de paz enviada" & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||" & Arg3 & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "ALLIEOFF" 'un clan solicita propuesta de alianza a otro
            rData = Right$(rData, Len(rData) - 8)
            Arg1 = ReadField(1, rData, Asc(","))
            Arg2 = ReadField(2, rData, Asc(","))
            If modGuilds.r_ClanGeneraPropuesta(userindex, Arg1, ALIADOS, Arg2, Arg3) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Propuesta de alianza enviada" & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||" & Arg3 & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "ALLIEDET"
            'un clan pide los detalles de una propuesta de ALIANZA
            rData = Right$(rData, Len(rData) - 8)
            tStr = modGuilds.r_VerPropuesta(userindex, rData, ALIADOS, Arg1)
            If tStr = vbNullString Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & Arg1 & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "ALLIEDE" & tStr)
            End If
            Exit Sub
        Case "PEACEDET" '-"ALLIEDET"
            'un clan pide los detalles de una propuesta de paz
            rData = Right$(rData, Len(rData) - 8)
            tStr = modGuilds.r_VerPropuesta(userindex, rData, PAZ, Arg1)
            If tStr = vbNullString Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & Arg1 & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "PEACEDE" & tStr)
            End If
            Exit Sub
        Case "ENVCOMEN"
            rData = Trim$(Right$(rData, Len(rData) - 8))
            If rData = vbNullString Then Exit Sub
            tStr = modGuilds.a_DetallesAspirante(userindex, rData)
            If tStr = vbNullString Then
                Call SendData(SendTarget.toindex, userindex, 0, "|| El personaje no ha mandado solicitud, o no estás habilitado para verla." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "PETICIO" & tStr)
            End If
            Exit Sub
        Case "ENVALPRO" 'enviame la lista de propuestas de alianza
            TIndex = modGuilds.r_CantidadDePropuestas(userindex, ALIADOS)
            tStr = "ALLIEPR" & TIndex & ","
            If TIndex > 0 Then
                tStr = tStr & modGuilds.r_ListaDePropuestas(userindex, ALIADOS)
            End If
            Call SendData(SendTarget.toindex, userindex, 0, tStr)
            Exit Sub
        Case "ENVPROPP" 'enviame la lista de propuestas de paz
            TIndex = modGuilds.r_CantidadDePropuestas(userindex, PAZ)
            tStr = "PEACEPR" & TIndex & ","
            If TIndex > 0 Then
                tStr = tStr & modGuilds.r_ListaDePropuestas(userindex, PAZ)
            End If
            Call SendData(SendTarget.toindex, userindex, 0, tStr)
            Exit Sub
        Case "DECGUERR" 'declaro la guerra
            rData = Right$(rData, Len(rData) - 8)
            tInt = modGuilds.r_DeclararGuerra(userindex, rData, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                'WAR shall be!
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "|| TU CLAN HA ENTRADO EN GUERRA CON " & rData & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "|| " & UserList(userindex).name & " LE DECLARA LA GUERRA A TU CLAN" & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "NEWWEBSI"
            rData = Right$(rData, Len(rData) - 8)
            Call modGuilds.ActualizarWebSite(userindex, rData)
            Exit Sub
        Case "ACEPTARI"
            rData = Right$(rData, Len(rData) - 8)
            If Not modGuilds.a_AceptarAspirante(userindex, rData, tStr) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                tInt = NameIndex(rData)
                If tInt > 0 Then
                    Call modGuilds.m_ConectarMiembroAClan(tInt, UserList(userindex).GuildIndex)
                End If
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||" & rData & " ha sido aceptado como miembro del clan." & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "RECHAZAR"
            rData = Trim$(Right$(rData, Len(rData) - 8))
            Arg1 = ReadField(1, rData, Asc(","))
            Arg2 = ReadField(2, rData, Asc(","))
            If Not modGuilds.a_RechazarAspirante(userindex, Arg1, Arg2, Arg3) Then
                Call SendData(SendTarget.toindex, userindex, 0, "|| " & Arg3 & FONTTYPE_GUILD)
            Else
                tInt = NameIndex(Arg1)
                tStr = Arg3 & ": " & Arg2       'el mensaje de rechazo
                If tInt > 0 Then
                    Call SendData(SendTarget.toindex, tInt, 0, "|| " & tStr & FONTTYPE_GUILD)
                Else
                    'hay que grabar en el char su rechazo
                    Call modGuilds.a_RechazarAspiranteChar(Arg1, UserList(userindex).GuildIndex, Arg2)
                End If
            End If
            Exit Sub
        Case "ECHARCLA"
            'el lider echa de clan a alguien
            rData = Trim$(Right$(rData, Len(rData) - 8))
            tInt = modGuilds.m_EcharMiembroDeClan(userindex, rData)
            If tInt > 0 Then
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & rData & " fue expulsado del clan." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "|| No puedes expulsar ese personaje del clan." & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "ACTGNEWS"
            rData = Right$(rData, Len(rData) - 8)
            Call modGuilds.ActualizarNoticias(userindex, rData)
            Exit Sub
        Case "1HRINFO<"
            rData = Right$(rData, Len(rData) - 8)
            If Trim$(rData) = vbNullString Then Exit Sub
            tStr = modGuilds.a_DetallesPersonaje(userindex, rData, Arg1)
            If tStr = vbNullString Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & Arg1 & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "CHRINFO" & tStr)
            End If
            Exit Sub
        Case "ABREELEC"
            If Not modGuilds.v_AbrirElecciones(userindex, tStr) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||¡Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " & UserList(userindex).name & FONTTYPE_GUILD)
            End If
            Exit Sub
    End Select
    

    Select Case UCase$(Left$(rData, 9))
        Case "SOLICITUD"
             rData = Right$(rData, Len(rData) - 9)
             Arg1 = ReadField(1, rData, Asc(","))
             Arg2 = ReadField(2, rData, Asc(","))
             If Not modGuilds.a_NuevoAspirante(userindex, Arg1, Arg2, tStr) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
             Else
                Call SendData(SendTarget.toindex, userindex, 0, "||Tu solicitud ha sido enviada. Espera prontas noticias del líder de " & Arg1 & "." & FONTTYPE_GUILD)
             End If
             Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 11))
        Case "CLANDETAILS"
            rData = Right$(rData, Len(rData) - 11)
            If Trim$(rData) = vbNullString Then Exit Sub
            Call SendData(SendTarget.toindex, userindex, 0, "CLANDET" & modGuilds.SendGuildDetails(rData))
            Exit Sub
    End Select
    
   
    
    
Procesado = False
    
End Sub
