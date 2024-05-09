Attribute VB_Name = "Acciones"
        

'Pablo Ignacio Márquez

Option Explicit

''
' Modulo para manejar las acciones (doble click) de los carteles, foro, puerta, ramitas
'

''eva
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Ys

Sub Accion(ByVal userindex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer)
On Error Resume Next

'¿Posicion valida?
If InMapBounds(Map, x, Y) Then
   
    Dim FoundChar As Byte
    Dim FoundSomething As Byte
    Dim TempCharIndex As Integer
       
    '¿Es un obj?
    If MapData(Map, x, Y).OBJInfo.ObjIndex > 0 Then
        UserList(userindex).flags.TargetObj = MapData(Map, x, Y).OBJInfo.ObjIndex
        
        Select Case ObjData(MapData(Map, x, Y).OBJInfo.ObjIndex).OBJType
            
            Case eOBJType.otPuertas 'Es una puerta
                Call AccionParaPuerta(Map, x, Y, userindex)
            Case eOBJType.otCARTELES 'Es un cartel
                Call AccionParaCartel(Map, x, Y, userindex)
            Case eOBJType.otFOROS 'Foro
                Call AccionParaForo(Map, x, Y, userindex)
            Case eOBJType.otLeña    'Leña
                If MapData(Map, x, Y).OBJInfo.ObjIndex = FOGATA_APAG And UserList(userindex).flags.Muerto = 0 Then
                    Call AccionParaRamita(Map, x, Y, userindex)
                End If
        End Select
    '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
    ElseIf MapData(Map, x + 1, Y).OBJInfo.ObjIndex > 0 Then
        UserList(userindex).flags.TargetObj = MapData(Map, x + 1, Y).OBJInfo.ObjIndex
        Call SendData(SendTarget.toIndex, userindex, 0, "SELE" & ObjData(MapData(Map, x + 1, Y).OBJInfo.ObjIndex).OBJType & "," & ObjData(MapData(Map, x + 1, Y).OBJInfo.ObjIndex).name & "," & "OBJ")
        Select Case ObjData(MapData(Map, x + 1, Y).OBJInfo.ObjIndex).OBJType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, x + 1, Y, userindex)
            
        End Select
    ElseIf MapData(Map, x + 1, Y + 1).OBJInfo.ObjIndex > 0 Then
        UserList(userindex).flags.TargetObj = MapData(Map, x + 1, Y + 1).OBJInfo.ObjIndex
        Call SendData(SendTarget.toIndex, userindex, 0, "SELE" & ObjData(MapData(Map, x + 1, Y + 1).OBJInfo.ObjIndex).OBJType & "," & ObjData(MapData(Map, x + 1, Y + 1).OBJInfo.ObjIndex).name & "," & "OBJ")
        Select Case ObjData(MapData(Map, x + 1, Y + 1).OBJInfo.ObjIndex).OBJType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, x + 1, Y + 1, userindex)
            
        End Select
    ElseIf MapData(Map, x, Y + 1).OBJInfo.ObjIndex > 0 Then
        UserList(userindex).flags.TargetObj = MapData(Map, x, Y + 1).OBJInfo.ObjIndex
        Call SendData(SendTarget.toIndex, userindex, 0, "SELE" & ObjData(MapData(Map, x, Y + 1).OBJInfo.ObjIndex).OBJType & "," & ObjData(MapData(Map, x, Y + 1).OBJInfo.ObjIndex).name & "," & "OBJ")
        Select Case ObjData(MapData(Map, x, Y + 1).OBJInfo.ObjIndex).OBJType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, x, Y + 1, userindex)
            
        End Select
    
    ElseIf MapData(Map, x, Y).userindex > 0 Then
    
    ElseIf MapData(Map, x, Y).NpcIndex > 0 Then     'Acciones NPCs
        'Set the target NPC
        UserList(userindex).flags.TargetNPC = MapData(Map, x, Y).NpcIndex
        
        If Npclist(MapData(Map, x, Y).NpcIndex).Comercia = 1 Then
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 3 Then
                Call SendData(SendTarget.toIndex, userindex, 0, "Z27")
                Exit Sub
            End If
            If UserList(userindex).flags.Montado = True Then
            Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes comerciar estando arriba de tu Mascota!" & FONTTYPE_INFO)
                Exit Sub
            End If
            'Iniciamos la rutina pa' comerciar.
            Call IniciarCOmercioNPC(userindex)
        
        ElseIf Npclist(MapData(Map, x, Y).NpcIndex).NPCtype = eNPCType.Banquero Then
            If Distancia(Npclist(MapData(Map, x, Y).NpcIndex).pos, UserList(userindex).pos) > 4 Then
                Call SendData(SendTarget.toIndex, userindex, 0, "Z27")
                Exit Sub
            End If
              If UserList(userindex).flags.Montado = True Then
            Call SendData(SendTarget.toIndex, userindex, 0, "||No puedes usar la boveda estando arriba de tu Mascota!" & FONTTYPE_INFO)
                Exit Sub
            End If
            'A depositar de una
            Call IniciarDeposito(userindex)
        
        ElseIf Npclist(MapData(Map, x, Y).NpcIndex).NPCtype = eNPCType.Revividor Then
            If Distancia(UserList(userindex).pos, Npclist(MapData(Map, x, Y).NpcIndex).pos) > 10 Then
                Call SendData(SendTarget.toIndex, userindex, 0, "Z32")
                Exit Sub
            End If
           If UserList(userindex).flags.Envenenado = 1 Then
           UserList(userindex).flags.Envenenado = 0
            Call SendData(SendTarget.toIndex, userindex, 0, "||Te has curado del envenenamiento." & FONTTYPE_INFO)
          End If
           'Revivimos si es necesario
            If UserList(userindex).flags.Muerto = 1 Then
                Call RevivirUsuario(userindex)
            End If
            
            'curamos totalmente
            UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
            
            Call EnviarHP(userindex)
        End If
    Else
        UserList(userindex).flags.TargetNPC = 0
        UserList(userindex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(userindex).flags.TargetUser = 0
        UserList(userindex).flags.TargetObj = 0
    End If
End If

End Sub

Sub AccionParaForo(ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal userindex As Integer)
On Error Resume Next

Dim pos As WorldPos
pos.Map = Map
pos.x = x
pos.Y = Y

If Distancia(pos, UserList(userindex).pos) > 2 Then
    Call SendData(SendTarget.toIndex, userindex, 0, "Z27")
    Exit Sub
End If

'¿Hay mensajes?
Dim f As String, tit As String, men As String, base As String, auxcad As String
f = App.Path & "\foros\" & UCase$(ObjData(MapData(Map, x, Y).OBJInfo.ObjIndex).ForoID) & ".for"
If FileExist(f, vbNormal) Then
    Dim num As Integer
    num = val(GetVar(f, "INFO", "CantMSG"))
    base = Left$(f, Len(f) - 4)
    Dim i As Integer
    Dim n As Integer
    For i = 1 To num
        n = FreeFile
        f = base & i & ".for"
        Open f For Input Shared As #n
        Input #n, tit
        men = ""
        auxcad = ""
        Do While Not EOF(n)
            Input #n, auxcad
            men = men & vbCrLf & auxcad
        Loop
        Close #n
        Call SendData(SendTarget.toIndex, userindex, 0, "FMSG" & tit & Chr(176) & men)
        
    Next
End If
Call SendData(SendTarget.toIndex, userindex, 0, "MFOR")
End Sub


Sub AccionParaPuerta(ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal userindex As Integer)
On Error Resume Next

Dim MiObj As Obj
Dim wp As WorldPos

If Not (Distance(UserList(userindex).pos.x, UserList(userindex).pos.Y, x, Y) > 2) Then
    If ObjData(MapData(Map, x, Y).OBJInfo.ObjIndex).Llave = 0 Then
        If ObjData(MapData(Map, x, Y).OBJInfo.ObjIndex).Cerrada = 1 Then
                'Abre la puerta
                If ObjData(MapData(Map, x, Y).OBJInfo.ObjIndex).Llave = 0 Then
                    
                    MapData(Map, x, Y).OBJInfo.ObjIndex = ObjData(MapData(Map, x, Y).OBJInfo.ObjIndex).IndexAbierta
                    
                    Call ModAreas.SendToAreaByPos(Map, x, Y, "HO" & ObjData(MapData(Map, x, Y).OBJInfo.ObjIndex).GrhIndex & "," & x & "," & Y)
                     
                    'Desbloquea
                    MapData(Map, x, Y).Blocked = 0
                    MapData(Map, x - 1, Y).Blocked = 0
                    
                    'Bloquea todos los mapas
                    Call Bloquear(SendTarget.ToMap, 0, Map, Map, x, Y, 0)
                    Call Bloquear(SendTarget.ToMap, 0, Map, Map, x - 1, Y, 0)
                    
                      
                    'Sonido
                    SendData SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_PUERTA
                    
                Else
                     Call SendData(SendTarget.toIndex, userindex, 0, "||La puerta esta cerrada con llave." & FONTTYPE_INFO)
                End If
        Else
                'Cierra puerta
                MapData(Map, x, Y).OBJInfo.ObjIndex = ObjData(MapData(Map, x, Y).OBJInfo.ObjIndex).IndexCerrada
                
                Call ModAreas.SendToAreaByPos(Map, x, Y, "HO" & ObjData(MapData(Map, x, Y).OBJInfo.ObjIndex).GrhIndex & "," & x & "," & Y)
                
                
                MapData(Map, x, Y).Blocked = 1
                MapData(Map, x - 1, Y).Blocked = 1
                
                
                Call Bloquear(SendTarget.ToMap, 0, Map, Map, x - 1, Y, 1)
                Call Bloquear(SendTarget.ToMap, 0, Map, Map, x, Y, 1)
                
                SendData SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_PUERTA
        End If
        
        UserList(userindex).flags.TargetObj = MapData(Map, x, Y).OBJInfo.ObjIndex
    Else
        Call SendData(SendTarget.toIndex, userindex, 0, "||La puerta esta cerrada con llave." & FONTTYPE_INFO)
    End If
Else
    Call SendData(SendTarget.toIndex, userindex, 0, "Z27")
End If

End Sub

Sub AccionParaCartel(ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal userindex As Integer)
On Error Resume Next


Dim MiObj As Obj

If ObjData(MapData(Map, x, Y).OBJInfo.ObjIndex).OBJType = 8 Then
  
  If Len(ObjData(MapData(Map, x, Y).OBJInfo.ObjIndex).texto) > 0 Then
       Call SendData(SendTarget.toIndex, userindex, 0, "MCAR" & _
        ObjData(MapData(Map, x, Y).OBJInfo.ObjIndex).texto & _
        Chr(176) & ObjData(MapData(Map, x, Y).OBJInfo.ObjIndex).GrhSecundario)
  End If
  
End If

End Sub

Sub AccionParaRamita(ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal userindex As Integer)
On Error Resume Next

Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj
Dim raise As Integer

Dim pos As WorldPos
pos.Map = Map
pos.x = x
pos.Y = Y

If Distancia(pos, UserList(userindex).pos) > 2 Then
    Call SendData(toIndex, userindex, 0, "Z27")
    Exit Sub
End If

If MapData(Map, x, Y).trigger = eTrigger.ZONASEGURA Or MapInfo(Map).Pk = False Then
    Call SendData(SendTarget.toIndex, userindex, 0, "||En zona segura no puedes hacer fogatas." & FONTTYPE_INFO)
    Exit Sub
End If

If UserList(userindex).Stats.UserSkills(Supervivencia) > 1 And UserList(userindex).Stats.UserSkills(Supervivencia) < 6 Then
            Suerte = 3
ElseIf UserList(userindex).Stats.UserSkills(Supervivencia) >= 6 And UserList(userindex).Stats.UserSkills(Supervivencia) <= 10 Then
            Suerte = 2
ElseIf UserList(userindex).Stats.UserSkills(Supervivencia) >= 10 And UserList(userindex).Stats.UserSkills(Supervivencia) Then
            Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    If MapInfo(UserList(userindex).pos.Map).Zona <> Ciudad Then
        Obj.ObjIndex = FOGATA
        Obj.Amount = 1
        
        Call SendData(toIndex, userindex, 0, "||Has prendido la fogata." & FONTTYPE_INFO)
        Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "FO")
        
        Call MakeObj(ToMap, 0, Map, Obj, Map, x, Y)
        
        'Las fogatas prendidas se deben eliminar
        Dim Fogatita As New cGarbage
        Fogatita.Map = Map
        Fogatita.x = x
        Fogatita.Y = Y
        Call TrashCollector.Add(Fogatita)
    Else
        Call SendData(toIndex, userindex, 0, "||La ley impide realizar fogatas en las ciudades." & FONTTYPE_INFO)
        Exit Sub
    End If
Else
    Call SendData(toIndex, userindex, 0, "||No has podido hacer fuego." & FONTTYPE_INFO)
End If

'Sino tiene hambre o sed quizas suba el skill supervivencia
If UserList(userindex).flags.Hambre = 0 And UserList(userindex).flags.Sed = 0 Then
    Call SubirSkill(userindex, Supervivencia)
End If

End Sub
