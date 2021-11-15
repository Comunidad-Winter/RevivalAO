Attribute VB_Name = "Trabajo"

Option Explicit

Public Sub DoPermanecerOculto(ByVal userindex As Integer)
On Error GoTo errhandler
Dim Suerte As Integer
Dim res As Integer
If UCase$(UserList(userindex).Clase) = "GUERRERO" Then Exit Sub
If UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 10 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= -1 Then
                    Suerte = 35
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 20 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 11 Then
                    Suerte = 30
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 30 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 21 Then
                    Suerte = 28
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 40 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 31 Then
                    Suerte = 24
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 50 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 41 Then
                    Suerte = 22
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 60 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 51 Then
                    Suerte = 20
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 70 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 61 Then
                    Suerte = 18
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 80 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 71 Then
                    Suerte = 15
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 90 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 81 Then
                    Suerte = 10
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 100 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 91 Then
                    Suerte = 10     'Lo atamos con alambre.... en la 11.6 el sistema de ocultarse debería de estar bien hecho
End If

If UCase$(UserList(userindex).Clase) <> "LADRON" Then Suerte = Suerte + 50

res = RandomNumber(1, Suerte)

If res > 9 Then
    UserList(userindex).flags.Oculto = 0
    If UserList(userindex).flags.Invisible = 0 Then
        'no hace falta encriptar este (se jode el gil que bypassea esto)
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).char.CharIndex & ",0")
        Call SendData(SendTarget.toindex, userindex, 0, "Z11")
    End If
End If


Exit Sub

errhandler:
    Call LogError("Error en Sub DoPermanecerOculto")


End Sub

Public Sub DoOcultarse(ByVal userindex As Integer)

On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer
If UCase$(UserList(userindex).pos.Map) = 86 Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes ocultarte en invocaciones!" & FONTTYPE_INFO)
Exit Sub
End If
If UCase$(UserList(userindex).pos.Map) = 66 Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes ocultarte en guerras!" & FONTTYPE_INFO)
Exit Sub
End If
If UCase$(UserList(userindex).pos.Map) = 79 Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes ocultarte en torneos!" & FONTTYPE_INFO)
Exit Sub
End If
If UCase$(UserList(userindex).pos.Map) = 88 Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes ocultarte en deatmatch!!" & FONTTYPE_INFO)
Exit Sub
End If
If UCase$(UserList(userindex).pos.Map) = 78 Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes ocultarte en retos!" & FONTTYPE_INFO)
Exit Sub
End If
If UCase$(UserList(userindex).pos.Map) = 61 Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes ocultarte en duelos!" & FONTTYPE_INFO)
Exit Sub
End If
If UCase$(UserList(userindex).pos.Map) = 87 Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes ocultarte en retos!" & FONTTYPE_INFO)
Exit Sub
End If
If UCase$(UserList(userindex).Clase) = "GUERRERO" Then
UserList(userindex).flags.Oculto = 1
#If SeguridadAlkon Then
    If EncriptarProtocolosCriticos Then
        Call SendCryptedData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).char.CharIndex & ",1")
    Else
#End If
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).char.CharIndex & ",1")
#If SeguridadAlkon Then
    End If
#End If
UserList(userindex).Counters.Ocultando = UserList(userindex).Counters.Ocultando + 100

Call SendData(SendTarget.toindex, userindex, 0, "||¡Te has escondido entre las sombras!" & FONTTYPE_INFO)
Exit Sub
End If

If UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 10 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= -1 Then
                    Suerte = 35
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 20 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 11 Then
                    Suerte = 30
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 30 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 21 Then
                    Suerte = 28
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 40 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 31 Then
                    Suerte = 24
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 50 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 41 Then
                    Suerte = 22
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 60 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 51 Then
                    Suerte = 20
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 70 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 61 Then
                    Suerte = 18
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 80 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 71 Then
                    Suerte = 15
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 90 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 81 Then
                    Suerte = 10
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 100 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 91 Then
                    Suerte = 7
End If

If UCase$(UserList(userindex).Clase) <> "LADRON" Then Suerte = Suerte + 50

res = RandomNumber(1, Suerte)

If res <= 5 Then
    UserList(userindex).flags.Oculto = 1
#If SeguridadAlkon Then
    If EncriptarProtocolosCriticos Then
        Call SendCryptedData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).char.CharIndex & ",1")
    Else
#End If
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).char.CharIndex & ",1")
#If SeguridadAlkon Then
    End If
#End If
    Call SendData(SendTarget.toindex, userindex, 0, "||¡Te has escondido entre las sombras!" & FONTTYPE_INFO)
    Call SubirSkill(userindex, Ocultarse)
Else
    '[CDT 17-02-2004]
    If Not UserList(userindex).flags.UltimoMensaje = 4 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||¡No has logrado esconderte!" & FONTTYPE_INFO)
        UserList(userindex).flags.UltimoMensaje = 4
    End If
    '[/CDT]
End If
If UCase$(UserList(userindex).Clase) = "GUERRERO" Then
UserList(userindex).Counters.Ocultando = UserList(userindex).Counters.Ocultando + 100
Else
UserList(userindex).Counters.Ocultando = UserList(userindex).Counters.Ocultando + 1
End If
Exit Sub

errhandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub


Public Sub DoNavega(ByVal userindex As Integer, ByRef Barco As ObjData, ByVal Slot As Integer)

Dim ModNave As Long
ModNave = ModNavegacion(UserList(userindex).Clase)

If UserList(userindex).Stats.UserSkills(eSkill.Navegacion) / ModNave < Barco.MinSkill Then
    Call SendData(SendTarget.toindex, userindex, 0, "||No tenes suficientes conocimientos para usar este barco." & FONTTYPE_INFO)
    Call SendData(SendTarget.toindex, userindex, 0, "||Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion." & FONTTYPE_INFO)
    Exit Sub
End If

UserList(userindex).Invent.BarcoObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
UserList(userindex).Invent.BarcoSlot = Slot

If UserList(userindex).flags.Navegando = 0 Then
    
    UserList(userindex).char.Head = 0
    
    If UserList(userindex).flags.Muerto = 0 Then
        UserList(userindex).char.Body = Barco.Ropaje
    Else
        UserList(userindex).char.Body = iFragataFantasmal
    End If
    
    UserList(userindex).char.ShieldAnim = NingunEscudo
    UserList(userindex).char.WeaponAnim = NingunArma
    UserList(userindex).char.CascoAnim = NingunCasco
     '[MaTeO 9]
    UserList(userindex).char.Alas = NingunAlas
    '[/MaTeO 9]
    UserList(userindex).flags.Navegando = 1
    
Else
    
    UserList(userindex).flags.Navegando = 0
    
    If UserList(userindex).flags.Muerto = 0 Then
        UserList(userindex).char.Head = UserList(userindex).OrigChar.Head
        
        If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(userindex).char.Body = ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(userindex)
        End If
        
        If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then _
            UserList(userindex).char.ShieldAnim = ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then _
            UserList(userindex).char.WeaponAnim = ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then _
            UserList(userindex).char.CascoAnim = ObjData(UserList(userindex).Invent.CascoEqpObjIndex).CascoAnim
    Else
        UserList(userindex).char.Body = iCuerpoMuerto
        UserList(userindex).char.Head = iCabezaMuerto
        UserList(userindex).char.ShieldAnim = NingunEscudo
        UserList(userindex).char.WeaponAnim = NingunArma
        UserList(userindex).char.CascoAnim = NingunCasco
    '[MaTeO 9]
        UserList(userindex).char.Alas = NingunAlas
        '[/MaTeO 9]
    End If
End If
'[MaTeO 9]
Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
'[/MaTeO 9]
Call SendData(SendTarget.toindex, userindex, 0, "NAVEG")

End Sub
Function TieneObjetos(ByVal ItemIndex As Integer, ByVal Cant As Integer, ByVal userindex As Integer) As Boolean
'Call LogTarea("Sub TieneObjetos")

Dim i As Integer
Dim Total As Long
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(userindex).Invent.Object(i).ObjIndex = ItemIndex Then
        Total = Total + UserList(userindex).Invent.Object(i).Amount
    End If
Next i

If Cant <= Total Then
    TieneObjetos = True
    Exit Function
End If
        
End Function

Function QuitarObjetos(ByVal ItemIndex As Integer, ByVal Cant As Integer, ByVal userindex As Integer) As Boolean
'Call LogTarea("Sub QuitarObjetos")

Dim i As Integer
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(userindex).Invent.Object(i).ObjIndex = ItemIndex Then
        
        Call Desequipar(userindex, i)
        
        UserList(userindex).Invent.Object(i).Amount = UserList(userindex).Invent.Object(i).Amount - Cant
        If (UserList(userindex).Invent.Object(i).Amount <= 0) Then
            Cant = Abs(UserList(userindex).Invent.Object(i).Amount)
            UserList(userindex).Invent.Object(i).Amount = 0
            UserList(userindex).Invent.Object(i).ObjIndex = 0
        Else
            Cant = 0
        End If
        
        Call UpdateUserInv(False, userindex, i)
        
        If (Cant = 0) Then
            QuitarObjetos = True
            Exit Function
        End If
    End If
Next i

End Function



Function ModNavegacion(ByVal Clase As String) As Integer

Select Case UCase$(Clase)
    Case "PIRATA"
        ModNavegacion = 1
    Case "PESCADOR"
        ModNavegacion = 1.2
    Case Else
        ModNavegacion = 2.3
End Select

End Function

Function ModDomar(ByVal Clase As String) As Integer
    Select Case UCase$(Clase)
        Case "DRUIDA"
            ModDomar = 6
        Case "CAZADOR"
            ModDomar = 6
        Case "CLERIGO"
            ModDomar = 7
        Case Else
            ModDomar = 10
    End Select
End Function

Function CalcularPoderDomador(ByVal userindex As Integer) As Long
    With UserList(userindex).Stats
        CalcularPoderDomador = .UserAtributos(eAtributos.Carisma) _
            * (.UserSkills(eSkill.Domar) / ModDomar(UserList(userindex).Clase)) _
            + RandomNumber(1, .UserAtributos(eAtributos.Carisma) / 3) _
            + RandomNumber(1, .UserAtributos(eAtributos.Carisma) / 3) _
            + RandomNumber(1, .UserAtributos(eAtributos.Carisma) / 3)
    End With
End Function

Function FreeMascotaIndex(ByVal userindex As Integer) As Integer
    Dim j As Integer
    For j = 1 To MAXMASCOTAS
        If UserList(userindex).MascotasIndex(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function
        End If
    Next j
End Function

Sub DoDomar(ByVal userindex As Integer, ByVal NpcIndex As Integer)
'Call LogTarea("Sub DoDomar")

If UserList(userindex).NroMacotas < MAXMASCOTAS Then
    
    If Npclist(NpcIndex).MaestroUser = userindex Then
        Call SendData(SendTarget.toindex, userindex, 0, "||La criatura ya te ha aceptado como su amo." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||La criatura ya tiene amo." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Npclist(NpcIndex).flags.Domable <= CalcularPoderDomador(userindex) Then
        Dim Index As Integer
        UserList(userindex).NroMacotas = UserList(userindex).NroMacotas + 1
        Index = FreeMascotaIndex(userindex)
        UserList(userindex).MascotasIndex(Index) = NpcIndex
        UserList(userindex).MascotasType(Index) = Npclist(NpcIndex).Numero
        
        Npclist(NpcIndex).MaestroUser = userindex
        
        Call FollowAmo(NpcIndex)
        
        Call SendData(SendTarget.toindex, userindex, 0, "||La criatura te ha aceptado como su amo." & FONTTYPE_INFO)
        Call SubirSkill(userindex, Domar)
    Else
        If Not UserList(userindex).flags.UltimoMensaje = 5 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No has logrado domar la criatura." & FONTTYPE_INFO)
            UserList(userindex).flags.UltimoMensaje = 5
        End If
    End If
Else
    Call SendData(SendTarget.toindex, userindex, 0, "||No podes controlar mas criaturas." & FONTTYPE_INFO)
End If
End Sub

Sub DoAdminInvisible(ByVal userindex As Integer)
    
    If UserList(userindex).flags.AdminInvisible = 0 Then
        
        ' Sacamos el mimetizmo
        If UserList(userindex).flags.Mimetizado = 1 Then
            UserList(userindex).char.Body = UserList(userindex).CharMimetizado.Body
            UserList(userindex).char.Head = UserList(userindex).CharMimetizado.Head
            UserList(userindex).char.CascoAnim = UserList(userindex).CharMimetizado.CascoAnim
            
            UserList(userindex).char.ShieldAnim = UserList(userindex).CharMimetizado.ShieldAnim
            UserList(userindex).char.WeaponAnim = UserList(userindex).CharMimetizado.WeaponAnim
            UserList(userindex).Counters.Mimetismo = 0
            UserList(userindex).flags.Mimetizado = 0
        End If
        
        UserList(userindex).flags.AdminInvisible = 1
        UserList(userindex).flags.Invisible = 1
        UserList(userindex).flags.Oculto = 1
        UserList(userindex).flags.OldBody = UserList(userindex).char.Body
        UserList(userindex).flags.OldHead = UserList(userindex).char.Head
        UserList(userindex).char.Body = 0
        UserList(userindex).char.Head = 0
        
    Else
        
        UserList(userindex).flags.AdminInvisible = 0
        UserList(userindex).flags.Invisible = 0
        UserList(userindex).flags.Oculto = 0
        UserList(userindex).char.Body = UserList(userindex).flags.OldBody
        UserList(userindex).char.Head = UserList(userindex).flags.OldHead
        
    End If
    
    'vuelve a ser visible por la fuerza
    '[MaTeO 9]
    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
    '[/MaTeO 9]
    Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).char.CharIndex & ",0")
End Sub

Sub TratarDeHacerFogata(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal userindex As Integer)

Dim Suerte As Byte
Dim exito As Byte
Dim raise As Byte
Dim Obj As Obj
Dim posMadera As WorldPos

If Not LegalPos(Map, x, y) Then Exit Sub

With posMadera
    .Map = Map
    .x = x
    .y = y
End With

If Distancia(posMadera, UserList(userindex).pos) > 2 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Estás demasiado lejos para prender la fogata." & FONTTYPE_INFO)
    Exit Sub
End If

If UserList(userindex).flags.Muerto = 1 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes hacer fogatas estando muerto." & FONTTYPE_INFO)
    Exit Sub
End If

If MapData(Map, x, y).OBJInfo.Amount < 3 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Necesitas por lo menos tres troncos para hacer una fogata." & FONTTYPE_INFO)
    Exit Sub
End If


If UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) >= 0 And UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) < 6 Then
    Suerte = 3
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) >= 6 And UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) <= 34 Then
    Suerte = 2
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) >= 35 Then
    Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    Obj.ObjIndex = FOGATA_APAG
    Obj.Amount = MapData(Map, x, y).OBJInfo.Amount \ 3
    
    Call SendData(SendTarget.toindex, userindex, 0, "||Has hecho " & Obj.Amount & " fogatas." & FONTTYPE_INFO)
    
    Call MakeObj(SendTarget.ToMap, 0, Map, Obj, Map, x, y)
    
    'Seteamos la fogata como el nuevo TargetObj del user
    UserList(userindex).flags.TargetObj = FOGATA_APAG
Else
    '[CDT 17-02-2004]
    If Not UserList(userindex).flags.UltimoMensaje = 10 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No has podido hacer la fogata." & FONTTYPE_INFO)
        UserList(userindex).flags.UltimoMensaje = 10
    End If
    '[/CDT]
End If

Call SubirSkill(userindex, Supervivencia)


End Sub

Public Sub DoPescar(ByVal userindex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer
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
If UCase$(UserList(userindex).Clase) = "PESCADOR" Then
    Call QuitarSta(userindex, EsfuerzoPescarPescador)
Else
    Call QuitarSta(userindex, EsfuerzoPescarGeneral)
End If

If UserList(userindex).Stats.UserSkills(eSkill.Pesca) <= 10 _
   And UserList(userindex).Stats.UserSkills(eSkill.Pesca) >= -1 Then
                    Suerte = 35
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Pesca) <= 20 _
   And UserList(userindex).Stats.UserSkills(eSkill.Pesca) >= 11 Then
                    Suerte = 30
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Pesca) <= 30 _
   And UserList(userindex).Stats.UserSkills(eSkill.Pesca) >= 21 Then
                    Suerte = 28
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Pesca) <= 40 _
   And UserList(userindex).Stats.UserSkills(eSkill.Pesca) >= 31 Then
                    Suerte = 24
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Pesca) <= 50 _
   And UserList(userindex).Stats.UserSkills(eSkill.Pesca) >= 41 Then
                    Suerte = 22
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Pesca) <= 60 _
   And UserList(userindex).Stats.UserSkills(eSkill.Pesca) >= 51 Then
                    Suerte = 20
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Pesca) <= 70 _
   And UserList(userindex).Stats.UserSkills(eSkill.Pesca) >= 61 Then
                    Suerte = 18
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Pesca) <= 80 _
   And UserList(userindex).Stats.UserSkills(eSkill.Pesca) >= 71 Then
                    Suerte = 15
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Pesca) <= 90 _
   And UserList(userindex).Stats.UserSkills(eSkill.Pesca) >= 81 Then
                    Suerte = 13
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Pesca) <= 100 _
   And UserList(userindex).Stats.UserSkills(eSkill.Pesca) >= 91 Then
                    Suerte = 7
End If
res = RandomNumber(1, Suerte)

If res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
    MiObj.Amount = 3
    MiObj.ObjIndex = Pescado
    
    If Not MeterItemEnInventario(userindex, MiObj) Then
        Call TirarItemAlPiso(UserList(userindex).pos, MiObj)
    End If
    
    Call SendData(SendTarget.toindex, userindex, 0, "||¡Has pescado un lindo pez!" & FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(userindex).flags.UltimoMensaje = 6 Then
      Call SendData(SendTarget.toindex, userindex, 0, "||¡No has pescado nada!" & FONTTYPE_INFO)
      UserList(userindex).flags.UltimoMensaje = 6
    End If
    '[/CDT]
End If

Call SubirSkill(userindex, Pesca)


Exit Sub

errhandler:
    Call LogError("Error en DoPescar")
End Sub

Public Sub DoPescarRed(ByVal userindex As Integer)
On Error GoTo errhandler

Dim iSkill As Integer
Dim Suerte As Integer
Dim res As Integer
Dim EsPescador As Boolean
                  
If UCase(UserList(userindex).Clase) = "PESCADOR" Then
    Call QuitarSta(userindex, EsfuerzoPescarPescador)
    EsPescador = True
Else
    Call QuitarSta(userindex, EsfuerzoPescarGeneral)
    EsPescador = False
End If

iSkill = UserList(userindex).Stats.UserSkills(eSkill.Pesca)

' m = (60-11)/(1-10)
' y = mx - m*10 + 11

Select Case iSkill
Case 0:         Suerte = 0
Case 1 To 10:   Suerte = 60
Case 11 To 20:  Suerte = 54
Case 21 To 30:  Suerte = 49
Case 31 To 40:  Suerte = 43
Case 41 To 50:  Suerte = 38
Case 51 To 60:  Suerte = 32
Case 61 To 70:  Suerte = 27
Case 71 To 80:  Suerte = 21
Case 81 To 90:  Suerte = 16
Case 91 To 100: Suerte = 11
Case Else:      Suerte = 0
End Select

If Suerte > 0 Then
    res = RandomNumber(1, Suerte)
    
    If res < 6 Then
        Dim nPos As WorldPos
        Dim MiObj As Obj
        Dim PecesPosibles(1 To 4) As Integer
        
        PecesPosibles(1) = PESCADO1
        PecesPosibles(2) = PESCADO2
        PecesPosibles(3) = PESCADO3
        PecesPosibles(4) = PESCADO4
        
        If EsPescador = True Then
            MiObj.Amount = RandomNumber(1, 5)
        Else
            MiObj.Amount = 5
        End If
        MiObj.ObjIndex = PecesPosibles(RandomNumber(LBound(PecesPosibles), UBound(PecesPosibles)))
        
        If Not MeterItemEnInventario(userindex, MiObj) Then
            Call TirarItemAlPiso(UserList(userindex).pos, MiObj)
        End If
        
        Call SendData(SendTarget.toindex, userindex, 0, "||¡Has pescado algunos peces!" & FONTTYPE_INFO)
        
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||¡No has pescado nada!" & FONTTYPE_INFO)
    End If
    
    Call SubirSkill(userindex, Pesca)
End If

Exit Sub

errhandler:
    Call LogError("Error en DoPescarRed")
End Sub

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)

If Not MapInfo(UserList(VictimaIndex).pos.Map).Pk Then Exit Sub

If UserList(LadrOnIndex).flags.Seguro Then
    Call SendData(SendTarget.toindex, LadrOnIndex, 0, "||Debes quitar el seguro para robar" & FONTTYPE_FIGHT)
    Exit Sub
End If

If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

If UserList(VictimaIndex).Faccion.FuerzasCaos = 1 And UserList(LadrOnIndex).Faccion.FuerzasCaos = 1 Then
    Call SendData(SendTarget.toindex, LadrOnIndex, 0, "||No puedes robar a otros miembros de las fuerzas del caos" & FONTTYPE_FIGHT)
    Exit Sub
End If

If UserList(VictimaIndex).flags.Privilegios = PlayerType.User Then
    Dim Suerte As Integer
    Dim res As Integer
    
    If UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 10 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= -1 Then
                        Suerte = 35
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 20 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 11 Then
                        Suerte = 30
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 30 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 21 Then
                        Suerte = 28
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 40 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 31 Then
                        Suerte = 24
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 50 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 41 Then
                        Suerte = 22
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 60 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 51 Then
                        Suerte = 20
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 70 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 61 Then
                        Suerte = 18
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 80 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 71 Then
                        Suerte = 15
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 90 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 81 Then
                        Suerte = 10
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 100 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 91 Then
                        Suerte = 5
    End If
    res = RandomNumber(1, Suerte)
    
    If res < 3 Then 'Exito robo
       
        If (RandomNumber(1, 50) < 25) And (UCase$(UserList(LadrOnIndex).Clase) = "LADRON") Then
            If TieneObjetosRobables(VictimaIndex) Then
                Call RobarObjeto(LadrOnIndex, VictimaIndex)
            Else
                Call SendData(SendTarget.toindex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).name & " no tiene objetos." & FONTTYPE_INFO)
            End If
        Else 'Roba oro
            If UserList(VictimaIndex).Stats.GLD > 0 Then
                Dim n As Integer
                
                If UCase$(UserList(LadrOnIndex).Clase) = "LADRON" Then
                    n = RandomNumber(100, 1000)
                Else
                    n = RandomNumber(1, 100)
                End If
                If n > UserList(VictimaIndex).Stats.GLD Then n = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - n
                
                UserList(LadrOnIndex).Stats.GLD = UserList(LadrOnIndex).Stats.GLD + n
                If UserList(LadrOnIndex).Stats.GLD > MaxOro Then _
                    UserList(LadrOnIndex).Stats.GLD = MaxOro
                
                Call SendData(SendTarget.toindex, LadrOnIndex, 0, "||Le has robado " & n & " monedas de oro a " & UserList(VictimaIndex).name & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toindex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).name & " no tiene oro." & FONTTYPE_INFO)
            End If
        End If
    Else
        Call SendData(SendTarget.toindex, LadrOnIndex, 0, "||¡No has logrado robar nada!" & FONTTYPE_INFO)
        Call SendData(SendTarget.toindex, VictimaIndex, 0, "||¡" & UserList(LadrOnIndex).name & " ha intentado robarte!" & FONTTYPE_INFO)
        Call SendData(SendTarget.toindex, VictimaIndex, 0, "||¡" & UserList(LadrOnIndex).name & " es un criminal!" & FONTTYPE_INFO)
    End If

    If Not Criminal(LadrOnIndex) Then
        Call VolverCriminal(LadrOnIndex)
    End If
    
    If UserList(LadrOnIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(LadrOnIndex)

    UserList(LadrOnIndex).Reputacion.LadronesRep = UserList(LadrOnIndex).Reputacion.LadronesRep + vlLadron
    If UserList(LadrOnIndex).Reputacion.LadronesRep > MAXREP Then _
        UserList(LadrOnIndex).Reputacion.LadronesRep = MAXREP
    Call SubirSkill(LadrOnIndex, Robar)
End If


End Sub


Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Integer) As Boolean
' Agregué los barcos
' Esta funcion determina qué objetos son robables.

Dim OI As Integer

OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex

ObjEsRobable = _
ObjData(OI).OBJType <> eOBJType.otLlaves And _
UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And _
ObjData(OI).Real = 0 And _
ObjData(OI).Caos = 0 And _
ObjData(OI).OBJType <> eOBJType.otBarcos

End Function

Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'Call LogTarea("Sub RobarObjeto")
Dim flag As Boolean
Dim i As Integer
flag = False

If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
    i = 1
    Do While Not flag And i <= MAX_INVENTORY_SLOTS
        'Hay objeto en este slot?
        If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
           If ObjEsRobable(VictimaIndex, i) Then
                 If RandomNumber(1, 10) < 4 Then flag = True
           End If
        End If
        If Not flag Then i = i + 1
    Loop
Else
    i = 20
    Do While Not flag And i > 0
      'Hay objeto en este slot?
      If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
         If ObjEsRobable(VictimaIndex, i) Then
               If RandomNumber(1, 10) < 4 Then flag = True
         End If
      End If
      If Not flag Then i = i - 1
    Loop
End If

If flag Then
    Dim MiObj As Obj
    Dim num As Byte
    'Cantidad al azar
    num = RandomNumber(1, 5)
                
    If num > UserList(VictimaIndex).Invent.Object(i).Amount Then
         num = UserList(VictimaIndex).Invent.Object(i).Amount
    End If
                
    MiObj.Amount = num
    MiObj.ObjIndex = UserList(VictimaIndex).Invent.Object(i).ObjIndex
    
    UserList(VictimaIndex).Invent.Object(i).Amount = UserList(VictimaIndex).Invent.Object(i).Amount - num
                
    If UserList(VictimaIndex).Invent.Object(i).Amount <= 0 Then
          Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)
    End If
            
    Call UpdateUserInv(False, VictimaIndex, CByte(i))
                
    If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(LadrOnIndex).pos, MiObj)
    End If
    
    Call SendData(SendTarget.toindex, LadrOnIndex, 0, "||Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name & FONTTYPE_INFO)
Else
    Call SendData(SendTarget.toindex, LadrOnIndex, 0, "||No has logrado robar un objetos." & FONTTYPE_INFO)
End If

End Sub
Public Sub DoApuñalar(ByVal userindex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
'[MaTeO 6]
If UserList(userindex).Clase = "GUERRERO" Then Exit Sub
'[/MaTeO 6]
Dim Suerte As Integer
Dim res As Integer

If UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) <= 10 _
   And UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= -1 Then
                    Suerte = 200
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) <= 20 _
   And UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= 11 Then
                    Suerte = 190
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) <= 30 _
   And UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= 21 Then
                    Suerte = 180
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) <= 40 _
   And UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= 31 Then
                    Suerte = 170
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) <= 50 _
   And UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= 41 Then
                    Suerte = 160
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) <= 60 _
   And UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= 51 Then
                    Suerte = 150
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) <= 70 _
   And UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= 61 Then
                    Suerte = 140
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) <= 80 _
   And UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= 71 Then
                    Suerte = 130
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) <= 90 _
   And UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= 81 Then
                    Suerte = 120
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) < 100 _
   And UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= 91 Then
                    Suerte = 110
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) = 100 Then
                    Suerte = 100
End If

If UCase$(UserList(userindex).Clase) = "ASESINO" Then
    res = RandomNumber(0, Suerte)
    If res < 23 Then res = 0
Else
    res = RandomNumber(0, Suerte * 1.2)
End If

If res < 15 Then
    If VictimUserIndex <> 0 Then
        UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - Int(daño * 1.4)
        Call SendData(SendTarget.toindex, userindex, 0, "||Has apuñalado a " & UserList(VictimUserIndex).name & " por " & Int(daño * 1.4) & FONTTYPE_APU)
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW15")
        Call SendData(SendTarget.toindex, VictimUserIndex, 0, "||Te ha apuñalado " & UserList(userindex).name & " por " & Int(daño * 1.4) & FONTTYPE_APU)
        Call SendData(SendTarget.ToPCArea, VictimUserIndex, UserList(VictimUserIndex).pos.Map, "CFX" & UserList(VictimUserIndex).char.CharIndex & "," & 17 & "," & 1)
Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbCyan & "°Apu! + " & Int(daño * 1.4) & "!" & "°" & str(UserList(userindex).char.CharIndex))
    Else
        Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - Int(daño * 2)
        Call SendData(SendTarget.toindex, userindex, 0, "||Has apuñalado la criatura por " & Int(daño * 2) & FONTTYPE_APU)
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW13")
        Call SendData(SendTarget.ToNPCArea, VictimNpcIndex, Npclist(VictimNpcIndex).pos.Map, "CFX" & Npclist(VictimNpcIndex).char.CharIndex & "," & 17 & "," & 1)
        Call SubirSkill(userindex, Apuñalar)
         Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbCyan & "°Apu! + " & Int(daño * 2) & "!" & "°" & str(UserList(userindex).char.CharIndex))
        '[Alejo]
        Call CalcularDarExp(userindex, VictimNpcIndex, Int(daño * 2))
    End If
Else
    Call SendData(SendTarget.toindex, userindex, 0, "||¡No has logrado apuñalar a tu enemigo!" & FONTTYPE_FIGHT)
End If

End Sub

Public Sub QuitarSta(ByVal userindex As Integer, ByVal Cantidad As Integer)
UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MinSta - Cantidad
If UserList(userindex).Stats.MinSta < 0 Then UserList(userindex).Stats.MinSta = 0
End Sub

Public Sub DoTalar(ByVal userindex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer

If UCase$(UserList(userindex).Clase) = "LEÑADOR" Then
    Call QuitarSta(userindex, EsfuerzoTalarLeñador)
Else
    Call QuitarSta(userindex, EsfuerzoTalarGeneral)
End If

If UserList(userindex).Stats.UserSkills(eSkill.Talar) <= 10 _
   And UserList(userindex).Stats.UserSkills(eSkill.Talar) >= -1 Then
                    Suerte = 35
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Talar) <= 20 _
   And UserList(userindex).Stats.UserSkills(eSkill.Talar) >= 11 Then
                    Suerte = 30
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Talar) <= 30 _
   And UserList(userindex).Stats.UserSkills(eSkill.Talar) >= 21 Then
                    Suerte = 28
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Talar) <= 40 _
   And UserList(userindex).Stats.UserSkills(eSkill.Talar) >= 31 Then
                    Suerte = 24
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Talar) <= 50 _
   And UserList(userindex).Stats.UserSkills(eSkill.Talar) >= 41 Then
                    Suerte = 22
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Talar) <= 60 _
   And UserList(userindex).Stats.UserSkills(eSkill.Talar) >= 51 Then
                    Suerte = 20
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Talar) <= 70 _
   And UserList(userindex).Stats.UserSkills(eSkill.Talar) >= 61 Then
                    Suerte = 18
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Talar) <= 80 _
   And UserList(userindex).Stats.UserSkills(eSkill.Talar) >= 71 Then
                    Suerte = 15
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Talar) <= 90 _
   And UserList(userindex).Stats.UserSkills(eSkill.Talar) >= 81 Then
                    Suerte = 13
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Talar) <= 100 _
   And UserList(userindex).Stats.UserSkills(eSkill.Talar) >= 91 Then
                    Suerte = 7
End If
res = RandomNumber(1, Suerte)

If res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
    If UCase$(UserList(userindex).Clase) = "LEÑADOR" Then
        MiObj.Amount = RandomNumber(1, 5)
    Else
        MiObj.Amount = 1
    End If
    
    MiObj.ObjIndex = Leña
    
    
    If Not MeterItemEnInventario(userindex, MiObj) Then
        
        Call TirarItemAlPiso(UserList(userindex).pos, MiObj)
        
    End If
    
    Call SendData(SendTarget.toindex, userindex, 0, "||¡Has conseguido algo de leña!" & FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(userindex).flags.UltimoMensaje = 8 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||¡No has obtenido leña!" & FONTTYPE_INFO)
        UserList(userindex).flags.UltimoMensaje = 8
    End If
    '[/CDT]
End If

Call SubirSkill(userindex, Talar)



Exit Sub

errhandler:
    Call LogError("Error en DoTalar")

End Sub

Sub VolverCriminal(ByVal userindex As Integer)

If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.y).trigger = 6 Then Exit Sub

If UserList(userindex).flags.Privilegios < PlayerType.SemiDios Then
    UserList(userindex).Reputacion.BurguesRep = 0
    UserList(userindex).Reputacion.NobleRep = 0
    UserList(userindex).Reputacion.PlebeRep = 0
    UserList(userindex).Reputacion.BandidoRep = UserList(userindex).Reputacion.BandidoRep + vlASALTO
    If UserList(userindex).Reputacion.BandidoRep > MAXREP Then _
        UserList(userindex).Reputacion.BandidoRep = MAXREP
    If UserList(userindex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(userindex)
End If

End Sub

Sub VolverCiudadano(ByVal userindex As Integer)

If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.y).trigger = 6 Then Exit Sub

UserList(userindex).Reputacion.LadronesRep = 0
UserList(userindex).Reputacion.BandidoRep = 0
UserList(userindex).Reputacion.AsesinoRep = 0
UserList(userindex).Reputacion.PlebeRep = UserList(userindex).Reputacion.PlebeRep + vlASALTO
If UserList(userindex).Reputacion.PlebeRep > MAXREP Then _
    UserList(userindex).Reputacion.PlebeRep = MAXREP
End Sub




Public Sub DoMeditar(ByVal userindex As Integer)

UserList(userindex).Counters.IdleCount = 0

Dim Suerte As Integer
Dim res As Integer
Dim Cant As Integer

'Barrin 3/10/03
'Esperamos a que se termine de concentrar
Dim TActual As Long
TActual = GetTickCount() And &H7FFFFFFF
If TActual - UserList(userindex).Counters.tInicioMeditar < TIEMPO_INICIOMEDITAR Then
    Exit Sub
End If

If UserList(userindex).Counters.bPuedeMeditar = False Then
    UserList(userindex).Counters.bPuedeMeditar = True
End If

If UserList(userindex).Stats.MinMAN >= UserList(userindex).Stats.MaxMAN Then
    Call SendData(SendTarget.toindex, userindex, 0, "Z16")
    Call SendData(SendTarget.toindex, userindex, 0, "MEDOK")
    UserList(userindex).flags.Meditando = False
    UserList(userindex).char.FX = 0
    UserList(userindex).char.loops = 0
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & 0 & "," & 0)
    Exit Sub
End If

If UserList(userindex).Stats.UserSkills(eSkill.Meditar) <= 10 _
   And UserList(userindex).Stats.UserSkills(eSkill.Meditar) >= -1 Then
                    Suerte = 35
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Meditar) <= 20 _
   And UserList(userindex).Stats.UserSkills(eSkill.Meditar) >= 11 Then
                    Suerte = 30
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Meditar) <= 30 _
   And UserList(userindex).Stats.UserSkills(eSkill.Meditar) >= 21 Then
                    Suerte = 28
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Meditar) <= 40 _
   And UserList(userindex).Stats.UserSkills(eSkill.Meditar) >= 31 Then
                    Suerte = 24
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Meditar) <= 50 _
   And UserList(userindex).Stats.UserSkills(eSkill.Meditar) >= 41 Then
                    Suerte = 22
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Meditar) <= 60 _
   And UserList(userindex).Stats.UserSkills(eSkill.Meditar) >= 51 Then
                    Suerte = 20
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Meditar) <= 70 _
   And UserList(userindex).Stats.UserSkills(eSkill.Meditar) >= 61 Then
                    Suerte = 18
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Meditar) <= 80 _
   And UserList(userindex).Stats.UserSkills(eSkill.Meditar) >= 71 Then
                    Suerte = 15
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Meditar) <= 90 _
   And UserList(userindex).Stats.UserSkills(eSkill.Meditar) >= 81 Then
                    Suerte = 10
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Meditar) <= 100 _
   And UserList(userindex).Stats.UserSkills(eSkill.Meditar) >= 91 Then
                    Suerte = 5
End If
res = RandomNumber(1, Suerte)

If res = 1 Then
    Cant = Porcentaje(UserList(userindex).Stats.MaxMAN, 3)
    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN + Cant
    If UserList(userindex).Stats.MinMAN > UserList(userindex).Stats.MaxMAN Then _
        UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MaxMAN
    
    If Not UserList(userindex).flags.UltimoMensaje = 22 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||¡Has recuperado " & Cant & " puntos de mana!" & FONTTYPE_INFO)
        UserList(userindex).flags.UltimoMensaje = 22
    End If
    
    Call SendData(SendTarget.toindex, userindex, 0, "ASM" & UserList(userindex).Stats.MinMAN)
    Call SubirSkill(userindex, Meditar)
End If

End Sub



Public Sub Desarmar(ByVal userindex As Integer, ByVal VictimIndex As Integer)

Dim Suerte As Integer
Dim res As Integer

If UserList(userindex).Stats.UserSkills(eSkill.Wresterling) <= 10 _
   And UserList(userindex).Stats.UserSkills(eSkill.Wresterling) >= -1 Then
                    Suerte = 35
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) <= 20 _
   And UserList(userindex).Stats.UserSkills(eSkill.Wresterling) >= 11 Then
                    Suerte = 30
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) <= 30 _
   And UserList(userindex).Stats.UserSkills(eSkill.Wresterling) >= 21 Then
                    Suerte = 28
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) <= 40 _
   And UserList(userindex).Stats.UserSkills(eSkill.Wresterling) >= 31 Then
                    Suerte = 24
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) <= 50 _
   And UserList(userindex).Stats.UserSkills(eSkill.Wresterling) >= 41 Then
                    Suerte = 22
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) <= 60 _
   And UserList(userindex).Stats.UserSkills(eSkill.Wresterling) >= 51 Then
                    Suerte = 20
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) <= 70 _
   And UserList(userindex).Stats.UserSkills(eSkill.Wresterling) >= 61 Then
                    Suerte = 18
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) <= 80 _
   And UserList(userindex).Stats.UserSkills(eSkill.Wresterling) >= 71 Then
                    Suerte = 15
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) <= 90 _
   And UserList(userindex).Stats.UserSkills(eSkill.Wresterling) >= 81 Then
                    Suerte = 10
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) <= 100 _
   And UserList(userindex).Stats.UserSkills(eSkill.Wresterling) >= 91 Then
                    Suerte = 5
End If
res = RandomNumber(1, Suerte)

If res <= 2 Then
        Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
        Call SendData(SendTarget.toindex, userindex, 0, "||Has logrado desarmar a tu oponente!" & FONTTYPE_FIGHT)
        If UserList(VictimIndex).Stats.ELV < 20 Then Call SendData(SendTarget.toindex, VictimIndex, 0, "||Tu oponente te ha desarmado!" & FONTTYPE_FIGHT)
    End If
End Sub

