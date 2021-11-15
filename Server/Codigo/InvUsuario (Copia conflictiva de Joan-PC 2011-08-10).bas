Attribute VB_Name = "InvUsuario"


Option Explicit

Public Function TieneObjetosRobables(ByVal userindex As Integer) As Boolean

'17/09/02
'Agregue que la función se asegure que el objeto no es un barco

On Error Resume Next

Dim i As Integer
Dim ObjIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
    ObjIndex = UserList(userindex).Invent.Object(i).ObjIndex
    If ObjIndex > 0 Then
            If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And _
                ObjData(ObjIndex).OBJType <> eOBJType.otBarcos) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
    
    End If
Next i


End Function

Function ClasePuedeUsarItem(ByVal userindex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo manejador

'Call LogTarea("ClasePuedeUsarItem")

Dim flag As Boolean

If ObjData(ObjIndex).ClaseProhibida(1) <> "" Then
    
    Dim i As Integer
    For i = 1 To NUMCLASES
        If ObjData(ObjIndex).ClaseProhibida(i) = UCase$(UserList(userindex).Clase) Then
                ClasePuedeUsarItem = False
                Exit Function
        End If
    Next i
    
Else
    
    

End If

ClasePuedeUsarItem = True

Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Sub QuitarNewbieObj(ByVal userindex As Integer)
Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(userindex).Invent.Object(j).ObjIndex > 0 Then
             
             If ObjData(UserList(userindex).Invent.Object(j).ObjIndex).Newbie = 1 Then _
                    Call QuitarUserInvItem(userindex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, userindex, j)
        
        End If
Next

'[Barrin 17-12-03] Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon
'es transportado a su hogar de origen ;)
If UserList(userindex).pos.Map = 37 Then
    
    Dim DeDonde As WorldPos
    
    Select Case UCase$(UserList(userindex).Hogar)
        Case "LINDOS" 'Vamos a tener que ir por todo el desierto... uff!
            DeDonde = Lindos
        Case "ULLATHORPE"
            DeDonde = Ullathorpe
        Case "BANDERBILL"
            DeDonde = Banderbill
        Case Else
            DeDonde = Nix
    End Select
       
    Call WarpUserChar(userindex, DeDonde.Map, DeDonde.x, DeDonde.y, True)

End If
'[/Barrin]

End Sub

Sub LimpiarInventario(ByVal userindex As Integer)


Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        UserList(userindex).Invent.Object(j).ObjIndex = 0
        UserList(userindex).Invent.Object(j).Amount = 0
        UserList(userindex).Invent.Object(j).Equipped = 0
        
Next

UserList(userindex).Invent.NroItems = 0

UserList(userindex).Invent.ArmourEqpObjIndex = 0
UserList(userindex).Invent.ArmourEqpSlot = 0

UserList(userindex).Invent.WeaponEqpObjIndex = 0
UserList(userindex).Invent.WeaponEqpSlot = 0

UserList(userindex).Invent.CascoEqpObjIndex = 0
UserList(userindex).Invent.CascoEqpSlot = 0

UserList(userindex).Invent.EscudoEqpObjIndex = 0
UserList(userindex).Invent.EscudoEqpSlot = 0

UserList(userindex).Invent.HerramientaEqpObjIndex = 0
UserList(userindex).Invent.HerramientaEqpSlot = 0

UserList(userindex).Invent.MunicionEqpObjIndex = 0
UserList(userindex).Invent.MunicionEqpSlot = 0

UserList(userindex).Invent.BarcoObjIndex = 0
UserList(userindex).Invent.BarcoSlot = 0

End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal userindex As Integer)
On Error GoTo errhandler

If Cantidad > 100000 Or UserList(userindex).Stats.ELV < 30 Then Exit Sub

'SI EL NPC TIENE ORO LO TIRAMOS
If (Cantidad > 0) And (Cantidad <= UserList(userindex).Stats.GLD) Then
        Dim i As Byte
        Dim MiObj As Obj
        'info debug
        Dim loops As Integer
        
        'Seguridad Alkon
        If Cantidad > 39999 Then
            Dim j As Integer
            Dim k As Integer
            Dim m As Integer
            Dim Cercanos As String
            m = UserList(userindex).pos.Map
            For j = UserList(userindex).pos.x - 10 To UserList(userindex).pos.x + 10
                For k = UserList(userindex).pos.y - 10 To UserList(userindex).pos.y + 10
                    If InMapBounds(m, j, k) Then
                        If MapData(m, j, k).userindex > 0 Then
                            Cercanos = Cercanos & UserList(MapData(m, j, k).userindex).name & ","
                        End If
                    End If
                Next k
            Next j
            Call LogDesarrollo(UserList(userindex).name & " tira oro. Cercanos: " & Cercanos)
        End If
        '/Seguridad
        
        Do While (Cantidad > 0) And (UserList(userindex).Stats.GLD > 0)
            
            If Cantidad > MAX_INVENTORY_OBJS And UserList(userindex).Stats.GLD > MAX_INVENTORY_OBJS Then
                MiObj.Amount = MAX_INVENTORY_OBJS
                UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - MAX_INVENTORY_OBJS
                Cantidad = Cantidad - MiObj.Amount
            Else
                MiObj.Amount = Cantidad
                UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - Cantidad
                Cantidad = Cantidad - MiObj.Amount
            End If

            MiObj.ObjIndex = iORO
            
            If UserList(userindex).flags.Privilegios > PlayerType.User Then Call LogGM(UserList(userindex).name, "Tiro cantidad:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).name, False)
            
            Call TirarItemAlPiso(UserList(userindex).pos, MiObj)
            
            'info debug
            loops = loops + 1
            If loops > 100 Then
                LogError ("Error en tiraroro")
                Exit Sub
            End If
            
        Loop
    
End If

Exit Sub

errhandler:

End Sub

Sub QuitarUserInvItem(ByVal userindex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)

Dim MiObj As Obj
'Desequipa
If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

If UserList(userindex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(userindex, Slot)

'Quita un objeto
UserList(userindex).Invent.Object(Slot).Amount = UserList(userindex).Invent.Object(Slot).Amount - Cantidad
'¿Quedan mas?
If UserList(userindex).Invent.Object(Slot).Amount <= 0 Then
    UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems - 1
    UserList(userindex).Invent.Object(Slot).ObjIndex = 0
    UserList(userindex).Invent.Object(Slot).Amount = 0
End If
    
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal userindex As Integer, ByVal Slot As Byte)

Dim NullObj As UserOBJ
Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(userindex).Invent.Object(Slot).ObjIndex > 0 Then
        Call ChangeUserInv(userindex, Slot, UserList(userindex).Invent.Object(Slot))
    Else
        Call ChangeUserInv(userindex, Slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For LoopC = 1 To MAX_INVENTORY_SLOTS

        'Actualiza el inventario
        If UserList(userindex).Invent.Object(LoopC).ObjIndex > 0 Then
            Call ChangeUserInv(userindex, LoopC, UserList(userindex).Invent.Object(LoopC))
        Else
            
            Call ChangeUserInv(userindex, LoopC, NullObj)
            
        End If

    Next LoopC

End If

End Sub

Sub DropObj(ByVal userindex As Integer, ByVal Slot As Byte, ByVal num As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)

Dim Obj As Obj

If num > 0 Then
  
  If num > UserList(userindex).Invent.Object(Slot).Amount Then num = UserList(userindex).Invent.Object(Slot).Amount
  
  'Check objeto en el suelo
  If MapData(UserList(userindex).pos.Map, x, y).OBJInfo.ObjIndex = 0 Or MapData(UserList(userindex).pos.Map, x, y).OBJInfo.ObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex Then
        If UserList(userindex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(userindex, Slot)
        Obj.ObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
        
'        If ObjData(Obj.ObjIndex).Newbie = 1 And EsNewbie(UserIndex) Then
'            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podes tirar el objeto." & FONTTYPE_INFO)
'            Exit Sub
'        End If
        
        If num + MapData(UserList(userindex).pos.Map, x, y).OBJInfo.Amount > MAX_INVENTORY_OBJS Then
            num = MAX_INVENTORY_OBJS - MapData(UserList(userindex).pos.Map, x, y).OBJInfo.Amount
        End If
        
        Obj.Amount = num
        
        Call MakeObj(SendTarget.ToMap, 0, Map, Obj, Map, x, y)
        Call QuitarUserInvItem(userindex, Slot, num)
        Call UpdateUserInv(False, userindex, Slot)
        
        If ObjData(Obj.ObjIndex).OBJType = eOBJType.otBarcos Then
            Call SendData(SendTarget.toindex, userindex, 0, "||¡¡ATENCION!! ¡ACABAS DE TIRAR TU BARCA!" & FONTTYPE_TALK)
        End If
        If ObjData(Obj.ObjIndex).OBJType = eOBJType.otMontura And UserList(userindex).flags.Montado = True Then
        UserList(userindex).char.Body = UserList(userindex).flags.NumeroMont
          '[MaTeO 9]
          Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
        '[/MaTeO 9]
          UserList(userindex).flags.Montado = False
        End If
        If ObjData(Obj.ObjIndex).Caos = 1 Or ObjData(Obj.ObjIndex).Real = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||¡ATENCION!! ¡¡ACABAS DE TIRAR TU ARMADURA FACCIONARIA!!" & FONTTYPE_TALK)
        End If
        
        If UserList(userindex).flags.Privilegios > PlayerType.User Then Call LogGM("EDITADOS", UserList(userindex).name & " tiró " & num & " " & ObjData(Obj.ObjIndex).name, False)
  Else
    Call SendData(SendTarget.toindex, userindex, 0, "||No hay espacio en el piso." & FONTTYPE_INFO)
  End If
    
End If

End Sub

Sub EraseObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal num As Integer, ByVal Map As Byte, ByVal x As Integer, ByVal y As Integer)

MapData(Map, x, y).OBJInfo.Amount = MapData(Map, x, y).OBJInfo.Amount - num

If MapData(Map, x, y).OBJInfo.Amount <= 0 Then
    MapData(Map, x, y).OBJInfo.ObjIndex = 0
    MapData(Map, x, y).OBJInfo.Amount = 0
    
    If sndRoute = SendTarget.ToMap Then
        Call SendToAreaByPos(Map, x, y, "BO" & x & "," & y)
   Else
        Call SendData(sndRoute, sndIndex, sndMap, "BO" & x & "," & y)
    End If
End If

End Sub

Sub MakeObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Obj As Obj, Map As Integer, ByVal x As Integer, ByVal y As Integer)

If Obj.ObjIndex > 0 And Obj.ObjIndex <= UBound(ObjData) Then

    If MapData(Map, x, y).OBJInfo.ObjIndex = Obj.ObjIndex Then
        MapData(Map, x, y).OBJInfo.Amount = MapData(Map, x, y).OBJInfo.Amount + Obj.Amount
    Else
        MapData(Map, x, y).OBJInfo = Obj
        
        If sndRoute = SendTarget.ToMap Then
            Call ModAreas.SendToAreaByPos(Map, x, y, "HO" & ObjData(Obj.ObjIndex).GrhIndex & "," & x & "," & y)
        Else
            Call SendData(sndRoute, sndIndex, sndMap, "HO" & ObjData(Obj.ObjIndex).GrhIndex & "," & x & "," & y)
        End If
    End If
End If

End Sub

Function MeterItemEnInventario(ByVal userindex As Integer, ByRef MiObj As Obj) As Boolean
On Error GoTo errhandler

'Call LogTarea("MeterItemEnInventario")
 
Dim x As Integer
Dim y As Integer
Dim Slot As Byte

'¿el user ya tiene un objeto del mismo tipo?
Slot = 1
Do Until UserList(userindex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And _
         UserList(userindex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
   Slot = Slot + 1
   If Slot > MAX_INVENTORY_SLOTS Then
         Exit Do
   End If
Loop
    
'Sino busca un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
   Slot = 1
   Do Until UserList(userindex).Invent.Object(Slot).ObjIndex = 0
       Slot = Slot + 1
       If Slot > MAX_INVENTORY_SLOTS Then
           Call SendData(SendTarget.toindex, userindex, 0, "Z24")
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems + 1
End If
    
'Mete el objeto
If UserList(userindex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
   'Menor que MAX_INV_OBJS
   UserList(userindex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
   UserList(userindex).Invent.Object(Slot).Amount = UserList(userindex).Invent.Object(Slot).Amount + MiObj.Amount
Else
   UserList(userindex).Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
End If
    
MeterItemEnInventario = True
       
Call UpdateUserInv(False, userindex, Slot)


Exit Function
errhandler:

End Function


Sub GetObj(ByVal userindex As Integer)

Dim Obj As ObjData
Dim MiObj As Obj

'¿Hay algun obj?
If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.y).OBJInfo.ObjIndex > 0 Then
    '¿Esta permitido agarrar este obj?
    If ObjData(MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.y).OBJInfo.ObjIndex).Agarrable <> 1 Then
        Dim x As Integer
        Dim y As Integer
        Dim Slot As Byte
        
        x = UserList(userindex).pos.x
        y = UserList(userindex).pos.y
        Obj = ObjData(MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.y).OBJInfo.ObjIndex)
        MiObj.Amount = MapData(UserList(userindex).pos.Map, x, y).OBJInfo.Amount
        MiObj.ObjIndex = MapData(UserList(userindex).pos.Map, x, y).OBJInfo.ObjIndex
        
        If Not MeterItemEnInventario(userindex, MiObj) Then
            Call SendData(SendTarget.toindex, userindex, 0, "Z24")
        Else
            'Quitamos el objeto
            Call EraseObj(SendTarget.ToMap, 0, UserList(userindex).pos.Map, MapData(UserList(userindex).pos.Map, x, y).OBJInfo.Amount, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.y)
            If UserList(userindex).flags.Privilegios > PlayerType.User Then Call LogGM(UserList(userindex).name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).name, False)
        End If
        
    End If
Else
End If

End Sub

Sub Desequipar(ByVal userindex As Integer, ByVal Slot As Byte)
'Desequipa el item slot del inventario
Dim Obj As ObjData


If (Slot < LBound(UserList(userindex).Invent.Object)) Or (Slot > UBound(UserList(userindex).Invent.Object)) Then
    Exit Sub
ElseIf UserList(userindex).Invent.Object(Slot).ObjIndex = 0 Then
    Exit Sub
End If

Obj = ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex)

Select Case Obj.OBJType
    Case eOBJType.otWeapon
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        UserList(userindex).Invent.WeaponEqpObjIndex = 0
        UserList(userindex).Invent.WeaponEqpSlot = 0
        If Not UserList(userindex).flags.Mimetizado = 1 Then
            UserList(userindex).char.WeaponAnim = NingunArma
             '[MaTeO 9]
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
            '[/MaTeO 9]
        End If
    
    Case eOBJType.otFlechas
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        UserList(userindex).Invent.MunicionEqpObjIndex = 0
        UserList(userindex).Invent.MunicionEqpSlot = 0
    
    Case eOBJType.otHerramientas
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        UserList(userindex).Invent.HerramientaEqpObjIndex = 0
        UserList(userindex).Invent.HerramientaEqpSlot = 0
    
    Case eOBJType.otArmadura
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        UserList(userindex).Invent.ArmourEqpObjIndex = 0
        UserList(userindex).Invent.ArmourEqpSlot = 0
        Call DarCuerpoDesnudo(userindex, UserList(userindex).flags.Mimetizado = 1)
          '[MaTeO 9]
        Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
        '[/MaTeO 9]
            
    Case eOBJType.otCASCO
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        UserList(userindex).Invent.CascoEqpObjIndex = 0
        UserList(userindex).Invent.CascoEqpSlot = 0
         
        If Not UserList(userindex).flags.Mimetizado = 1 Then
            UserList(userindex).char.CascoAnim = NingunCasco
            '[MaTeO 9]
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
            '[/MaTeO 9]
        End If
    
    
    Case eOBJType.otESCUDO
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        UserList(userindex).Invent.EscudoEqpObjIndex = 0
        UserList(userindex).Invent.EscudoEqpSlot = 0
        If Not UserList(userindex).flags.Mimetizado = 1 Then
            UserList(userindex).char.ShieldAnim = NingunEscudo
          '[MaTeO 9]
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
            '[/MaTeO 9]
        End If
           '[MaTeO 9]
    Case eOBJType.otAlas
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        UserList(userindex).Invent.AlaEqpObjIndex = 0
        UserList(userindex).Invent.AlaEqpSlot = 0
        If Not UserList(userindex).flags.Mimetizado = 1 Then
            UserList(userindex).char.Alas = NingunAlas
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
        End If
    '[/MaTeO 9]
End Select

Call EnviarSta(userindex)
Call UpdateUserInv(False, userindex, Slot)
Call SendUserHitBox(userindex)

End Sub

Function SexoPuedeUsarItem(ByVal userindex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo errhandler

If ObjData(ObjIndex).Mujer = 1 Then
    SexoPuedeUsarItem = UCase$(UserList(userindex).Genero) <> "HOMBRE"
ElseIf ObjData(ObjIndex).Hombre = 1 Then
    SexoPuedeUsarItem = UCase$(UserList(userindex).Genero) <> "MUJER"
Else
    SexoPuedeUsarItem = True
End If

Exit Function
errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal userindex As Integer, ByVal ObjIndex As Integer) As Boolean

If ObjData(ObjIndex).Real = 1 Then
    If Not Criminal(userindex) Then
        FaccionPuedeUsarItem = (UserList(userindex).Faccion.ArmadaReal = 1)
    Else
        FaccionPuedeUsarItem = False
    End If
ElseIf ObjData(ObjIndex).Caos = 1 Then
    If Criminal(userindex) Then
        FaccionPuedeUsarItem = (UserList(userindex).Faccion.FuerzasCaos = 1)
    Else
        FaccionPuedeUsarItem = False
    End If
Else
    FaccionPuedeUsarItem = True
End If

End Function

Sub EquiparInvItem(ByVal userindex As Integer, ByVal Slot As Byte)
On Error GoTo errhandler

'Equipa un item del inventario
Dim Obj As ObjData
Dim ObjIndex As Integer

ObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
Obj = ObjData(ObjIndex)

If Obj.Newbie = 1 And Not EsNewbie(userindex) Then
     Call SendData(SendTarget.toindex, userindex, 0, "||Solo los newbies pueden usar este objeto." & FONTTYPE_INFO)
     Exit Sub
End If
        
Select Case Obj.OBJType
    '[MaTeO 9]
    Case eOBJType.otAlas
       If ClasePuedeUsarItem(userindex, ObjIndex) And _
          FaccionPuedeUsarItem(userindex, ObjIndex) Then
                'Si esta equipado lo quita
                If UserList(userindex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(userindex, Slot)
                    'Animacion por defecto
                    If UserList(userindex).flags.Mimetizado = 1 Then
                        UserList(userindex).CharMimetizado.WeaponAnim = NingunAlas
                    Else
                        UserList(userindex).char.WeaponAnim = NingunAlas
                        '[MaTeO 9]
                        Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
                        '[/MaTeO 9]
                    End If
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(userindex).Invent.AlaEqpObjIndex > 0 Then
                    Call Desequipar(userindex, UserList(userindex).Invent.AlaEqpSlot)
                End If
        
                UserList(userindex).Invent.Object(Slot).Equipped = 1
                UserList(userindex).Invent.AlaEqpObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
                UserList(userindex).Invent.AlaEqpSlot = Slot
        
                If UserList(userindex).flags.Mimetizado = 1 Then
                    UserList(userindex).CharMimetizado.Alas = Obj.Ropaje
                Else
                    UserList(userindex).char.Alas = Obj.Ropaje
                    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
                End If
       Else
            Call SendData(SendTarget.toindex, userindex, 0, "Z42")
       End If
    '[/MaTeO 9]
    Case eOBJType.otWeapon
    If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then
If ObjData(ObjIndex).DosManos = 1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Debes desequiparte tu escudo para poder usar esta arma." & FONTTYPE_INFO)
Exit Sub
End If
End If
       If ClasePuedeUsarItem(userindex, ObjIndex) And _
          FaccionPuedeUsarItem(userindex, ObjIndex) Then
                'Si esta equipado lo quita
                If UserList(userindex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(userindex, Slot)
                    'Animacion por defecto
                    If UserList(userindex).flags.Mimetizado = 1 Then
                        UserList(userindex).CharMimetizado.WeaponAnim = NingunArma
                    Else
                        UserList(userindex).char.WeaponAnim = NingunArma
                            '[MaTeO 9]
                        Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
                        '[/MaTeO 9]
                    End If
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                    Call Desequipar(userindex, UserList(userindex).Invent.WeaponEqpSlot)
                End If
        
                UserList(userindex).Invent.Object(Slot).Equipped = 1
                UserList(userindex).Invent.WeaponEqpObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
                UserList(userindex).Invent.WeaponEqpSlot = Slot
                
                'Sonido
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_SACARARMA)
        
                If UserList(userindex).flags.Mimetizado = 1 Then
                    UserList(userindex).CharMimetizado.WeaponAnim = Obj.WeaponAnim
                Else
                    UserList(userindex).char.WeaponAnim = Obj.WeaponAnim
                     '[MaTeO 9]
                    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
                    '[/MaTeO 9]
                End If
       Else
            Call SendData(SendTarget.toindex, userindex, 0, "Z42")
       End If

         Case eOBJType.otPARAA
If UserList(userindex).flags.Muerto = 1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo." & FONTTYPE_INFO)
Exit Sub
End If
If UserList(userindex).flags.Paralizado = 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡¡No estás Paralizado!! " & FONTTYPE_INFO)
Exit Sub
End If
If UserList(userindex).flags.Paralizado = 1 Then
UserList(userindex).flags.Paralizado = 0
Call SendData(SendTarget.toindex, userindex, 0, "PARADOW")
Call SendData(SendTarget.toindex, userindex, 0, "PU" & UserList(userindex).pos.x & "," & UserList(userindex).pos.y)
Call SendData(SendTarget.toindex, userindex, 0, "||Te has quitado la paralisis." & FONTTYPE_INFO)
Call QuitarUserInvItem(userindex, Slot, 1)
Call UpdateUserInv(False, userindex, Slot)
End If
    Case eOBJType.otHerramientas
       If ClasePuedeUsarItem(userindex, ObjIndex) And _
          FaccionPuedeUsarItem(userindex, ObjIndex) Then
                'Si esta equipado lo quita
                If UserList(userindex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(userindex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(userindex).Invent.HerramientaEqpObjIndex > 0 Then
                    Call Desequipar(userindex, UserList(userindex).Invent.HerramientaEqpSlot)
                End If
        
                UserList(userindex).Invent.Object(Slot).Equipped = 1
                UserList(userindex).Invent.HerramientaEqpObjIndex = ObjIndex
                UserList(userindex).Invent.HerramientaEqpSlot = Slot
                
       Else
            Call SendData(SendTarget.toindex, userindex, 0, "Z42")
       End If
    
    Case eOBJType.otFlechas
       If ClasePuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) And _
          FaccionPuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) Then
                
                'Si esta equipado lo quita
                If UserList(userindex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(userindex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(userindex).Invent.MunicionEqpObjIndex > 0 Then
                    Call Desequipar(userindex, UserList(userindex).Invent.MunicionEqpSlot)
                End If
        
                UserList(userindex).Invent.Object(Slot).Equipped = 1
                UserList(userindex).Invent.MunicionEqpObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
                UserList(userindex).Invent.MunicionEqpSlot = Slot
                
       Else
            Call SendData(SendTarget.toindex, userindex, 0, "Z42")
       End If
    
    Case eOBJType.otArmadura
        If UserList(userindex).flags.Navegando = 1 Then Exit Sub
        'Nos aseguramos que puede usarla
        If ClasePuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) And _
           SexoPuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) And _
           CheckRazaUsaRopa(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) And _
           FaccionPuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) Then
           
           'Si esta equipado lo quita
            If UserList(userindex).Invent.Object(Slot).Equipped Then
                Call Desequipar(userindex, Slot)
                Call DarCuerpoDesnudo(userindex, UserList(userindex).flags.Mimetizado = 1)
                If Not UserList(userindex).flags.Mimetizado = 1 Then
                      '[MaTeO 9]
                    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
                    '[/MaTeO 9]
                End If
                Exit Sub
            End If
    
            'Quita el anterior
            If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
                Call Desequipar(userindex, UserList(userindex).Invent.ArmourEqpSlot)
            End If
    
            'Lo equipa
           
            
            UserList(userindex).Invent.Object(Slot).Equipped = 1
            UserList(userindex).Invent.ArmourEqpObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
            UserList(userindex).Invent.ArmourEqpSlot = Slot
                
            If UserList(userindex).flags.Mimetizado = 1 Then
                UserList(userindex).CharMimetizado.Body = Obj.Ropaje
            Else
                UserList(userindex).char.Body = Obj.Ropaje
                    '[MaTeO 9]
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
                '[/MaTeO 9]
            End If
            UserList(userindex).flags.Desnudo = 0
            

        Else
            Call SendData(SendTarget.toindex, userindex, 0, "Z42")
        End If
    
    Case eOBJType.otCASCO
        If UserList(userindex).flags.Navegando = 1 Then Exit Sub
        If ClasePuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) Then
            'Si esta equipado lo quita
            If UserList(userindex).Invent.Object(Slot).Equipped Then
                Call Desequipar(userindex, Slot)
                If UserList(userindex).flags.Mimetizado = 1 Then
                    UserList(userindex).CharMimetizado.CascoAnim = NingunCasco
                Else
                    UserList(userindex).char.CascoAnim = NingunCasco
                    '[MaTeO 9]
                    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
                    '[/MaTeO 9]
                End If
                Exit Sub
            End If
    
            'Quita el anterior
            If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then
                Call Desequipar(userindex, UserList(userindex).Invent.CascoEqpSlot)
            End If
    
            'Lo equipa
       
         
            UserList(userindex).Invent.Object(Slot).Equipped = 1
            UserList(userindex).Invent.CascoEqpObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
            UserList(userindex).Invent.CascoEqpSlot = Slot
            If UserList(userindex).flags.Mimetizado = 1 Then
                UserList(userindex).CharMimetizado.CascoAnim = Obj.CascoAnim
            Else
                UserList(userindex).char.CascoAnim = Obj.CascoAnim
                '[MaTeO 9]
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
                '[/MaTeO 9]
            End If
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "Z42")
        End If
        
    
    Case eOBJType.otESCUDO
    If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).DosManos = 1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Debes desequiparte tu arma para poder usar este escudo." & FONTTYPE_INFO)
Exit Sub
End If
End If

        If UserList(userindex).flags.Navegando = 1 Then Exit Sub
         If ClasePuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) And _
             FaccionPuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) Then

             'Si esta equipado lo quita
             If UserList(userindex).Invent.Object(Slot).Equipped Then
                 Call Desequipar(userindex, Slot)
                 If UserList(userindex).flags.Mimetizado = 1 Then
                     UserList(userindex).CharMimetizado.ShieldAnim = NingunEscudo
                 Else
                     UserList(userindex).char.ShieldAnim = NingunEscudo
                     '[MaTeO 9]
                     Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
                     '[/MaTeO 9]
                 End If
                 Exit Sub
             End If
     
             'Quita el anterior
             If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then
                 Call Desequipar(userindex, UserList(userindex).Invent.EscudoEqpSlot)
             End If
     
             'Lo equipa
             
             UserList(userindex).Invent.Object(Slot).Equipped = 1
             UserList(userindex).Invent.EscudoEqpObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
             UserList(userindex).Invent.EscudoEqpSlot = Slot
             
             If UserList(userindex).flags.Mimetizado = 1 Then
                 UserList(userindex).CharMimetizado.ShieldAnim = Obj.ShieldAnim
             Else
                 UserList(userindex).char.ShieldAnim = Obj.ShieldAnim
                 '[MaTeO 9]
                 Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
                 '[/MaTeO 9]
             End If
         Else
             Call SendData(SendTarget.toindex, userindex, 0, "Z42")
         End If
End Select

'Actualiza
Call UpdateUserInv(False, userindex, Slot)
Call SendUserHitBox(userindex)
Exit Sub
errhandler:
Call LogError("EquiparInvItem Slot:" & Slot)
End Sub

Private Function CheckRazaUsaRopa(ByVal userindex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo errhandler

'Verifica si la raza puede usar la ropa
If UserList(userindex).Raza = "Humano" Or _
   UserList(userindex).Raza = "Elfo" Or _
   UserList(userindex).Raza = "Elfo Oscuro" Then
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
End If


Exit Function
errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal userindex As Integer, ByVal Slot As Byte)

'Usa un item del inventario
Dim Obj As ObjData
Dim ObjIndex As Integer
Dim TargObj As ObjData
Dim MiObj As Obj

If UserList(userindex).Invent.Object(Slot).Amount = 0 Then Exit Sub

Obj = ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex)

If Obj.Newbie = 1 And Not EsNewbie(userindex) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Solo los newbies pueden usar estos objetos." & FONTTYPE_INFO)
    Exit Sub
End If

If Obj.OBJType = eOBJType.otWeapon Then

    If Obj.proyectil = 1 Then
        'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
        If Not IntervaloPermiteUsarArcos(userindex, False) Then Exit Sub
    Else
        'dagas
        If Not IntervaloPermiteUsar(userindex) Then Exit Sub
    End If
Else
    If Not IntervaloPermiteUsar(userindex) Then Exit Sub
End If

ObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
UserList(userindex).flags.TargetObjInvIndex = ObjIndex
UserList(userindex).flags.TargetObjInvSlot = Slot

Select Case Obj.OBJType
    Case eOBJType.otUseOnce
        If UserList(userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "Z12")
            Exit Sub
        End If

        'Usa el item
        UserList(userindex).Stats.MinHam = UserList(userindex).Stats.MinHam + Obj.MinHam
        If UserList(userindex).Stats.MinHam > UserList(userindex).Stats.MaxHam Then _
            UserList(userindex).Stats.MinHam = UserList(userindex).Stats.MaxHam
        UserList(userindex).flags.Hambre = 0
        Call EnviarHambreYsed(userindex)
        'Sonido
        
        If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, e_SoundIndex.MORFAR_MANZANA)
        Else
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, e_SoundIndex.SOUND_COMIDA)
        End If
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(userindex, Slot, 1)
        
        Call UpdateUserInv(False, userindex, Slot)

    Case eOBJType.otGuita
        If UserList(userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "Z12")
            Exit Sub
        End If
        
        UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + UserList(userindex).Invent.Object(Slot).Amount
        UserList(userindex).Invent.Object(Slot).Amount = 0
        UserList(userindex).Invent.Object(Slot).ObjIndex = 0
        UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems - 1
        
        Call UpdateUserInv(False, userindex, Slot)
        Call EnviarOro(userindex)
        
    Case eOBJType.otWeapon
        If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
        End If
        
        If ObjData(ObjIndex).proyectil = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "T01" & Proyectiles)
            If UserList(userindex).flags.Invisible = 1 Or UserList(userindex).flags.Oculto = 1 Then
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).char.CharIndex & ",0")
        Call SendData(SendTarget.toindex, userindex, 0, "Z11")
        UserList(userindex).flags.Invisible = 0
        UserList(userindex).flags.Oculto = 0
        UserList(userindex).Counters.Invisibilidad = 0
        End If
        Else
            If UserList(userindex).flags.TargetObj = 0 Then Exit Sub
            
            '¿El target-objeto es leña?
            If UserList(userindex).flags.TargetObj = Leña Then
                If UserList(userindex).Invent.Object(Slot).ObjIndex = DAGA Then
                    Call TratarDeHacerFogata(UserList(userindex).flags.TargetObjMap, _
                         UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY, userindex)
                End If
            End If
        End If
    
    Case eOBJType.otPociones
'If UserList(userindex).Lac.LPociones.Puedo = False Then Exit Sub
        If UserList(userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "Z12")
            Exit Sub
        End If
        
        UserList(userindex).flags.TomoPocion = True
        UserList(userindex).flags.TipoPocion = Obj.TipoPocion
                
        Select Case UserList(userindex).flags.TipoPocion
        
            Case 1 'Modif la agilidad
                UserList(userindex).flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then _
                    UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
                If UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) > 2 * UserList(userindex).Stats.UserAtributosBackUP(Agilidad) Then UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = 2 * UserList(userindex).Stats.UserAtributosBackUP(Agilidad)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, Slot, 1)
                Call EnviarDopa(userindex)
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_BEBER)
                
            Case 2 'Modif la fuerza
                UserList(userindex).flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then _
                    UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
                If UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) > 2 * UserList(userindex).Stats.UserAtributosBackUP(Fuerza) Then UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = 2 * UserList(userindex).Stats.UserAtributosBackUP(Fuerza)
                
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, Slot, 1)
                Call EnviarDopa(userindex)
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_BEBER)
               
            Case 3 'Pocion roja, restaura HP
             If UserList(userindex).flags.Potea = False Then
 If Not UserList(userindex).pos.Map = 67 Then
 Call SendData(SendTarget.toindex, userindex, 0, "||Has sido encarcelado 3 Horas por uso de editor de paquetes, Agradece que no te baneamos =)" & FONTTYPE_INFO)
 Call Encarcelar(userindex, 180, "Servidor")
 End If
 Exit Sub
 End If
                'Usa el item
                UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If UserList(userindex).Stats.MinHP > UserList(userindex).Stats.MaxHP Then _
                    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_BEBER)
                Call EnviarHP(userindex)
                UserList(userindex).flags.Potea = False
            Case 4 'Pocion azul, restaura MANA
                         If UserList(userindex).flags.Potea = False Then
 If Not UserList(userindex).pos.Map = 67 Then
 Call SendData(SendTarget.toindex, userindex, 0, "||Has sido encarcelado 3 Horas por uso de editor de paquetes, Agradece que no te baneamos =)" & FONTTYPE_INFO)
 Call Encarcelar(userindex, 180, "Servidor")
 End If
 Exit Sub
 End If
                'Usa el item
                UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN + Porcentaje(UserList(userindex).Stats.MaxMAN, 5)
                If UserList(userindex).Stats.MinMAN > UserList(userindex).Stats.MaxMAN Then _
                    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MaxMAN
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_BEBER)
                Call EnviarMn(userindex)
                UserList(userindex).flags.Potea = False
            Case 5 ' Pocion violeta
                If UserList(userindex).flags.Envenenado = 1 Then
                    UserList(userindex).flags.Envenenado = 0
                    Call SendData(SendTarget.toindex, userindex, 0, "||Te has curado del envenenamiento." & FONTTYPE_INFO)
                End If
                
            Case 7
            If UserList(userindex).flags.Paralizado = 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡¡No estás Paralizado!! " & FONTTYPE_INFO)
Exit Sub
End If
If UserList(userindex).flags.Paralizado = 1 Then
UserList(userindex).flags.Paralizado = 0
Call SendData(SendTarget.toindex, userindex, 0, "PARADOW")
Call SendData(SendTarget.toindex, userindex, 0, "PU" & UserList(userindex).pos.x & "," & UserList(userindex).pos.y)
Call SendData(SendTarget.toindex, userindex, 0, "||Te has quitado la paralisis." & FONTTYPE_INFO)
End If
                'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_BEBER)
                
            Case 6  ' Pocion Negra
                If UserList(userindex).flags.Privilegios = PlayerType.User Then
                    Call QuitarUserInvItem(userindex, Slot, 1)
                    Call UserDie(userindex)
                    Call SendData(SendTarget.toindex, userindex, 0, "||Sientes un gran mareo y pierdes el conocimiento." & FONTTYPE_FIGHT)
                End If
               
       End Select
       Call UpdateUserInv(False, userindex, Slot)

     Case eOBJType.otBebidas
        If UserList(userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "Z12")
            Exit Sub
        End If
        UserList(userindex).Stats.MinAGU = UserList(userindex).Stats.MinAGU + Obj.MinSed
        If UserList(userindex).Stats.MinAGU > UserList(userindex).Stats.MaxAGU Then _
            UserList(userindex).Stats.MinAGU = UserList(userindex).Stats.MaxAGU
        UserList(userindex).flags.Sed = 0
        Call EnviarHambreYsed(userindex)
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(userindex, Slot, 1)
        
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_BEBER)
        
        Call UpdateUserInv(False, userindex, Slot)
    
    Case eOBJType.otLlaves
        If UserList(userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "Z12")
            Exit Sub
        End If
        
        If UserList(userindex).flags.TargetObj = 0 Then Exit Sub
        TargObj = ObjData(UserList(userindex).flags.TargetObj)
        '¿El objeto clickeado es una puerta?
        If TargObj.OBJType = eOBJType.otPuertas Then
            '¿Esta cerrada?
            If TargObj.Cerrada = 1 Then
                  '¿Cerrada con llave?
                  If TargObj.Llave > 0 Then
                     If TargObj.clave = Obj.clave Then
         
                        MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.ObjIndex _
                        = ObjData(MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.ObjIndex).IndexCerrada
                        UserList(userindex).flags.TargetObj = MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.ObjIndex
                        Call SendData(SendTarget.toindex, userindex, 0, "||Has abierto la puerta." & FONTTYPE_INFO)
                        Exit Sub
                     Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||La llave no sirve." & FONTTYPE_INFO)
                        Exit Sub
                     End If
                  Else
                     If TargObj.clave = Obj.clave Then
                        MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.ObjIndex _
                        = ObjData(MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.ObjIndex).IndexCerradaLlave
                        Call SendData(SendTarget.toindex, userindex, 0, "||Has cerrado con llave la puerta." & FONTTYPE_INFO)
                        UserList(userindex).flags.TargetObj = MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.ObjIndex
                        Exit Sub
                     Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||La llave no sirve." & FONTTYPE_INFO)
                        Exit Sub
                     End If
                  End If
            Else
                  Call SendData(SendTarget.toindex, userindex, 0, "||No esta cerrada." & FONTTYPE_INFO)
                  Exit Sub
            End If
            
        End If
    
        Case eOBJType.otBotellaVacia
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
            End If
            If Not HayAgua(UserList(userindex).pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No hay agua allí." & FONTTYPE_INFO)
                Exit Sub
            End If
            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex).IndexAbierta
            Call QuitarUserInvItem(userindex, Slot, 1)
            If Not MeterItemEnInventario(userindex, MiObj) Then
                Call TirarItemAlPiso(UserList(userindex).pos, MiObj)
            End If
            
            Call UpdateUserInv(False, userindex, Slot)
    
        Case eOBJType.otBotellaLlena
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
            End If
            UserList(userindex).Stats.MinAGU = UserList(userindex).Stats.MinAGU + Obj.MinSed
            If UserList(userindex).Stats.MinAGU > UserList(userindex).Stats.MaxAGU Then _
                UserList(userindex).Stats.MinAGU = UserList(userindex).Stats.MaxAGU
            UserList(userindex).flags.Sed = 0
            Call EnviarHambreYsed(userindex)
            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex).IndexCerrada
            Call QuitarUserInvItem(userindex, Slot, 1)
            If Not MeterItemEnInventario(userindex, MiObj) Then
                Call TirarItemAlPiso(UserList(userindex).pos, MiObj)
            End If
            
            
        Case eOBJType.otHerramientas
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
            End If
            If Not UserList(userindex).Stats.MinSta > 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Estas muy cansado" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(userindex).Invent.Object(Slot).Equipped = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Antes de usar la herramienta deberias equipartela." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            UserList(userindex).Reputacion.PlebeRep = UserList(userindex).Reputacion.PlebeRep + vlProleta
            If UserList(userindex).Reputacion.PlebeRep > MAXREP Then _
                UserList(userindex).Reputacion.PlebeRep = MAXREP
            
            Select Case ObjIndex
                Case CAÑA_PESCA, RED_PESCA
                    Call SendData(SendTarget.toindex, userindex, 0, "T01" & Pesca)
                Case HACHA_LEÑADOR
                    Call SendData(SendTarget.toindex, userindex, 0, "T01" & Talar)
            End Select
        
Case eOBJType.otPergaminos
                If UserList(userindex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                    Exit Sub
                End If
               
                If UserList(userindex).flags.Hambre = 0 And _
                   UserList(userindex).flags.Sed = 0 Then
                If ClasePuedeUsarItem(userindex, ObjIndex) And _
                    FaccionPuedeUsarItem(userindex, ObjIndex) Then

                        Call AgregarHechizo(userindex, Slot)
                        Call UpdateUserInv(False, userindex, Slot)
                    Else
                        Call SendData(toindex, userindex, 0, "||Tú clase no puede aprender este hechizo." & FONTTYPE_INFO)
                    End If
                Else
                   Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado hambriento y sediento." & FONTTYPE_INFO)
                End If
       
       Case eOBJType.otMinerales
           If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
           End If
           Call SendData(SendTarget.toindex, userindex, 0, "T01" & FundirMetal)
       
       Case eOBJType.otInstrumentos
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
            End If
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & Obj.Snd1)
       
       Case eOBJType.otBarcos
       If UserList(userindex).flags.Montado = True Then
           Call SendData(SendTarget.toindex, userindex, 0, "||Estas Montado!" & FONTTYPE_INFO)
           Exit Sub
           End If
    'Verifica si esta aproximado al agua antes de permitirle navegar
        If UserList(userindex).Stats.ELV < 35 Then
            If UCase$(UserList(userindex).Clase) <> "PESCADOR" And UCase$(UserList(userindex).Clase) <> "PIRATA" Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Para recorrer los mares debes ser nivel 35 o superior." & FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        If ((LegalPos(UserList(userindex).pos.Map, UserList(userindex).pos.x - 1, UserList(userindex).pos.y, True) Or _
            LegalPos(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.y - 1, True) Or _
            LegalPos(UserList(userindex).pos.Map, UserList(userindex).pos.x + 1, UserList(userindex).pos.y, True) Or _
            LegalPos(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.y + 1, True)) And _
            UserList(userindex).flags.Navegando = 0) _
            Or UserList(userindex).flags.Navegando = 1 Then
           Call DoNavega(userindex, Obj, Slot)
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||¡Debes aproximarte al agua para usar el barco!" & FONTTYPE_INFO)
        End If
           Case eOBJType.otMontura
          ObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
        Obj = ObjData(ObjIndex)
        
           If UserList(userindex).flags.Muerto = 1 Then
           Call SendData(SendTarget.toindex, userindex, 0, "||Estas muerto!" & FONTTYPE_INFO)
           Exit Sub
           End If
           If UserList(userindex).flags.Navegando = 1 Then
           Call SendData(SendTarget.toindex, userindex, 0, "||Estas navegando!!" & FONTTYPE_INFO)
           Exit Sub
           End If
           If UserList(userindex).pos.Map = 66 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes usar tu mascota en guerra!!" & FONTTYPE_INFO)
           Exit Sub
           End If
          If UserList(userindex).flags.Montado = True Then
          UserList(userindex).char.Body = UserList(userindex).flags.NumeroMont
              '[MaTeO 9]
          Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
          '[/MaTeO 9]
          UserList(userindex).flags.Montado = False
          Exit Sub
          End If
          
           If UserList(userindex).flags.Montado = False Then
           UserList(userindex).flags.NumeroMont = UserList(userindex).char.Body
           UserList(userindex).char.Body = Obj.Ropaje
                '[MaTeO 9]
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
                '[/MaTeO 9]
                UserList(userindex).flags.Montado = True
                End If
End Select
'[MaTeO 13]
If ObjIndex = 4295 Then
    Call WarpUserChar(userindex, 117, 50, 72, True)
    Call QuitarUserInvItem(userindex, Slot, 1)
End If
'[/MaTeO 13]
'Actualiza
'call scenduserstatsbox(UserIndex)
'Call UpdateUserInv(False, UserIndex, Slot)

End Sub
Sub TirarTodo(ByVal userindex As Integer)
On Error Resume Next

'If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

Call TirarTodosLosItems(userindex)
Call TirarOro(UserList(userindex).Stats.GLD, userindex)

End Sub

Public Function ItemSeCae(ByVal Index As Integer) As Boolean

ItemSeCae = (ObjData(Index).Real <> 1 Or ObjData(Index).NoSeCae = 0) And _
            (ObjData(Index).Caos <> 1 Or ObjData(Index).NoSeCae = 0) And _
            ObjData(Index).OBJType <> eOBJType.otLlaves And _
            ObjData(Index).OBJType <> eOBJType.otBarcos And _
            ObjData(Index).NoSeCae = 0


End Function

Sub TirarTodosLosItems(ByVal userindex As Integer)
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    
    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(userindex).Invent.Object(i).ObjIndex
        If ItemIndex > 0 Then
             If ItemSeCae(ItemIndex) Then
                NuevaPos.x = 0
                NuevaPos.y = 0
                
                'Creo el Obj
                MiObj.Amount = UserList(userindex).Invent.Object(i).Amount
                MiObj.ObjIndex = ItemIndex
                
                Tilelibre UserList(userindex).pos, NuevaPos, MiObj
                If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then
                    Call DropObj(userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.x, NuevaPos.y)
                End If
             End If
        End If
    Next i
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean

ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function

Sub TirarTodosLosItemsNoNewbies(ByVal userindex As Integer)
Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As Obj
Dim ItemIndex As Integer

If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.y).trigger = 6 Then Exit Sub

For i = 1 To MAX_INVENTORY_SLOTS
    ItemIndex = UserList(userindex).Invent.Object(i).ObjIndex
    If ItemIndex > 0 Then
        If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
            NuevaPos.x = 0
            NuevaPos.y = 0
            
            'Creo MiObj
            MiObj.Amount = UserList(userindex).Invent.Object(i).ObjIndex
            MiObj.ObjIndex = ItemIndex
            
            Tilelibre UserList(userindex).pos, NuevaPos, MiObj
            If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then
                If MapData(NuevaPos.Map, NuevaPos.x, NuevaPos.y).OBJInfo.ObjIndex = 0 Then Call DropObj(userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.x, NuevaPos.y)
            End If
        End If
    End If
Next i

End Sub
