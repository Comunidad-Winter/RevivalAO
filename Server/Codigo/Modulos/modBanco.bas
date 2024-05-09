Attribute VB_Name = "modBanco"
Option Explicit

'MODULO PROGRAMADO POR NEB
'Kevin Birmingham
'kbneb@hotmail.com

Sub IniciarDeposito(ByVal userindex As Integer)
On Error GoTo errhandler

'Hacemos un Update del inventario del usuario
Call UpdateBanUserInv(True, userindex, 0)
'Atcualizamos el dinero
Call EnviarOro(userindex)
'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
SendData SendTarget.toIndex, userindex, 0, "INITBANCO"
UserList(userindex).flags.Comerciando = True

errhandler:

End Sub

Sub SendBanObj(userindex As Integer, Slot As Byte, Object As UserOBJ)


UserList(userindex).BancoInvent.Object(Slot) = Object

If Object.ObjIndex > 0 Then

    Call SendData(SendTarget.toIndex, userindex, 0, "SBO" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).name & "," & Object.Amount & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
    & ObjData(Object.ObjIndex).OBJType & "," _
    & ObjData(Object.ObjIndex).MaxHIT & "," _
    & ObjData(Object.ObjIndex).MinHIT & "," _
    & ObjData(Object.ObjIndex).MaxDef)

Else

    Call SendData(SendTarget.toIndex, userindex, 0, "SBO" & Slot & "," & "0" & "," & "(Vacío)" & "," & "0" & "," & "0")

End If


End Sub

Sub UpdateBanUserInv(ByVal UpdateAll As Boolean, ByVal userindex As Integer, ByVal Slot As Byte)

Dim NullObj As UserOBJ
Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(userindex).BancoInvent.Object(Slot).ObjIndex > 0 Then
        Call SendBanObj(userindex, Slot, UserList(userindex).BancoInvent.Object(Slot))
    Else
        Call SendBanObj(userindex, Slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS

        'Actualiza el inventario
        If UserList(userindex).BancoInvent.Object(LoopC).ObjIndex > 0 Then
            Call SendBanObj(userindex, LoopC, UserList(userindex).BancoInvent.Object(LoopC))
        Else
            
            Call SendBanObj(userindex, LoopC, NullObj)
            
        End If

    Next LoopC

End If

End Sub

Sub UserRetiraItem(ByVal userindex As Integer, ByVal i As Integer, ByVal Cantidad As Integer)
On Error GoTo errhandler


If Cantidad < 1 Then Exit Sub

   
       If UserList(userindex).BancoInvent.Object(i).Amount > 0 Then
            If Cantidad > UserList(userindex).BancoInvent.Object(i).Amount Then Cantidad = UserList(userindex).BancoInvent.Object(i).Amount
            'Agregamos el obj que compro al inventario
            Call UserReciveObj(userindex, CInt(i), Cantidad)
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, userindex, 0)
            'Actualizamos el banco
            Call UpdateBanUserInv(True, userindex, 0)
            'Actualizamos la ventana de comercio
            Call UpdateVentanaBanco(i, 0, userindex)
       End If



errhandler:

End Sub

Sub UserReciveObj(ByVal userindex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)

Dim Slot As Integer
Dim obji As Integer


If UserList(userindex).BancoInvent.Object(ObjIndex).Amount <= 0 Then Exit Sub

obji = UserList(userindex).BancoInvent.Object(ObjIndex).ObjIndex


'¿Ya tiene un objeto de este tipo?
Slot = 1
Do Until UserList(userindex).Invent.Object(Slot).ObjIndex = obji And _
   UserList(userindex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
    
    Slot = Slot + 1
    If Slot > MAX_INVENTORY_SLOTS Then
        Exit Do
    End If
Loop

'Sino se fija por un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(userindex).Invent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > MAX_INVENTORY_SLOTS Then
                Call SendData(SendTarget.toIndex, userindex, 0, "||No podés tener mas objetos." & FONTTYPE_INFO)
                Exit Sub
            End If
        Loop
        UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems + 1
End If



'Mete el obj en el slot
If UserList(userindex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
    
    'Menor que MAX_INV_OBJS
    UserList(userindex).Invent.Object(Slot).ObjIndex = obji
    UserList(userindex).Invent.Object(Slot).Amount = UserList(userindex).Invent.Object(Slot).Amount + Cantidad
    
    Call QuitarBancoInvItem(userindex, CByte(ObjIndex), Cantidad)
Else
    Call SendData(SendTarget.toIndex, userindex, 0, "||No podés tener mas objetos." & FONTTYPE_INFO)
End If


End Sub

Sub QuitarBancoInvItem(ByVal userindex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)



Dim ObjIndex As Integer
ObjIndex = UserList(userindex).BancoInvent.Object(Slot).ObjIndex

    'Quita un Obj

       UserList(userindex).BancoInvent.Object(Slot).Amount = UserList(userindex).BancoInvent.Object(Slot).Amount - Cantidad
        
        If UserList(userindex).BancoInvent.Object(Slot).Amount <= 0 Then
            UserList(userindex).BancoInvent.NroItems = UserList(userindex).BancoInvent.NroItems - 1
            UserList(userindex).BancoInvent.Object(Slot).ObjIndex = 0
            UserList(userindex).BancoInvent.Object(Slot).Amount = 0
        End If

    
    
End Sub

Sub UpdateVentanaBanco(ByVal Slot As Integer, ByVal NpcInv As Byte, ByVal userindex As Integer)
 
 
 Call SendData(SendTarget.toIndex, userindex, 0, "BANCOOK" & Slot & "," & NpcInv)
 
End Sub

Sub UserDepositaItem(ByVal userindex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer)

On Error GoTo errhandler

'El usuario deposita un item
   
If UserList(userindex).Invent.Object(Item).Amount > 0 And UserList(userindex).Invent.Object(Item).Equipped = 0 Then
            
            If Cantidad > 0 And Cantidad > UserList(userindex).Invent.Object(Item).Amount Then Cantidad = UserList(userindex).Invent.Object(Item).Amount
            'Agregamos el obj que compro al inventario
            Call UserDejaObj(userindex, CInt(Item), Cantidad)
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, userindex, 0)
            'Actualizamos el inventario del banco
            Call UpdateBanUserInv(True, userindex, 0)
            'Actualizamos la ventana del banco
            
            Call UpdateVentanaBanco(Item, 1, userindex)
            
End If

errhandler:

End Sub

Sub UserDejaObj(ByVal userindex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)

Dim Slot As Integer
Dim obji As Integer

If Cantidad < 1 Then Exit Sub

obji = UserList(userindex).Invent.Object(ObjIndex).ObjIndex

'¿Ya tiene un objeto de este tipo?
Slot = 1
Do Until UserList(userindex).BancoInvent.Object(Slot).ObjIndex = obji And _
         UserList(userindex).BancoInvent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
            Slot = Slot + 1
        
            If Slot > MAX_BANCOINVENTORY_SLOTS Then
                Exit Do
            End If
Loop

'Sino se fija por un slot vacio antes del slot devuelto
If Slot > MAX_BANCOINVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(userindex).BancoInvent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > MAX_BANCOINVENTORY_SLOTS Then
                Call SendData(SendTarget.toIndex, userindex, 0, "||No tienes mas espacio en el banco!!" & FONTTYPE_INFO)
                Exit Sub
                Exit Do
            End If
        Loop
        If Slot <= MAX_BANCOINVENTORY_SLOTS Then UserList(userindex).BancoInvent.NroItems = UserList(userindex).BancoInvent.NroItems + 1
        
        
End If

If Slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido
    'Mete el obj en el slot
    If UserList(userindex).BancoInvent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
        
        'Menor que MAX_INV_OBJS
        UserList(userindex).BancoInvent.Object(Slot).ObjIndex = obji
        UserList(userindex).BancoInvent.Object(Slot).Amount = UserList(userindex).BancoInvent.Object(Slot).Amount + Cantidad
        
        Call QuitarUserInvItem(userindex, CByte(ObjIndex), Cantidad)

    Else
        Call SendData(SendTarget.toIndex, userindex, 0, "||El banco no puede cargar tantos objetos." & FONTTYPE_INFO)
    End If

Else
    Call QuitarUserInvItem(userindex, CByte(ObjIndex), Cantidad)
End If

End Sub

Sub SendUserBovedaTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
On Error Resume Next
Dim j As Integer
Call SendData(SendTarget.toIndex, sendIndex, 0, "||" & UserList(userindex).name & FONTTYPE_INFO)
Call SendData(SendTarget.toIndex, sendIndex, 0, "|| Tiene " & UserList(userindex).BancoInvent.NroItems & " objetos." & FONTTYPE_INFO)
For j = 1 To MAX_BANCOINVENTORY_SLOTS
    If UserList(userindex).BancoInvent.Object(j).ObjIndex > 0 Then
        Call SendData(SendTarget.toIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(UserList(userindex).BancoInvent.Object(j).ObjIndex).name & " Cantidad:" & UserList(userindex).BancoInvent.Object(j).Amount & FONTTYPE_INFO)
    End If
Next

End Sub

Sub SendUserBovedaTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
On Error Resume Next
Dim j As Integer
Dim CharFile As String, Tmp As String
Dim ObjInd As Long, ObjCant As Long

CharFile = CharPath & CharName & ".chr"

If FileExist(CharFile, vbNormal) Then
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||" & CharName & FONTTYPE_INFO)
    Call SendData(SendTarget.toIndex, sendIndex, 0, "|| Tiene " & GetVar(CharFile, "BancoInventory", "CantidadItems") & " objetos." & FONTTYPE_INFO)
    For j = 1 To MAX_BANCOINVENTORY_SLOTS
        Tmp = GetVar(CharFile, "BancoInventory", "Obj" & j)
        ObjInd = ReadField(1, Tmp, Asc("-"))
        ObjCant = ReadField(2, Tmp, Asc("-"))
        If ObjInd > 0 Then
            Call SendData(SendTarget.toIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(ObjInd).name & " Cantidad:" & ObjCant & FONTTYPE_INFO)
        End If
    Next
Else
    Call SendData(SendTarget.toIndex, sendIndex, 0, "||Usuario inexistente: " & CharName & FONTTYPE_INFO)
End If

End Sub

