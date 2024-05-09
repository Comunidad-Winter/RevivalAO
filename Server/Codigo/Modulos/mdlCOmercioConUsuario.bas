Attribute VB_Name = "mdlCOmercioConUsuario"
'Modulo para comerciar con otro usuario
'Por Alejo (Alejandro Santos)
'
'
'[Alejo]
Option Explicit

Private Const MAX_ORO_LOGUEABLE As Long = 90000

Public Type tCOmercioUsuario
    DestUsu As Integer 'El otro Usuario
    DestNick As String
    Objeto As Integer 'Indice del inventario a comerciar, que objeto desea dar
    
    'El tipo de datos de Cant ahora es Long (antes Integer)
    'asi se puede comerciar con oro > 32k
    '[CORREGIDO]
    Cant As Long 'Cuantos comerciar, cuantos objetos desea dar
    '[/CORREGIDO]
    Acepto As Boolean
End Type

'origen: origen de la transaccion, originador del comando
'destino: receptor de la transaccion
Public Sub IniciarComercioConUsuario(ByVal Origen As Integer, ByVal Destino As Integer)
On Error GoTo errhandler

'Si ambos pusieron /comerciar entonces
If UserList(Origen).ComUsu.DestUsu = Destino And _
   UserList(Destino).ComUsu.DestUsu = Origen Then
    'Actualiza el inventario del usuario
    Call UpdateUserInv(True, Origen, 0)
    'Decirle al origen que abra la ventanita.
    Call SendData(SendTarget.toIndex, Origen, 0, "INITCOMUSU")
    UserList(Origen).flags.Comerciando = True

    'Actualiza el inventario del usuario
    Call UpdateUserInv(True, Destino, 0)
    'Decirle al origen que abra la ventanita.
    Call SendData(SendTarget.toIndex, Destino, 0, "INITCOMUSU")
    UserList(Destino).flags.Comerciando = True

    'Call EnviarObjetoTransaccion(Origen)
Else
    'Es el primero que comercia ?
    Call SendData(SendTarget.toIndex, Destino, 0, "||" & UserList(Origen).name & " desea comerciar. Si deseas aceptar, Escribe /COMERCIAR." & FONTTYPE_TALK)
    UserList(Destino).flags.TargetUser = Origen
    
End If

Exit Sub
errhandler:
    Call LogError("Error en IniciarComercioConUsuario: " & Err.Description)
End Sub

'envia a AQuien el objeto del otro
Public Sub EnviarObjetoTransaccion(ByVal AQuien As Integer)
Dim ObjInd As Integer
Dim ObjCant As Long

'[Alejo]: En esta funcion se centralizaba el problema
'         de no poder comerciar con mas de 32k de oro.
'         Ahora si funciona!!!

ObjCant = UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Cant
If UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto = FLAGORO Then
    ObjInd = iORO
Else
    ObjInd = UserList(UserList(AQuien).ComUsu.DestUsu).Invent.Object(UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto).ObjIndex
End If

If ObjCant <= 0 Or ObjInd <= 0 Then Exit Sub

If ObjInd > 0 And ObjCant > 0 Then
    Call SendData(SendTarget.toIndex, AQuien, 0, "COMUSUPET" & 1 & "," & ObjInd & "," & ObjData(ObjInd).name & "," & ObjCant & "," & 0 & "," & ObjData(ObjInd).GrhIndex & "," _
    & ObjData(ObjInd).OBJType & "," _
    & ObjData(ObjInd).MaxHIT & "," _
    & ObjData(ObjInd).MinHIT & "," _
    & ObjData(ObjInd).MaxDef & "," _
    & ObjData(ObjInd).Valor \ 3)
End If

End Sub

Public Sub FinComerciarUsu(ByVal userindex As Integer)
With UserList(userindex)
    If .ComUsu.DestUsu > 0 Then
        Call SendData(SendTarget.toIndex, userindex, 0, "FINCOMUSUOK")
    End If
    
    .ComUsu.Acepto = False
    .ComUsu.Cant = 0
    .ComUsu.DestUsu = 0
    .ComUsu.Objeto = 0
    .ComUsu.DestNick = ""
    .flags.Comerciando = False
End With

End Sub

Public Sub AceptarComercioUsu(ByVal userindex As Integer)
Dim Obj1 As Obj, Obj2 As Obj
Dim OtroUserIndex As Integer
Dim TerminarAhora As Boolean

TerminarAhora = False

If UserList(userindex).ComUsu.DestUsu <= 0 Then
    TerminarAhora = True
End If

OtroUserIndex = UserList(userindex).ComUsu.DestUsu


If UserList(OtroUserIndex).flags.UserLogged = False Or UserList(userindex).flags.UserLogged = False Then
    TerminarAhora = True
End If
If UserList(OtroUserIndex).ComUsu.DestUsu <> userindex Then
    TerminarAhora = True
End If
If UserList(OtroUserIndex).name <> UserList(userindex).ComUsu.DestNick Then
    TerminarAhora = True
End If
If UserList(userindex).name <> UserList(OtroUserIndex).ComUsu.DestNick Then
    TerminarAhora = True
End If

If TerminarAhora = True Then
    Call FinComerciarUsu(userindex)
    Call FinComerciarUsu(OtroUserIndex)
    Exit Sub
End If

UserList(userindex).ComUsu.Acepto = True
TerminarAhora = False

If UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.Acepto = False Then
    Call SendData(SendTarget.toIndex, userindex, 0, "||El otro usuario aun no ha aceptado tu oferta." & FONTTYPE_TALK)
    Exit Sub
End If

If UserList(userindex).ComUsu.Objeto = FLAGORO Then
    Obj1.ObjIndex = iORO
    If UserList(userindex).ComUsu.Cant > UserList(userindex).Stats.GLD Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
        TerminarAhora = True
    End If
Else
    Obj1.Amount = UserList(userindex).ComUsu.Cant
    Obj1.ObjIndex = UserList(userindex).Invent.Object(UserList(userindex).ComUsu.Objeto).ObjIndex
    If Obj1.Amount > UserList(userindex).Invent.Object(UserList(userindex).ComUsu.Objeto).Amount Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
        TerminarAhora = True
    End If
End If
If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
    Obj2.ObjIndex = iORO
    If UserList(OtroUserIndex).ComUsu.Cant > UserList(OtroUserIndex).Stats.GLD Then
        Call SendData(SendTarget.toIndex, OtroUserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
        TerminarAhora = True
    End If
Else
    Obj2.Amount = UserList(OtroUserIndex).ComUsu.Cant
    Obj2.ObjIndex = UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).ObjIndex
    If Obj2.Amount > UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).Amount Then
        Call SendData(SendTarget.toIndex, OtroUserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
        TerminarAhora = True
    End If
End If

'Por si las moscas...
If TerminarAhora = True Then
    Call FinComerciarUsu(userindex)
    Call FinComerciarUsu(OtroUserIndex)
    Exit Sub
End If

'[CORREGIDO]
'Desde acá corregí el bug que cuando se ofrecian mas de
'10k de oro no le llegaban al destinatario.

'pone el oro directamente en la billetera
If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
    'quito la cantidad de oro ofrecida
    UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD - UserList(OtroUserIndex).ComUsu.Cant
    If UserList(OtroUserIndex).ComUsu.Cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(Date & " " & UserList(OtroUserIndex).name & " solto oro en comercio seguro con " & UserList(userindex).name & ". Cantidad: " & UserList(OtroUserIndex).ComUsu.Cant)
    Call EnviarOro(OtroUserIndex)
    'y se la doy al otro
    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + UserList(OtroUserIndex).ComUsu.Cant
    If UserList(OtroUserIndex).ComUsu.Cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(Date & " " & UserList(userindex).name & " recibio oro en comercio seguro con " & UserList(OtroUserIndex).name & ". Cantidad: " & UserList(OtroUserIndex).ComUsu.Cant)
    Call EnviarOro(userindex)
Else
    'Quita el objeto y se lo da al otro
    If MeterItemEnInventario(userindex, Obj2) = False Then
        Call TirarItemAlPiso(UserList(userindex).pos, Obj2)
    End If
    Call QuitarObjetos(Obj2.ObjIndex, Obj2.Amount, OtroUserIndex)
End If

'pone el oro directamente en la billetera
If UserList(userindex).ComUsu.Objeto = FLAGORO Then
    'quito la cantidad de oro ofrecida
    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - UserList(userindex).ComUsu.Cant
    If UserList(userindex).ComUsu.Cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(Date & " " & UserList(userindex).name & " solto oro en comercio seguro con " & UserList(OtroUserIndex).name & ". Cantidad: " & UserList(userindex).ComUsu.Cant)
    Call EnviarOro(userindex)
    'y se la doy al otro
    UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD + UserList(userindex).ComUsu.Cant
    If UserList(userindex).ComUsu.Cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(Date & " " & UserList(OtroUserIndex).name & " recibio oro en comercio seguro con " & UserList(userindex).name & ". Cantidad: " & UserList(userindex).ComUsu.Cant)
    Call EnviarOro(OtroUserIndex)
Else
    'Quita el objeto y se lo da al otro
    If MeterItemEnInventario(OtroUserIndex, Obj1) = False Then
        Call TirarItemAlPiso(UserList(OtroUserIndex).pos, Obj1)
    End If
    Call QuitarObjetos(Obj1.ObjIndex, Obj1.Amount, userindex)
End If

'[/CORREGIDO] :p

Call UpdateUserInv(True, userindex, 0)
Call UpdateUserInv(True, OtroUserIndex, 0)

Call FinComerciarUsu(userindex)
Call FinComerciarUsu(OtroUserIndex)
 
End Sub

'[/Alejo]

