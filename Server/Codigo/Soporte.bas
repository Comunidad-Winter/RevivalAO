Attribute VB_Name = "Soporte"
Public xaoindex As Integer
Public Sub MostrarSop(ByVal userindex As Integer, ByVal marika As Integer, ByVal nombre As String)
xaoindex = NameIndex(nombre)
 If xaoindex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El usuario se encuentra offline." & FONTTYPE_INFO)
Exit Sub
Else
SendData SendTarget.toindex, userindex, 0, "SOPO" & _
            UserList(marika).Pregunta _
             & Chr$(2) & UserList(marika).name
End If
End Sub
Public Sub EnviaRespuesta(ByVal elmarika As String)
xaoindex = NameIndex(elmarika)
If Not xaoindex <= 0 Then
SendData SendTarget.toindex, xaoindex, 0, "LLE"
 Call SendData(SendTarget.toindex, xaoindex, 0, "TW126")
End If
End Sub
Public Sub EnviarResp(ByVal userindex As Integer)
SendData SendTarget.toindex, userindex, 0, "RESP" & _
UserList(userindex).Respuesta
End Sub
Public Sub ResetSop(ByVal userindex As Integer)
UserList(userindex).Pregunta = "Ninguna"
UserList(userindex).Respuesta = "Ninguna"
UserList(userindex).flags.Soporteo = False
End Sub

