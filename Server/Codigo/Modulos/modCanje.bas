Attribute VB_Name = "modCanje"
Option Explicit
Public pathCanje As String
Public Sub EnviarCanje(ByVal userindex As Integer)
   Dim i As Variant
   For i = 1 To GetVar(pathCanje, "CANTIDAD", "CANTIDAD")
   Debug.Print GetVar(pathCanje, "CANJE" & i, "NOMBRE")
     Call SendData(SendTarget.toIndex, userindex, 0, "TUKI" & GetVar(pathCanje, "CANJE" & i, "NOMBRE") & "," & GetVar(pathCanje, "CANJE" & i, "MIN") & "," & GetVar(pathCanje, "CANJE" & i, "MAX") & "," & GetVar(pathCanje, "CANJE" & i, "VALOR") & "," & GetVar(pathCanje, "CANJE" & i, "GRHINDEX") & "," & UserList(userindex).Stats.PuntosCanje) '& "," & GetVar(pathCanje, "CANJE" & i, "NUMERO")
    Next i
    Call SendData(SendTarget.toIndex, userindex, 0, "INIC")
   
End Sub

