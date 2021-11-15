Attribute VB_Name = "Existess"
Function Existe(sArchivo As String) As Boolean
    Existe = Len(Dir$(sArchivo))
End Function

