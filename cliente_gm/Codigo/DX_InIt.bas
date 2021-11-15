Attribute VB_Name = "Mod_DX"
Option Explicit

Public DirectX As New DirectX7
Public DirectDraw As DirectDraw7

Public PrimarySurface As DirectDrawSurface7
Public PrimaryClipper As DirectDrawClipper
Public SecundaryClipper As DirectDrawClipper
Public BackBufferSurface As DirectDrawSurface7

Public SupBMiniMap As DirectDrawSurface7
Public SupMiniMap As DirectDrawSurface7

Public oldResHeight As Long, oldResWidth As Long
Public bNoResChange As Boolean

Private Sub IniciarDXobject(DX As DirectX7)

Err.Clear

On Error Resume Next

Set DX = New DirectX7

If Err Then
    MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
    LogError "Error producido por Set DX = New DirectX7"
    End
End If

End Sub

Private Sub IniciarDDobject(DD As DirectDraw7)
Err.Clear
On Error Resume Next
Set DD = DirectX.DirectDrawCreate("")
If Err Then
    MsgBox "No se puede iniciar DirectDraw. Por favor asegurese de tener la ultima version correctamente instalada."
    LogError "Error producido en Private Sub IniciarDDobject(DD As DirectDraw7)"
    End
End If
End Sub

Public Sub IniciarObjetosDirectX()

On Error Resume Next

Call AddtoRichTextBox(frmCargando.status, "Iniciando DirectX....", 0, 0, 0, 0, 0, True)
Call IniciarDXobject(DirectX)
Call AddtoRichTextBox(frmCargando.status, "Hecho", , , , 1, , False)

Call AddtoRichTextBox(frmCargando.status, "Iniciando DirectDraw....", 0, 0, 0, 0, 0, True)
Call IniciarDDobject(DirectDraw)
Call AddtoRichTextBox(frmCargando.status, "Hecho", , , , 1, , False)

    Call AddtoRichTextBox(frmCargando.status, "Iniciando DirectSound....", 0, 0, 0, 0, 0, True)
    Call AddtoRichTextBox(frmCargando.status, "Hecho", , , , 1, , False)

Call AddtoRichTextBox(frmCargando.status, "Analizando y preparando la placa de video....", 0, 0, 0, 0, 0, True)
    
Dim lRes As Long
Dim MidevM As typDevMODE
lRes = EnumDisplaySettings(0, 0, MidevM)
    
Dim intWidth As Integer
Dim intHeight As Integer

oldResWidth = Screen.Width \ Screen.TwipsPerPixelX
oldResHeight = Screen.Height \ Screen.TwipsPerPixelY

If oldResWidth <> 800 Or oldResHeight <> 600 Then
    If MsgBox("¿Quieres jugar en modo Pantalla Completa?", vbYesNo, "Resolución") = vbYes Then
        bNoResChange = True
        frmMain.Height = 12000
        With MidevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
            .dmPelsWidth = 800
            .dmPelsHeight = 600
'            .dmBitsPerPel = 16
        End With
       lRes = ChangeDisplaySettings(MidevM, CDS_TEST)
    Else
frmMain.Height = 9000
        bNoResChange = False
    End If
End If
 

Call AddtoRichTextBox(frmCargando.status, "¡DirectX OK!", 0, 251, 0, 1, 0)

Exit Sub

End Sub

Public Sub LiberarObjetosDX()
Err.Clear
On Error GoTo fin:
Dim loopc As Integer

Set PrimarySurface = Nothing
Set PrimaryClipper = Nothing
Set BackBufferSurface = Nothing

Set DirectDraw = Nothing

Set DirectX = Nothing
Exit Sub
fin: LogError "Error producido en Public Sub LiberarObjetosDX()"
End Sub

