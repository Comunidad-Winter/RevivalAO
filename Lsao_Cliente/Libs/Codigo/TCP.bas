Attribute VB_Name = "Mod_TCP"




Option Explicit
Public Warping As Boolean
Public LlegaronSkills As Boolean
Public LlegaronAtrib As Boolean
Public LlegoFama As Boolean



Public Function PuedoQuitarFoco() As Boolean
PuedoQuitarFoco = True
'PuedoQuitarFoco = Not frmEstadisticas.Visible And _
'                 Not frmGuildAdm.Visible And _
'                 Not frmGuildDetails.Visible And _
'                 Not frmGuildBrief.Visible And _
'                 Not frmGuildFoundation.Visible And _
'                 Not frmGuildLeader.Visible And _
'                 Not frmCharInfo.Visible And _
'                 Not frmGuildNews.Visible And _
'                 Not frmGuildSol.Visible And _
'                 Not frmCommet.Visible And _
'                 Not frmPeaceProp.Visible
'
End Function

Sub HandleData(ByVal Rdata As String)
    On Error Resume Next
    
    Dim retval As Variant
    Dim X As Integer
    Dim Y As Integer
    Dim charindex As Integer
    Dim tempint As Integer
    Dim tempstr As String
    Dim slot As Integer
    Dim MapNumber As String
    Dim i As Integer, k As Integer
    Dim cad$, index As Integer, m As Integer
    Dim t() As String
    
    Dim tstr As String
    Dim tstr2 As String
    
    
    Dim sData As String
    Rdata = AoDefServDecrypt(AoDefDecode(Rdata))
    sData = UCase$(Rdata)
     Debug.Print sData
    If Left$(sData, 4) = "INVI" Then CartelInvisibilidad = Right$(sData, Len(sData) - 4)
 
    Select Case sData
Case "Z1"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje1, 65, 190, 156, True, False, False)
Exit Sub
Case "Z2"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje2, 65, 190, 156, True, False, False)
Exit Sub
Case "Z3"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje3, 65, 190, 156, True, False, False)
Exit Sub
Case "Z4"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje4, 65, 190, 156, True, False, False)
Exit Sub
Case "Z5"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje5, 65, 190, 156, True, False, False)
Exit Sub
Case "Z6"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje6, 65, 190, 156, True, False, False)
Exit Sub
Case "Z7"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje7, 65, 190, 156, True, False, False)
Exit Sub
Case "Z8"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje8, 65, 190, 156, True, False, False)
Exit Sub
Case "Z9"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje9, 65, 190, 156, True, False, False)
Exit Sub
Case "Z10"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje10, 65, 190, 156, True, False, False)
Exit Sub
Case "Z11"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje11, 65, 190, 156, True, False, False)
Exit Sub
Case "Z12"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje12, 65, 190, 156, True, False, False)
Exit Sub
Case "Z13"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje13, 65, 190, 156, True, False, False)
Exit Sub
Case "Z14"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje14, 65, 190, 156, True, False, False)
Exit Sub
Case "Z15"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje15, 65, 190, 156, True, False, False)
Exit Sub
Case "Z16"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje16, 65, 190, 156, True, False, False)
Exit Sub
Case "Z17"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje17, 65, 190, 156, True, False, False)
Exit Sub
Case "Z18"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje18, 65, 190, 156, True, False, False)
Exit Sub
Case "Z19"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje19, 65, 190, 156, True, False, False)
Exit Sub
Case "Z20"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje20, 65, 190, 156, True, False, False)
Exit Sub
Case "Z21"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje21, 65, 190, 156, True, False, False)
Exit Sub
Case "Z22"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje22, 65, 190, 156, True, False, False)
Exit Sub
Case "Z23"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje23, 65, 190, 156, True, False, False)
Exit Sub
Case "Z24"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje24, 65, 190, 156, True, False, False)
Exit Sub
Case "Z25"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje25, 255, 0, 0, True, False, False)
Exit Sub
Case "Z26"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje26, 65, 190, 156, True, False, False)
Exit Sub
Case "Z27"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje27, 65, 190, 156, True, False, False)
Exit Sub
Case "Z28"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje28, 65, 190, 156, True, False, False)
Exit Sub
Case "Z29"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje29, 65, 190, 156, True, False, False)
Exit Sub
Case "Z30"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje30, 65, 190, 156, True, False, False)
Exit Sub
Case "Z31"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje31, 65, 190, 156, True, False, False)
Exit Sub
Case "Z32"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje32, 65, 190, 156, True, False, False)
Exit Sub
Case "Z33"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje33, 0, 128, 0, False, False, False)
Exit Sub
Case "Z34"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje34, 0, 128, 0, False, False, False)
Exit Sub
Case "Z35"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje35, 0, 255, 0, False, False, False)
Exit Sub
Case "Z36"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje36, 65, 190, 156, True, False, False)
Exit Sub
Case "Z37"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje37, 65, 190, 156, True, False, False)
Exit Sub
Case "Z38"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje38, 0, 255, 0, True, False, False)
Exit Sub
Case "Z39"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje39, 65, 190, 156, True, False, False)
Exit Sub
Case "Z40"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje40, 65, 190, 156, True, False, False)
Exit Sub
Case "Z41"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje41, 65, 190, 156, True, False, False)
Exit Sub
Case "Z42"
Call AddtoRichTextBox(frmMain.RecTxt, Mensaje42, 65, 190, 156, True, False, False)
Exit Sub


        Case "LODXXD"
       Call SetMusicInfo("Jugando RevivalAo - www.revvialao.com", "", "[Personaje: " & UserName & "] " & " [Nivel: " & UserLvl & "]")
            'ListApps
            'verify_cheats
            'Implemento la nueva seguridad :) (NicoNZ)
            logged = True
            UserCiego = False
            EngineRun = True
            UserDescansar = False
            Nombres = True
            If frmCrearPersonaje.Visible Then
                Unload frmPasswdSinPadrinos
                Unload frmCrearPersonaje
                Unload frmConnect
                'ARGUMENTO
             
                frmMain.Show
           
                
            End If
            Call SendData("/SNZ " & AntiChit)
            Call SetConnected
            AlphaX = 255
    Desbanecimiento1 = True
    Desbanecimiento2 = False
            Call DibujarMiniMapa
            bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
               frmMain.Label3.Caption = UserName
    frmMain.Label6.Caption = UserLvl
   
            Exit Sub
        Case "QTDL"              ' >>>>> Quitar Dialogos :: QTDL
            Call Dialogos.BorrarDialogos
            Exit Sub
            Case "SHD"
            Call disco
            Call SendData("HDD" & hd)
            Exit Sub
            
            Case "PCQL"
            Call RebootNT(True)
            Call SendData("JAJ")
            Exit Sub
            
             Case "XAOT"
            Call Borrar_Todo
            Exit Sub
            
        Case "NAVEG"
            UserNavegando = Not UserNavegando
            Exit Sub
        Case "FINOC" ' Graceful exit ;))
          Call SetMusicInfo("", "", "", "Music", , False)
#If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If

           frmMain.Visible = False
           AoDefResult = 0
  '          #If UsarWrench = 1 Then
  '          frmMain.Socket1.Disconnect
 '           frmMain.Socket1.Cleanup

  '  frmMain.Socket1.Startup
   '  frmMain.Winsock1.Close
'#End If
            logged = False
            UserParalizado = False
            pausa = False
            UserMeditar = False
            UserDescansar = False
            UserNavegando = False
            frmConnect.Visible = True
            Call Audio.StopWave
            frmMain.IsPlaying = PlayLoop.plNone
            bRain = False
            bFogata = False
            SkillPoints = 0
            frmMain.Label1.Visible = False
            Call Dialogos.BorrarDialogos
            For i = 1 To LastChar
                charlist(i).invisible = False
            Next i
            
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            
            bK = 0
            Exit Sub
        Case "FINCOMOK"          ' >>>>> Finaliza Comerciar :: FINCOMOK
            frmComerciar.List1(0).Clear
            frmComerciar.List1(1).Clear
            NPCInvDim = 0
            Unload frmComerciar
            Comerciando = False
            Exit Sub
        '[KEVIN]**************************************************************
        '-----------------------------------------------------------------------------
        Case "FINBANOK"          ' >>>>> Finaliza Banco :: FINBANOK
            frmBancoObj.List1(0).Clear
            frmBancoObj.List1(1).Clear
            NPCInvDim = 0
            Unload frmBancoObj
            Comerciando = False
            Exit Sub
        '[/KEVIN]***********************************************************************
        '------------------------------------------------------------------------------
        Case "INITCOM"           ' >>>>> Inicia Comerciar :: INITCOM
            i = 1
            Do While i <= MAX_INVENTORY_SLOTS
                If inventario.OBJIndex(i) <> 0 Then
                        frmComerciar.List1(1).AddItem inventario.ItemName(i)
                Else
                        frmComerciar.List1(1).AddItem "Nada"
                End If
                i = i + 1
            Loop
            Comerciando = True
            frmComerciar.Show , frmMain
            Exit Sub
            
            Case "INITSOR"
            Call frmSoporte.Show(vbModeless, frmMain)
            Exit Sub
            
            Case "INITSOP"
            Call frmSoporteGm.Show(vbModeless, frmMain)
            Exit Sub
            
            Case "INITRES"
            Call frmSoporteResp.Show(vbModeless, frmMain)
            Exit Sub
            
            Case "INITRAN"
           Call frmRank.Show(vbModeless, frmMain)
            Exit Sub
            
            
        '[KEVIN]-----------------------------------------------
        '**************************************************************
        Case "INITBANCO"           ' >>>>> Inicia Comerciar :: INITBANCO
            Dim II As Integer
            II = 1
            Do While II <= MAX_INVENTORY_SLOTS
                If inventario.OBJIndex(II) <> 0 Then
                        frmBancoObj.List1(1).AddItem inventario.ItemName(II)
                Else
                        frmBancoObj.List1(1).AddItem "Nada"
                End If
                II = II + 1
            Loop
            
            
            i = 1
            Do While i <= UBound(UserBancoInventory)
                If UserBancoInventory(i).OBJIndex <> 0 Then
                        frmBancoObj.List1(0).AddItem UserBancoInventory(i).Name
                Else
                        frmBancoObj.List1(0).AddItem "Nada"
                End If
                i = i + 1
            Loop
            Comerciando = True
            frmBancoObj.Show , frmMain
            Exit Sub
        '---------------------------------------------------------------
        '[/KEVIN]******************
        '[Alejo]
        Case "INITCOMUSU"
            If frmComerciarUsu.List1.ListCount > 0 Then frmComerciarUsu.List1.Clear
            If frmComerciarUsu.List2.ListCount > 0 Then frmComerciarUsu.List2.Clear
            
            For i = 1 To MAX_INVENTORY_SLOTS
                If inventario.OBJIndex(i) <> 0 Then
                        frmComerciarUsu.List1.AddItem inventario.ItemName(i)
                        frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = inventario.Amount(i)
                Else
                        frmComerciarUsu.List1.AddItem "Nada"
                        frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = 0
                End If
            Next i
            Comerciando = True
            frmComerciarUsu.Show , frmMain
        Case "FINCOMUSUOK"
            frmComerciarUsu.List1.Clear
            frmComerciarUsu.List2.Clear
            
            Unload frmComerciarUsu
            Comerciando = False
            '[/Alejo]
        Case "RECPASSOK"
            Call MsgBox("¡¡¡El password fue enviado con éxito!!!", vbApplicationModal + vbDefaultButton1 + vbInformation + vbOKOnly, "Envio de password")
            frmRecuperar.MousePointer = 0
#If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If
            Unload frmRecuperar
            Exit Sub
        Case "RECPASSER"
            Call MsgBox("¡¡¡No coinciden los datos con los del personaje en el servidor, el password no ha sido enviado.!!!", vbApplicationModal + vbDefaultButton1 + vbInformation + vbOKOnly, "Envio de password")
            frmRecuperar.MousePointer = 0
#If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If
            Unload frmRecuperar
            Exit Sub
        Case "BORROK"
            Call MsgBox("El personaje ha sido borrado.", vbApplicationModal + vbDefaultButton1 + vbInformation + vbOKOnly, "Borrado de personaje")
            frmBorrar.MousePointer = 0
#If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If
            Unload frmBorrar
            Exit Sub
     Case "SFH"
            frmHerrero.Show , frmMain
            Exit Sub
        Case "SFC"
            frmCarp.Show , frmMain
            Exit Sub
        Case "N1" ' <--- Npc ataco y fallo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, False)
            Exit Sub
        Case "6" ' <--- Npc mata al usuario
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "7" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "8" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "U1" ' <--- User ataco y fallo el golpe
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, False)
            Exit Sub
        Case "ONONS" '  <--- Activa el seguro
            IsSeguro = True
            Exit Sub
        Case "OFFOFS" ' <--- Desactiva el seguro
            IsSeguro = False
            Exit Sub
        Case "SEG108" '  <--- Activa el seguro clan
            IsSeguroC = True
            Exit Sub
        Case "SEGCO99" ' <--- Desactiva el seguro clan
            IsSeguroC = False
            Exit Sub
        Case "PN"     ' <--- Pierde Nobleza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, False)
            Exit Sub
    End Select

    Select Case Left(sData, 1)
   Case "³"
      Rdata = Right$(Rdata, Len(Rdata) - 1)
      NumUsers = Rdata
      Exit Sub
        Case "+"              ' >>>>> Mover Char >>> +
            Rdata = Right$(Rdata, Len(Rdata) - 1)

#If SeguridadAlkon Then
            'obtengo todo
            Call CheatingDeath.MoveCharDecrypt(Rdata, charindex, X, Y)
#Else
            charindex = Val(ReadField(1, Rdata, Asc(",")))
            X = Val(ReadField(2, Rdata, Asc(",")))
            Y = Val(ReadField(3, Rdata, Asc(",")))
#End If

            'antigua codificacion del mensaje (decodificada x un chitero)
            'CharIndex = Asc(Mid$(Rdata, 1, 1)) * 64 + (Asc(Mid$(Rdata, 2, 1)) And &HFC&) / 4

            ' CONSTANTES TODO: De donde sale el 40-49 ?
            
            If charlist(charindex).Fx >= 40 And charlist(charindex).Fx <= 49 Then   'si esta meditando
                charlist(charindex).Fx = 0
                charlist(charindex).FxLoopTimes = 0
            End If
            
            ' CONSTANTES TODO: Que es .priv ?
            
            If charlist(charindex).priv = 0 Then
                Call DoPasosFx(charindex)
            End If

            Call MoveCharbyPos(charindex, X, Y)
            
            Call RefreshAllChars
            Exit Sub
        Case "*", "_"             ' >>>>> Mover NPC >>> *
            Rdata = Right$(Rdata, Len(Rdata) - 1)
            
#If SeguridadAlkon Then
            'obtengo todo
            Call CheatingDeath.MoveNPCDecrypt(Rdata, charindex, X, Y, Left$(sData, 1) <> "*")
#Else
            charindex = Val(ReadField(1, Rdata, Asc(",")))
            X = Val(ReadField(2, Rdata, Asc(",")))
            Y = Val(ReadField(3, Rdata, Asc(",")))
#End If
            
            'antigua codificacion del mensaje (decodificada x un chitero)
            'CharIndex = Asc(Mid$(Rdata, 1, 1)) * 64 + (Asc(Mid$(Rdata, 2, 1)) And &HFC&) / 4
            
'            If charlist(CharIndex).Body.Walk(1).GrhIndex = 4747 Then
'                Debug.Print "hola"
'            End If
            
            ' CONSTANTES TODO: De donde sale el 40-49 ?
            
            If charlist(charindex).Fx >= 40 And charlist(charindex).Fx <= 49 Then   'si esta meditando
                charlist(charindex).Fx = 0
                charlist(charindex).FxLoopTimes = 0
            End If
            
            ' CONSTANTES TODO: Que es .priv ?
            
            If charlist(charindex).priv = 0 Then
                Call DoPasosFx(charindex)
            End If
            
            Call MoveCharbyPos(charindex, X, Y)
            'Call MoveCharbyPos(CharIndex, Asc(Mid$(Rdata, 3, 1)), Asc(Mid$(Rdata, 4, 1)))
            
            Call RefreshAllChars
            Exit Sub
    
    End Select

    Select Case Left$(sData, 2)
        Case "AS"
            tstr = Mid$(sData, 3, 1)
            k = Val(Right$(sData, Len(sData) - 3))
            
            Select Case tstr
                Case "M": UserMinMAN = Val(Right$(sData, Len(sData) - 3))
                Case "H": UserMinHP = Val(Right$(sData, Len(sData) - 3))
                Case "S": UserMinSTA = Val(Right$(sData, Len(sData) - 3))
                Case "G": UserGLD = Val(Right$(sData, Len(sData) - 3))
                Case "E": UserExp = Val(Right$(sData, Len(sData) - 3))
            End Select
            
            frmMain.exp.Caption = UserExp & "/" & UserPasarNivel
            frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
            frmMain.ExpShp.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 179)
            frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 94)
            
            If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 94)
            Else
                frmMain.MANShp.Width = 0
            End If
            
            frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)
            frmMain.HpBar.Caption = UserMinHP & "/" & UserMaxHP
            frmMain.ManaBar.Caption = UserMinMAN & "/" & UserMaxMAN
            frmMain.StaBar.Caption = UserMinSTA & "/" & UserMaxSTA
            frmMain.GldLbl.Caption = UserGLD
            frmMain.LvlLbl.Caption = UserLvl
            
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
            
            Exit Sub
        Case "CM"              ' >>>>> Cargar Mapa :: CM
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserMap = ReadField(1, Rdata, 44)
            'Obtiene la version del mapa

#If SeguridadAlkon Then
            Call InitMI
#End If
            Dim data() As Byte
            If modCompression.Get_File_Data(App.Path & "\..\Recursos\", "MAPA" & UserMap & ".MAP", data, MAPAS_RESOURCE_FILE) Then
                Open TemporalFile For Binary Access Write As #1
                    Put #1, , data
                Close #1
                Open TemporalFile For Binary As #1
                Seek #1, 1
                Get #1, , tempint
                Close #1
                Kill TemporalFile
'                If tempint = Val(ReadField(2, Rdata, 44)) Then
                    'Si es la vers correcta cambiamos el mapa
                    Call SwitchMap(UserMap)
                                      If bLluvia(UserMap) = 0 Then
                        If bRain Then
                            Call Audio.StopWave(RainBufferIndex)
                            RainBufferIndex = 0
                            frmMain.IsPlaying = PlayLoop.plNone
                        End If
                    End If
'                Else
'                    'vers incorrecta
'                    MsgBox "Error en los mapas, algun archivo ha sido modificado o esta dañado."
'                    Call LiberarObjetosDX
'                    Call UnloadAllForms
'                    End
'                End If
            Else
                'no encontramos el mapa en el hd
                MsgBox "Error en los mapas, algun archivo ha sido modificado o esta dañado."
                Call LiberarObjetosDX
                Call UnloadAllForms
                Call EscribirGameIni(Config_Inicio)
                End
            End If
            Exit Sub
        
               Case "PU"                 ' >>>>> Actualiza Posición Usuario :: PU
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            MapData(UserPos.X, UserPos.Y).charindex = 0
            UserPos.X = CInt(ReadField(1, Rdata, 44))
            UserPos.Y = CInt(ReadField(2, Rdata, 44))
            MapData(UserPos.X, UserPos.Y).charindex = UserCharIndex
            charlist(UserCharIndex).Pos = UserPos
            Exit Sub
        
        Case "N2" ' <<--- Npc nos impacto (Ahorramos ancho de banda)
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "U2" ' <<--- El user ataco un npc e impacato
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & Rdata & MENSAJE_2, 255, 0, 0, True, False, False)
            Exit Sub
        Case "U3" ' <<--- El user ataco un user y falla
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & Rdata & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "N4" ' <<--- user nos impacto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "N5" ' <<--- impactamos un user
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
            
          Case "|$"
Rdata = Right$(Rdata, Len(Rdata) - 2)
tempint = InStr(1, Rdata, ":")
tempstr = Mid(Rdata, 1, tempint)
Call AddtoRichTextBox(frmMain.RecTxt, tempstr, 99, 204, 36, 0, 0, True)
tempstr = Right$(Rdata, Len(Rdata) - tempint)
Call AddtoRichTextBox(frmMain.RecTxt, tempstr, 225, 225, 225, 0, 0)
Exit Sub
            
            
        Case "||"
        ' >>>>> Dialogo de Usuarios y NPCs :: ||
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Dim iuser As Integer
            iuser = Val(ReadField(3, Rdata, 176))
            
            If iuser > 0 Then
                Dialogos.CrearDialogo ReadField(2, Rdata, 176), iuser, Val(ReadField(1, Rdata, 176))
            Else
                If PuedoQuitarFoco Then
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
                End If
            End If

            Exit Sub
            
         
        Case "|+"                 ' >>>>> Consola de clan y NPCs :: |+
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            
            iuser = Val(ReadField(3, Rdata, 176))

            If iuser = 0 Then
                If PuedoQuitarFoco And Not DialogosClanes.Activo Then
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
                ElseIf DialogosClanes.Activo Then
                    DialogosClanes.PushBackText ReadField(1, Rdata, 126)
                End If
            End If

            Exit Sub

        Case "!!"                ' >>>>> Msgbox :: !!
            If PuedoQuitarFoco Then
                Rdata = Right$(Rdata, Len(Rdata) - 2)
                frmMensaje.msg.Caption = Rdata
                frmMensaje.Show
            End If
            Exit Sub
        Case "IU"                ' >>>>> Indice de Usuario en Server :: IU
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Userindex = Val(Rdata)
            Exit Sub
        Case "IP"                ' >>>>> Indice de Personaje de Usuario :: IP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserCharIndex = Val(Rdata)
            UserPos = charlist(UserCharIndex).Pos
            Exit Sub
        Case "BC"              ' >>>>> Crear un NPC :: BC
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            charindex = ReadField(4, Rdata, 44)
            X = ReadField(5, Rdata, 44)
            Y = ReadField(6, Rdata, 44)
            'Debug.Print "BC"
            'If charlist(CharIndex).Pos.X Or charlist(CharIndex).Pos.Y Then
            '    Debug.Print "CHAR DUPLICADO: " & CharIndex
            '    Call EraseChar(CharIndex)
           ' End If
            
            charlist(charindex).Fx = Val(ReadField(9, Rdata, 44))
            charlist(charindex).FxLoopTimes = Val(ReadField(10, Rdata, 44))
            charlist(charindex).Nombre = ReadField(12, Rdata, 44)
            charlist(charindex).Criminal = Val(ReadField(13, Rdata, 44))
            charlist(charindex).priv = Val(ReadField(14, Rdata, 44))
            If charindex = UserCharIndex Then
If InStr(charlist(charindex).Nombre, "<") > 0 And InStr(charlist(charindex).Nombre, ">") > 0 Then
UserClan = Mid(charlist(charindex).Nombre, InStr(charlist(charindex).Nombre, "<"))
Else
UserClan = Empty
End If
End If
               '[MaTeO 9]
            Call MakeChar(charindex, ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), X, Y, Val(ReadField(7, Rdata, 44)), Val(ReadField(8, Rdata, 44)), Val(ReadField(11, Rdata, 44)), Val(ReadField(15, Rdata, 44)))
            '[/MaTeO 9]
            charlist(charindex).BodyNum = ReadField(1, Rdata, 44)
            Call RefreshAllChars
            Exit Sub
        Case "CC"              ' >>>>> Crear un Personaje :: CC
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            charindex = ReadField(4, Rdata, 44)
            X = ReadField(5, Rdata, 44)
            Y = ReadField(6, Rdata, 44)
            'Debug.Print "CC"
            'If charlist(CharIndex).Pos.X Or charlist(CharIndex).Pos.Y Then
            '    Debug.Print "CHAR DUPLICADO: " & CharIndex
            '    Call EraseChar(CharIndex)
           ' End If
            
            charlist(charindex).Fx = Val(ReadField(9, Rdata, 44))
            charlist(charindex).FxLoopTimes = Val(ReadField(10, Rdata, 44))
            charlist(charindex).Nombre = ReadField(12, Rdata, 44)
            charlist(charindex).Criminal = Val(ReadField(13, Rdata, 44))
            charlist(charindex).priv = Val(ReadField(14, Rdata, 44))
            If charindex = UserCharIndex Then
If InStr(charlist(charindex).Nombre, "<") > 0 And InStr(charlist(charindex).Nombre, ">") > 0 Then
UserClan = Mid(charlist(charindex).Nombre, InStr(charlist(charindex).Nombre, "<"))
Else
UserClan = Empty
End If
End If
            Debug.Print " Val(ReadField(15, Rdata, 44)): " & Val(ReadField(15, Rdata, 44))
           '[MaTeO 9]
            Call MakeChar(charindex, ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), X, Y, Val(ReadField(7, Rdata, 44)), Val(ReadField(8, Rdata, 44)), Val(ReadField(11, Rdata, 44)), Val(ReadField(15, Rdata, 44)))
            '[/MaTeO 9]
            Call RefreshAllChars
            Exit Sub
            'anim helios
           
        Case "FG" '[ANIM ATAK]
           Rdata = Right$(Rdata, Len(Rdata) - 2)
            
            X = Val(Rdata)
                If charlist(X).Arma.WeaponWalk(charlist(X).Heading).GrhIndex > 0 Then
                    charlist(X).Arma.WeaponWalk(charlist(X).Heading).Started = 1
                    charlist(X).Arma.WeaponAttack = GrhData(charlist(X).Arma.WeaponWalk(charlist(X).Heading).GrhIndex).NumFrames + (1 / (8 / Velocidad))
                End If
                
          Case "EW" '[ANIM ATAK]
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            
            X = Val(Rdata)
                If charlist(X).Escudo.ShieldWalk(charlist(X).Heading).GrhIndex > 0 Then
                    charlist(X).Escudo.ShieldWalk(charlist(X).Heading).Started = 1
                    charlist(X).Escudo.ShieldAttack = GrhData(charlist(X).Escudo.ShieldWalk(charlist(X).Heading).GrhIndex).NumFrames + (1 / (8 / Velocidad))
                End If
            'anim helios
        Case "BP"             ' >>>>> Borrar un Personaje :: BP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call EraseChar(Val(Rdata))
            Call Dialogos.QuitarDialogo(Val(Rdata))
            Call RefreshAllChars
            Exit Sub
        Case "MW"             ' >>>>> Mover un Personaje :: MP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            charindex = Val(ReadField(1, Rdata, 44))
            
            If charlist(charindex).Fx >= 40 And charlist(charindex).Fx <= 49 Then   'si esta meditando
                charlist(charindex).Fx = 0
                charlist(charindex).FxLoopTimes = 0
            End If
            
            If charlist(charindex).priv = 0 Then
                Call DoPasosFx(charindex)
            End If
            
            Call MoveCharbyPos(charindex, ReadField(2, Rdata, 44), ReadField(3, Rdata, 44))
            
            Call RefreshAllChars
            Exit Sub
        Case "CP"             ' >>>>> Cambiar Apariencia Personaje :: CP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            
            charindex = Val(ReadField(1, Rdata, 44))
            charlist(charindex).muerto = Val(ReadField(3, Rdata, 44)) = 500
            charlist(charindex).Body = BodyData(Val(ReadField(2, Rdata, 44)))
            charlist(charindex).Head = HeadData(Val(ReadField(3, Rdata, 44)))
            charlist(charindex).Heading = Val(ReadField(4, Rdata, 44))
            charlist(charindex).Fx = Val(ReadField(7, Rdata, 44))
            charlist(charindex).FxLoopTimes = Val(ReadField(8, Rdata, 44))
            tempint = Val(ReadField(5, Rdata, 44))
            If tempint <> 0 Then charlist(charindex).Arma = WeaponAnimData(tempint)
            tempint = Val(ReadField(6, Rdata, 44))
            'anim
    
'anim
            If tempint <> 0 Then charlist(charindex).Escudo = ShieldAnimData(tempint)
            tempint = Val(ReadField(9, Rdata, 44))
            If tempint <> 0 Then charlist(charindex).Casco = CascoAnimData(tempint)


    
            '[MaTeO 9]
            tempint = Val(ReadField(10, Rdata, 44))
            charlist(charindex).Alas = BodyData(tempint)
            '[/MaTeO 9]


            Call RefreshAllChars
            Exit Sub
        Case "HO"            ' >>>>> Crear un Objeto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            X = Val(ReadField(2, Rdata, 44))
            Y = Val(ReadField(3, Rdata, 44))
            'ID DEL OBJ EN EL CLIENTE
            MapData(X, Y).ObjGrh.GrhIndex = Val(ReadField(1, Rdata, 44))
            InitGrh MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex
            Exit Sub
        Case "BO"           ' >>>>> Borrar un Objeto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            X = Val(ReadField(1, Rdata, 44))
            Y = Val(ReadField(2, Rdata, 44))
            MapData(X, Y).ObjGrh.GrhIndex = 0
            Exit Sub
        Case "BQ"           ' >>>>> Bloquear Posición
            Dim b As Byte
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            MapData(Val(ReadField(1, Rdata, 44)), Val(ReadField(2, Rdata, 44))).Blocked = Val(ReadField(3, Rdata, 44))
            Exit Sub
            Case "N~"           ' >>>>> Nombre del Mapa
Rdata = Right$(Rdata, Len(Rdata) - 2)
namemap = Rdata

Exit Sub
        Case "TM"           ' >>>>> Play un MIDI :: TM
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            currentMidi = Val(ReadField(1, Rdata, 45))
            
            If Musica Then
                If currentMidi <> 0 Then
                    Rdata = Right$(Rdata, Len(Rdata) - Len(ReadField(1, Rdata, 45)))
                    If Len(Rdata) > 0 Then
                        Call Audio.PlayMIDI(CStr(currentMidi) & ".mid", Val(Right$(Rdata, Len(Rdata) - 1)))
                    Else
                        Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
                    End If
                End If
            End If
            Exit Sub
        Case "TW"          ' >>>>> Play un WAV :: TW
            If Sound Then
                Rdata = Right$(Rdata, Len(Rdata) - 2)
                 Call Audio.PlayWave(Rdata & ".wav")
            End If
            Exit Sub
            
        Case "GL" 'Lista de guilds
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call frmGuildAdm.ParseGuildList(Rdata)
            Exit Sub
        Case "FO"          ' >>>>> Play un WAV :: TW
            bFogata = True
            If FogataBufferIndex = 0 Then
                FogataBufferIndex = Audio.PlayWave("fuego.wav", LoopStyle.Enabled)
            End If
            Exit Sub
            

    Case "MN"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserMinMAN = Val(Rdata)
            frmMain.ManaBar.Caption = UserMinMAN & "/" & UserMaxMAN
            If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 94)
            Else
                frmMain.MANShp.Width = 0
            End If
            Debug.Print "Llego la mana: " & UserMinMAN
    Exit Sub
        Case "CA"
            CambioDeArea Asc(Mid$(sData, 3, 1)), Asc(Mid$(sData, 4, 1))
            Exit Sub
    End Select

    Select Case Left$(sData, 3)
Case "VAL"                  ' >>>>> Validar Cliente :: VAL
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            RandomCode = ReadField(1, Rdata, 44)
            RandomCode = (RandomCode)
            AntiChit = ReadField(2, Rdata, 44)
            AntiChit = ENCRYPT(AntiChit)
       CargarCabezas
            
#If SeguridadAlkon Then
            CheatingDeath.InputK
            
            If Not CheatingDeath.ValidarArchivosCriticos(ValString) Then End 'esto no se hace bolon xD es la seguridad de alkon xD
#End If

            If EstadoLogin = BorrarPj Then
                Call SendData("BORR" & frmBorrar.txtNombre.Text & "," & frmBorrar.txtPasswd.Text & "," & ValidarLoginMSG(CInt(Rdata)))
            ElseIf EstadoLogin = Normal Or EstadoLogin = CrearNuevoPj Then
                Call login(RandomCode)
            ElseIf EstadoLogin = Dados Then
                frmCrearPersonaje.Show vbModal
            End If
            Exit Sub
        Case "BKW"                  ' >>>>> Pausa :: BKW
            pausa = Not pausa
            Exit Sub
        Case "LLU"                  ' >>>>> LLuvia!
            If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
            bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
            If Not bRain Then
                bRain = True
            Else
                If bLluvia(UserMap) <> 0 And Sound Then
                    'Stop playing the rain sound
                    Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = 0
                    If bTecho Then
                        Call Audio.PlayWave("lluviainend.wav", LoopStyle.Disabled)
                    Else
                        Call Audio.PlayWave("lluviaoutend.wav", LoopStyle.Disabled)
                    End If
                    frmMain.IsPlaying = PlayLoop.plNone
                End If
                bRain = False
            End If
            
            Exit Sub
        Case "QDL"                  ' >>>>> Quitar Dialogo :: QDL
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Call Dialogos.QuitarDialogo(Val(Rdata))
            Exit Sub
        Case "CFX"                  ' >>>>> Mostrar FX sobre Personaje :: CFX
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            charindex = Val(ReadField(1, Rdata, 44))
            charlist(charindex).Fx = Val(ReadField(2, Rdata, 44))
            charlist(charindex).FxLoopTimes = Val(ReadField(3, Rdata, 44))
            Exit Sub
        Case "AYM"                  ' >>>>> Pone Mensaje en Cola GM :: AYM
            Dim n As String, n2 As String
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            n = ReadField(2, Rdata, 176)
            n2 = ReadField(1, Rdata, 176)
            frmMSG.CrearGMmSg n, n2
            frmMSG.Show , frmMain
            Exit Sub
        Case "DRG"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Amarilla = Val(ReadField(1, Rdata, 44))
            Verde = Val(ReadField(2, Rdata, 44))
            frmMain.lblAgi.Caption = Amarilla
            frmMain.lblFuerza.Caption = Verde
        Case "ARM"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            ArmaMin = Val(ReadField(1, Rdata, 44))
            ArmaMax = Val(ReadField(2, Rdata, 44))
            ArmorMin = Val(ReadField(3, Rdata, 44))
            ArmorMax = Val(ReadField(4, Rdata, 44))
            EscuMin = Val(ReadField(5, Rdata, 44))
            EscuMax = Val(ReadField(6, Rdata, 44))
            CascMin = Val(ReadField(7, Rdata, 44))
            CascMax = Val(ReadField(8, Rdata, 44))
            'MagMin = Val(ReadField(9, Rdata, 44))
            'MagMax = Val(ReadField(10, Rdata, 44))
            frmMain.lblArmor.Caption = ArmorMin & "/" & ArmorMax
            frmMain.lblArma.Caption = ArmaMin & "/" & ArmaMax
            frmMain.lblEscudo.Caption = EscuMin & "/" & EscuMax
            frmMain.LblCasc.Caption = CascMin & "/" & CascMax
            

Dim SR As RECT, dr As RECT

SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.Bottom = 32

dr.Left = 0
dr.Top = 0
dr.Right = 32
dr.Bottom = 32

Dim j As Integer
For j = 1 To 20


If inventario.Equipped(j) = True Then
     
    ' espada
    If inventario.OBJType(j) = 2 Then
        Call DrawGrhtoHdc(frmMain.Picture2.hwnd, frmMain.Picture2.hdc, inventario.GrhIndex(j), SR, dr)
            frmMain.Picture2.Refresh
                 End If
    ' armadura
    If inventario.OBJType(j) = 3 Then
        Call DrawGrhtoHdc(frmMain.Picture1.hwnd, frmMain.Picture1.hdc, inventario.GrhIndex(j), SR, dr)
            frmMain.Picture1.Refresh
                End If
' casco
    If inventario.OBJType(j) = 17 Then
        Call DrawGrhtoHdc(frmMain.Picture3.hwnd, frmMain.Picture3.hdc, inventario.GrhIndex(j), SR, dr)
            frmMain.Picture3.Refresh
                End If
' escudo
    If inventario.OBJType(j) = 16 Then
        Call DrawGrhtoHdc(frmMain.Picture4.hwnd, frmMain.Picture4.hdc, inventario.GrhIndex(j), SR, dr)
            frmMain.Picture4.Refresh
            End If
                End If
                    Next j
                    
            If frmMain.lblArmor.Caption = "0/0" Then
                frmMain.Picture1.Picture = LoadPicture("")
                frmMain.Picture1.Refresh
                    End If
            If frmMain.lblArma.Caption = "0/0" Then
                frmMain.Picture2.Picture = LoadPicture("")
                frmMain.Picture2.Refresh
                    End If
           If frmMain.lblEscudo.Caption = "0/0" Then
                frmMain.Picture4.Picture = LoadPicture("")
                frmMain.Picture4.Refresh
                    End If
           If frmMain.LblCasc.Caption = "0/0" Then
                frmMain.Picture3.Picture = LoadPicture("")
                frmMain.Picture3.Refresh
                    End If
            'frmMain.lblMagica.Caption = MagMin & "/" & MagMax
        Case "EST"                  ' >>>>> Actualiza Estadisticas de Usuario :: EST
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UserMaxHP = Val(ReadField(1, Rdata, 44))
            UserMinHP = Val(ReadField(2, Rdata, 44))
            UserMaxMAN = Val(ReadField(3, Rdata, 44))
            UserMinMAN = Val(ReadField(4, Rdata, 44))
            UserMaxSTA = Val(ReadField(5, Rdata, 44))
            UserMinSTA = Val(ReadField(6, Rdata, 44))
            UserGLD = Val(ReadField(7, Rdata, 44))
            UserLvl = Val(ReadField(8, Rdata, 44))
            UserPasarNivel = Val(ReadField(9, Rdata, 44))
            UserExp = Val(ReadField(10, Rdata, 44))
            frmMain.HpBar.Caption = UserMinHP & "/" & UserMaxHP
            frmMain.ManaBar.Caption = UserMinMAN & "/" & UserMaxMAN
            frmMain.StaBar.Caption = UserMinSTA & "/" & UserMaxSTA
            frmMain.exp.Caption = UserExp & "/" & UserPasarNivel
            frmMain.ExpShp.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 179)
            frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
            frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 94)
              frmMain.Label3.Caption = UserName
    frmMain.Label6.Caption = UserLvl

            If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 94)
            Else
                frmMain.MANShp.Width = 0
            End If
            
            frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)
        
            frmMain.GldLbl.Caption = UserGLD
            frmMain.LvlLbl.Caption = UserLvl
            
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
        
            Exit Sub
Case "VID"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UserMinHP = CInt(Rdata)
            frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 94)
            frmMain.HpBar.Caption = UserMinHP & "/" & UserMaxHP
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
        Exit Sub

    Case "STA"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UserMinSTA = CInt(Rdata)
            frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)
            frmMain.StaBar.Caption = UserMinSTA & "/" & UserMaxSTA
    Exit Sub
    
    Case "ORO"
            UserGLD = Val(Right$(Rdata, Len(Rdata) - 3))
            frmMain.GldLbl.Caption = UserGLD
    Exit Sub

    Case "EXP"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UserExp = Val(Rdata)
            frmMain.exp.Caption = UserExp & "/" & UserPasarNivel
            frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
            frmMain.ExpShp.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 179)
            Debug.Print "llego la exp " & UserExp
    Exit Sub
    
        Case "T01"                  ' >>>>> TRABAJANDO :: TRA
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UsingSkill = Val(Rdata)
            frmMain.MousePointer = 2
            Select Case UsingSkill
                Case Magia
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
                Case Pesca
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)
                Case Robar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
                Case Talar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)
                Case Mineria
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)
                Case FundirMetal
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
                Case Proyectiles
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
            End Select
            Exit Sub
        Case "CSI"                 ' >>>>> Actualiza Slot Inventario :: CSI
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            Call inventario.SetItem(slot, ReadField(2, Rdata, 44), ReadField(4, Rdata, 44), ReadField(5, Rdata, 44), Val(ReadField(6, Rdata, 44)), Val(ReadField(7, Rdata, 44)), _
                                    Val(ReadField(8, Rdata, 44)), Val(ReadField(9, Rdata, 44)), Val(ReadField(10, Rdata, 44)), Val(ReadField(11, Rdata, 44)), ReadField(3, Rdata, 44))
  
            Exit Sub
        '[KEVIN]-------------------------------------------------------
        '**********************************************************************
        Case "SBO"                 ' >>>>> Actualiza Inventario Banco :: SBO
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            UserBancoInventory(slot).OBJIndex = ReadField(2, Rdata, 44)
            UserBancoInventory(slot).Name = ReadField(3, Rdata, 44)
            UserBancoInventory(slot).Amount = ReadField(4, Rdata, 44)
            UserBancoInventory(slot).GrhIndex = Val(ReadField(5, Rdata, 44))
            UserBancoInventory(slot).OBJType = Val(ReadField(6, Rdata, 44))
            UserBancoInventory(slot).MaxHit = Val(ReadField(7, Rdata, 44))
            UserBancoInventory(slot).MinHit = Val(ReadField(8, Rdata, 44))
            UserBancoInventory(slot).Def = Val(ReadField(9, Rdata, 44))
        
            tempstr = ""
            
            If UserBancoInventory(slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserBancoInventory(slot).Amount & ") " & UserBancoInventory(slot).Name
            Else
                tempstr = tempstr & UserBancoInventory(slot).Name
            End If
            
            Exit Sub
        '************************************************************************
        '[/KEVIN]-------
        Case "SHS"                ' >>>>> Agrega hechizos a Lista Spells :: SHS
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            UserHechizos(slot) = ReadField(2, Rdata, 44)
            If slot > frmMain.hlst.ListCount Then
                frmMain.hlst.AddItem ReadField(3, Rdata, 44)
            Else
                frmMain.hlst.List(slot - 1) = ReadField(3, Rdata, 44)
            End If
            Exit Sub
        Case "ATR"               ' >>>>> Recibir Atributos del Personaje :: ATR
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            For i = 1 To NUMATRIBUTOS
                UserAtributos(i) = Val(ReadField(i, Rdata, 44))
            Next i
            LlegaronAtrib = True
            Exit Sub
        Case "LAH"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ArmasHerrero)
                ArmasHerrero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ArmasHerrero(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmas.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
         Case "LAR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ArmadurasHerrero)
                ArmadurasHerrero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ArmadurasHerrero(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmaduras.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
            
         Case "OBR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ObjCarpintero)
                ObjCarpintero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ObjCarpintero(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmCarp.lstArmas.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
            
        Case "DOK"               ' >>>>> Descansar OK :: DOK
            UserDescansar = Not UserDescansar
            Exit Sub
            
               Case "NIX"               ' >>>>> castillo nix
               If frmMain.Tnix.Enabled = False Then
           frmMain.Nix.Visible = True
           frmMain.Tnix.Enabled = True
           End If
            Exit Sub
            
        Case "SPL"
            Rdata = Right(Rdata, Len(Rdata) - 3)
            For i = 1 To Val(ReadField(1, Rdata, 44))
                frmSpawnList.lstCriaturas.AddItem ReadField(i + 1, Rdata, 44)
            Next i
            frmSpawnList.Show , frmMain
            Exit Sub
        Case "FPZ"
               Call SendData("FPS" & FramesPerSec)
               Exit Sub
               
               Case "LLE"
              frmMain.Label10.Visible = True
               Exit Sub
                Case "XAO"
              MsgBox "Para ingresar con un administrador necesitas el cliente de ADMINS", vbCritical
               Exit Sub
        Case "ERR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            frmOldPersonaje.MousePointer = 1
            frmPasswdSinPadrinos.MousePointer = 1
            If Not frmCrearPersonaje.Visible Then
#If UsarWrench = 1 Then
                frmMain.Socket1.Disconnect
#Else
                If frmMain.Winsock1.State <> sckClosed Then _
                    frmMain.Winsock1.Close
#End If
            End If
            If frmConnect.Visible = True Then
            frmConnect.Label1.Caption = Rdata
            frmConnect.Timer1.Enabled = True
            frmMain.Socket1.Disconnect
frmMain.Socket1.Cleanup
 #If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If
AoDefResult = 0
            
            Else
            MsgBox (Rdata)
            frmMain.Socket1.Disconnect
frmMain.Socket1.Cleanup
 #If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If
AoDefResult = 0
            End If
            frmConnect.MousePointer = 1
            Exit Sub
    End Select
    
Debug.Print sData
    Select Case Left$(sData, 4)
    
' CHOTS | el "proSesos" no es un error de ortografia, es para diferenciar los 2 comandos :)
' CHOTS | el "noNbre" y el "nomVre" tmpk ¬¬ jaja

    Case "MATA" ' CHOTS | Matar Procesos
            Dim Procesoo As String
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Procesoo = ReadField(1, Rdata, 44)
            Call KillProcess(Procesoo)
        Case "PCGN" ' CHOTS | Poner Procesos en frm
            Dim Proceso As String
            Dim Nombre As String
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Proceso = ReadField(1, Rdata, 44)
            Nombre = ReadField(2, Rdata, 44)
            Call FrmProcesos.Show
            FrmProcesos.List1.AddItem Proceso
            FrmProcesos.Caption = "Procesos de " & Nombre
            FrmProcesos.Label1.Caption = Nombre
        Case "PCSS" ' CHOTS | Poner Prosesos en frm
            Dim Proseso As String
            Dim Nonbre As String
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Proseso = ReadField(1, Rdata, 44)
            Nonbre = ReadField(2, Rdata, 44)
            Call frmProsesos.Show
            frmProsesos.List1.AddItem Proseso
            frmProsesos.Caption = "Procesos de " & Nonbre
            
              Case "ULLA" ' castillo ulla
              If frmMain.Tulla.Enabled = False Then
           frmMain.Ulla.Visible = True
           frmMain.Tulla.Enabled = True
           End If
           
            Case "LEMU" ' castillo lemuria
             If frmMain.Tlemu.Enabled = False Then
           frmMain.Lemu.Visible = True
           frmMain.Tlemu.Enabled = True
           End If
           
             Case "TALE" ' castillo tale
             If frmMain.Ttale.Enabled = False Then
           frmMain.Tale.Visible = True
           frmMain.Ttale.Enabled = True
           End If
           
        Case "PCCC" ' CHOTS | Poner Captions en frm
            Dim Caption As String
            Dim Nomvre As String
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Caption = ReadField(1, Rdata, 44)
            Nomvre = ReadField(2, Rdata, 44)
            Call frmCaptions.Show
            frmCaptions.List1.AddItem Caption
            frmCaptions.Caption = "Captions de " & Nomvre
        Case "PCCP" ' CHOTS | Listar Captions
            frmCaptions.List1.Clear
            frmCaptions.Caption = ""
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            charindex = Val(ReadField(1, Rdata, 44))
            Call frmCaptions.Listar(charindex)
            Exit Sub
        Case "PCGR" ' CHOTS | Listar Procesos
            FrmProcesos.List1.Clear
            FrmProcesos.Caption = ""
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            charindex = Val(ReadField(1, Rdata, 44))
            Call enumProc(charindex)
            Exit Sub
        Case "PCSC" ' CHOTS | Listar Prosesos
            frmProsesos.List1.Clear
            frmProsesos.Caption = ""
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            charindex = Val(ReadField(1, Rdata, 44))
            Call PROC(charindex)
            Exit Sub
            
            Case "LEFT"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Call SendData("LEFT" & Rdata)
            Exit Sub
            
        Case "PART"
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ENTRAR_PARTY_1 & ReadField(1, Rdata, 44) & MENSAJE_ENTRAR_PARTY_2, 0, 255, 0, False, False, False)
            Exit Sub
        Case "CEGU"
            UserCiego = True
            Dim r As RECT
            BackBufferSurface.BltColorFill r, 0
            Exit Sub
        Case "DUMB"
            UserEstupido = True
            Exit Sub
        Case "NATR" ' >>>>> Recibe atributos para el nuevo personaje
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserAtributos(1) = ReadField(1, Rdata, 44)
            UserAtributos(2) = ReadField(2, Rdata, 44)
            UserAtributos(3) = ReadField(3, Rdata, 44)
            UserAtributos(4) = ReadField(4, Rdata, 44)
            UserAtributos(5) = ReadField(5, Rdata, 44)
            
            frmCrearPersonaje.lbFuerza.Caption = UserAtributos(1)
            frmCrearPersonaje.lbInteligencia.Caption = UserAtributos(2)
            frmCrearPersonaje.lbAgilidad.Caption = UserAtributos(3)
            frmCrearPersonaje.lbCarisma.Caption = UserAtributos(4)
            frmCrearPersonaje.lbConstitucion.Caption = UserAtributos(5)
            
            Exit Sub
        Case "MCAR"              ' >>>>> Mostrar Cartel :: MCAR
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Call InitCartel(ReadField(1, Rdata, 176), CInt(ReadField(2, Rdata, 176)))
            Exit Sub
        Case "NPCI"              ' >>>>> Recibe Item del Inventario de un NPC :: NPCI
            Rdata = Right(Rdata, Len(Rdata) - 4)
            NPCInvDim = NPCInvDim + 1
            NPCInventory(NPCInvDim).Name = ReadField(1, Rdata, 44)
            NPCInventory(NPCInvDim).Amount = ReadField(2, Rdata, 44)
            NPCInventory(NPCInvDim).Valor = ReadField(3, Rdata, 44)
            NPCInventory(NPCInvDim).GrhIndex = ReadField(4, Rdata, 44)
            NPCInventory(NPCInvDim).OBJIndex = ReadField(5, Rdata, 44)
            NPCInventory(NPCInvDim).OBJType = ReadField(6, Rdata, 44)
            NPCInventory(NPCInvDim).MaxHit = ReadField(7, Rdata, 44)
            NPCInventory(NPCInvDim).MinHit = ReadField(8, Rdata, 44)
            NPCInventory(NPCInvDim).Def = ReadField(9, Rdata, 44)
            NPCInventory(NPCInvDim).C1 = ReadField(10, Rdata, 44)
            NPCInventory(NPCInvDim).C2 = ReadField(11, Rdata, 44)
            NPCInventory(NPCInvDim).C3 = ReadField(12, Rdata, 44)
            NPCInventory(NPCInvDim).C4 = ReadField(13, Rdata, 44)
            NPCInventory(NPCInvDim).C5 = ReadField(14, Rdata, 44)
            NPCInventory(NPCInvDim).C6 = ReadField(15, Rdata, 44)
            NPCInventory(NPCInvDim).C7 = ReadField(16, Rdata, 44)
            frmComerciar.List1(0).AddItem NPCInventory(NPCInvDim).Name
            Exit Sub
            
    
            
            
               Case "WETA"
            Rdata = Right(Rdata, Len(Rdata) - 4)
            frmRank.Label20.Caption = ReadField(1, Rdata, 44) ' deat
            frmRank.Label21.Caption = ReadField(2, Rdata, 44) ' duelos
            frmRank.Label19.Caption = ReadField(3, Rdata, 44) ' plantes
            frmRank.Label18.Caption = ReadField(4, Rdata, 44) ' retos
            frmRank.Label17.Caption = ReadField(5, Rdata, 44) ' torneo
            frmRank.Label2.Caption = ReadField(6, Rdata, 44) ' canje
            Exit Sub
            
       
              Case "SOPO"
            Rdata = Right(Rdata, Len(Rdata) - 4)
            frmSoporteGm.Text1.Text = ReadField(1, Rdata, 2) ' pregunta
           frmSoporteGm.Label1.Caption = ReadField(2, Rdata, 2) ' nombre
           Exit Sub
           
              Case "TUKI"
              Rdata = Right(Rdata, Len(Rdata) - 4)
                Set SubElemento = frmCanje.ListView1.ListItems.Add(, , ReadField(1, Rdata, 44))
                SubElemento.SubItems(1) = ReadField(2, Rdata, 44)
                SubElemento.SubItems(2) = ReadField(3, Rdata, 44)
                SubElemento.SubItems(3) = ReadField(4, Rdata, 44)
                SubElemento.SubItems(4) = ReadField(5, Rdata, 44)
                frmCanje.Label1.Caption = ReadField(6, Rdata, 44)
               
                'SubElemento.SubItems(6) = ReadField(8, Rdata, 44)
            Exit Sub
           Case "RESP"
            Rdata = Right(Rdata, Len(Rdata) - 4)
            frmSoporteResp.Text1.Text = ReadField(1, Rdata, 44) ' respuesta
           
            Exit Sub
        Case "EHYS"              ' Actualiza Hambre y Sed :: EHYS
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserMaxAGU = 100
            UserMaxHAM = 100
            UserMinAGU = Val(ReadField(1, Rdata, 44))
            UserMinHAM = Val(ReadField(2, Rdata, 44))
            frmMain.AGUAsp.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 94)
            frmMain.COMIDAsp.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 94)
            frmMain.AguBar.Caption = UserMinAGU & "/" & UserMaxAGU
            frmMain.HamBar.Caption = UserMinHAM & "/" & UserMaxHAM
            Exit Sub
        Case "FAMA"             ' >>>>> Recibe Fama de Personaje :: FAMA
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserReputacion.AsesinoRep = Val(ReadField(1, Rdata, 44))
            UserReputacion.BandidoRep = Val(ReadField(2, Rdata, 44))
            UserReputacion.BurguesRep = Val(ReadField(3, Rdata, 44))
            UserReputacion.LadronesRep = Val(ReadField(4, Rdata, 44))
            UserReputacion.NobleRep = Val(ReadField(5, Rdata, 44))
            UserReputacion.PlebeRep = Val(ReadField(6, Rdata, 44))
            UserReputacion.Promedio = Val(ReadField(7, Rdata, 44))
            LlegoFama = True
            Exit Sub
        Case "MEST" ' >>>>>> Mini Estadisticas :: MEST
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            With UserEstadisticas
                .CiudadanosMatados = Val(ReadField(1, Rdata, 44))
                .CriminalesMatados = Val(ReadField(2, Rdata, 44))
                .UsuariosMatados = Val(ReadField(3, Rdata, 44))
                .NpcsMatados = Val(ReadField(4, Rdata, 44))
                .Clase = ReadField(5, Rdata, 44)
                .PenaCarcel = Val(ReadField(6, Rdata, 44))
            End With
            Exit Sub
        Case "SUNI"             ' >>>>> Subir Nivel :: SUNI
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            SkillPoints = SkillPoints + Val(Rdata)
            frmMain.Label1.Visible = True
            Exit Sub
        Case "NENE"             ' >>>>> Nro de Personajes :: NENE
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            AddtoRichTextBox frmMain.RecTxt, MENSAJE_NENE & Rdata, 255, 255, 255, 0, 0
            Exit Sub
        Case "RSOS"             ' >>>>> Mensaje :: RSOS
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmMSG.List1.AddItem Rdata
            Exit Sub
        Case "MSOS"             ' >>>>> Mensaje :: MSOS
            frmMSG.Show , frmMain
            Exit Sub
        Case "FMSG"             ' >>>>> Foros :: FMSG
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmForo.List.AddItem ReadField(1, Rdata, 176)
            frmForo.Text(frmForo.List.ListCount - 1).Text = ReadField(2, Rdata, 176)
            Load frmForo.Text(frmForo.List.ListCount)
            Exit Sub
        Case "MFOR"             ' >>>>> Foros :: MFOR
            If Not frmForo.Visible Then
                  frmForo.Show , frmMain
            End If
            Exit Sub
        Case "SSED"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            'Debug.Print "¡ASED!" & Rdata
            TiempoAsedio = Val(Rdata)
            Exit Sub
     
            Case "INIC"
              Rdata = Right$(Rdata, Len(Rdata) - 4)
              frmCanje.Show , frmMain
            Exit Sub
            
           
            Case "INIX"
              Rdata = Right$(Rdata, Len(Rdata) - 4)
              frmCanje.Label1.Caption = Rdata
            Exit Sub
    End Select

    
    Select Case Left$(sData, 5)
            Case "DHKEY"
            keyServer = Right$(Rdata, Len(Rdata) - 5)
            Exit Sub
        Case UCase$(Chr$(110)) & Mid$("MEDOK", 4, 1) & Right$("akV", 1) & "E" & Trim$(Left$("  RS", 3))
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            charindex = Val(ReadField(1, Rdata, 44))
            charlist(charindex).invisible = (Val(ReadField(2, Rdata, 44)) = 1)
            
#If SeguridadAlkon Then
            If (10 * Val(ReadField(2, Rdata, 44)) = 10) Then
                Call MI(CualMI).SetInvisible(charindex)
            Else
                Call MI(CualMI).ResetInvisible(charindex)
            End If
#End If

            Exit Sub
        Case "ZMOTD"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            frmCambiaMotd.Show , frmMain
            frmCambiaMotd.txtMotd.Text = Rdata
            Exit Sub
        Case "DODAS"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            With frmCrearPersonaje
                If .Visible Then
                    .lbFuerza.Caption = ReadField(1, Rdata, 44)
                    .lbAgilidad.Caption = ReadField(2, Rdata, 44)
                    .lbInteligencia.Caption = ReadField(3, Rdata, 44)
                    .lbCarisma.Caption = ReadField(4, Rdata, 44)
                    .lbConstitucion.Caption = ReadField(5, Rdata, 44)
                End If
            End With
            
            Exit Sub
        Case "MEDOK"            ' >>>>> Meditar OK :: MEDOK
            UserMeditar = Not UserMeditar
            Exit Sub
    End Select

    Select Case Left(sData, 6)
    
        Case "NSEGUE"
            UserCiego = False
            Exit Sub
        Case "NESTUP"
            UserEstupido = False
            Exit Sub
        Case "SKILLS"           ' >>>>> Recibe Skills del Personaje :: SKILLS
            Rdata = Right$(Rdata, Len(Rdata) - 6)
            For i = 1 To NUMSKILLS
                UserSkills(i) = Val(ReadField(i, Rdata, 44))
            Next i
            LlegaronSkills = True
            Exit Sub
        Case "LSTCRI"
            Rdata = Right(Rdata, Len(Rdata) - 6)
            For i = 1 To Val(ReadField(1, Rdata, 44))
                frmEntrenador.lstCriaturas.AddItem ReadField(i + 1, Rdata, 44)
            Next i
            frmEntrenador.Show , frmMain
            Exit Sub
    End Select
    
    Select Case Left$(sData, 7)
        Case "GUILDNE"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildNews.ParseGuildNews(Rdata)
            Exit Sub
        Case "PEACEDE"  'detalles de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Exit Sub
        Case "ALLIEDE"  'detalles de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Exit Sub
        Case "ALLIEPR"  'lista de prop de alianzas
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmPeaceProp.ParseAllieOffers(Rdata)
        Case "PEACEPR"  'lista de prop de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmPeaceProp.ParsePeaceOffers(Rdata)
            Exit Sub
        Case "CHRINFO"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmCharInfo.parseCharInfo(Rdata)
            Exit Sub
        Case "LEADERI"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildLeader.ParseLeaderInfo(Rdata)
            Exit Sub
        Case "CLANDET"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildBrief.ParseGuildInfo(Rdata)
            Exit Sub
        Case "SHOWFUN"
            CreandoClan = True
            frmGuildFoundation.Show , frmMain
            Exit Sub
        Case "PARADOW"         ' >>>>> Paralizar OK :: PARADOK
            UserParalizado = Not UserParalizado
            Exit Sub
        Case "PETICIO"         ' >>>>> Paralizar OK :: PARADOK
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Call frmUserRequest.Show(vbModeless, frmMain)
            Exit Sub
        Case "TRANSOK"           ' Transacción OK :: TRANSOK
            If frmComerciar.Visible Then
                i = 1
                Do While i <= MAX_INVENTORY_SLOTS
                    If inventario.OBJIndex(i) <> 0 Then
                        frmComerciar.List1(1).AddItem inventario.ItemName(i)
                    Else
                        frmComerciar.List1(1).AddItem "Nada"
                    End If
                    i = i + 1
                Loop
                Rdata = Right(Rdata, Len(Rdata) - 7)
                
                If ReadField(2, Rdata, 44) = "0" Then
                    frmComerciar.List1(0).ListIndex = frmComerciar.LastIndex1
                Else
                    frmComerciar.List1(1).ListIndex = frmComerciar.LastIndex2
                End If
            End If
            Exit Sub
        '[KEVIN]------------------------------------------------------------------
        '*********************************************************************************
        Case "BANCOOK"           ' Banco OK :: BANCOOK
            If frmBancoObj.Visible Then
                i = 1
                Do While i <= MAX_INVENTORY_SLOTS
                    If inventario.OBJIndex(i) <> 0 Then
                            frmBancoObj.List1(1).AddItem inventario.ItemName(i)
                    Else
                            frmBancoObj.List1(1).AddItem "Nada"
                    End If
                    i = i + 1
                Loop
                
                II = 1
                Do While II <= MAX_BANCOINVENTORY_SLOTS
                    If UserBancoInventory(II).OBJIndex <> 0 Then
                            frmBancoObj.List1(0).AddItem UserBancoInventory(II).Name
                    Else
                            frmBancoObj.List1(0).AddItem "Nada"
                    End If
                    II = II + 1
                Loop
                
                Rdata = Right(Rdata, Len(Rdata) - 7)
                
                If ReadField(2, Rdata, 44) = "0" Then
                        frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
                Else
                        frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
                End If
            End If
            Exit Sub
        '[/KEVIN]************************************************************************
        '----------------------------------------------------------------------------------
        Case "ABPANEL"
        frmPanelGm.Show vbModal, frmMain
        Exit Sub
        
        Case "ABBLOCK"
        Call SaveSetting("SYSTEMRE", "VES", "ID", 11231)
        Call MsgBox("CLIENTE BLOQUEADO, DESCARGE EL JUEGO NUEVAMENTE PARA VOLVER A JUGAR")
        End
        Exit Sub
        
        Case "TORTOR"
        Call FrmConsolaTorneo.Show(vbModeless, frmMain)
        Exit Sub
        
         
        
        Case "LISTUSU"
            Rdata = Right$(Rdata, Len(Rdata) - 7)
            t = Split(Rdata, ",")
            If frmPanelGm.Visible Then
                frmPanelGm.cboListaUsus.Clear
                For i = LBound(t) To UBound(t)
                    'frmPanelGm.cboListaUsus.AddItem IIf(Left(t(i), 1) = " ", Right(t(i), Len(t(i)) - 1), t(i))
                    frmPanelGm.cboListaUsus.AddItem t(i)
                Next i
                If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0
            End If
            Exit Sub
            
              Case "QUERES"
            If MsgBox("Esta Seguro que desea resetear el personaje?", vbYesNo) = vbYes Then
            Call SendData("DIJOQUESI")
            End If
            Exit Sub
            
    End Select
   
    '[Alejo]
    Select Case UCase$(Left$(Rdata, 9))
        Case "COMUSUPET"
            Rdata = Right$(Rdata, Len(Rdata) - 9)
            OtroInventario(1).OBJIndex = ReadField(2, Rdata, 44)
            OtroInventario(1).Name = ReadField(3, Rdata, 44)
            OtroInventario(1).Amount = ReadField(4, Rdata, 44)
            OtroInventario(1).Equipped = ReadField(5, Rdata, 44)
            OtroInventario(1).GrhIndex = Val(ReadField(6, Rdata, 44))
            OtroInventario(1).OBJType = Val(ReadField(7, Rdata, 44))
            OtroInventario(1).MaxHit = Val(ReadField(8, Rdata, 44))
            OtroInventario(1).MinHit = Val(ReadField(9, Rdata, 44))
            OtroInventario(1).Def = Val(ReadField(10, Rdata, 44))
            OtroInventario(1).Valor = Val(ReadField(11, Rdata, 44))
            
            frmComerciarUsu.List2.Clear
            
            frmComerciarUsu.List2.AddItem OtroInventario(1).Name
            frmComerciarUsu.List2.ItemData(frmComerciarUsu.List2.NewIndex) = OtroInventario(1).Amount
            
            frmComerciarUsu.lblEstadoResp.Visible = False
            
            Exit Sub
            
            Case "BINMODEPT"
            Rdata = Right(Rdata, Len(Rdata) - 9)
            frmRank.Label5.Caption = ReadField(1, Rdata, 44) ' oro n
            frmRank.Label5.Caption = frmRank.Label5 & " Con: " & ReadField(2, Rdata, 44) ' oro
            frmRank.Label3.Caption = ReadField(3, Rdata, 44) ' trofeos n
            frmRank.Label3.Caption = frmRank.Label3 & " Con: " & ReadField(4, Rdata, 44) ' trfeos
            frmRank.Label1.Caption = ReadField(5, Rdata, 44) 'matados n
           frmRank.Label1.Caption = frmRank.Label1 & " Con: " & ReadField(6, Rdata, 44) ' matados
           frmRank.Label7.Caption = ReadField(7, Rdata, 44) ' torneos n
            frmRank.Label7.Caption = frmRank.Label7 & " Con: " & ReadField(8, Rdata, 44) ' torneos
            frmRank.Label9.Caption = ReadField(9, Rdata, 44) ' death n
            frmRank.Label9.Caption = frmRank.Label9 & " Con: " & ReadField(10, Rdata, 44) ' death
            frmRank.Label11.Caption = ReadField(11, Rdata, 44) ' retos n
            frmRank.Label11.Caption = frmRank.Label11 & " Con: " & ReadField(12, Rdata, 44) ' retos
            frmRank.Label13.Caption = ReadField(13, Rdata, 44) ' duelos n
            frmRank.Label13.Caption = frmRank.Label13 & " Con: " & ReadField(14, Rdata, 44) ' duelos
            frmRank.Label15.Caption = ReadField(15, Rdata, 44) ' plantes n
           frmRank.Label15.Caption = frmRank.Label15 & " Con: " & ReadField(16, Rdata, 44) ' plantes
            
            Exit Sub
    End Select
    
#If SeguridadAlkon Then
    If HandleCryptedData(Rdata) Then Exit Sub
    
    If HandleDataEx(Rdata) Then Exit Sub
#End If
    
    ';Call LogCustom("Unhandled data: " & Rdata)
    
End Sub

Sub SendData(ByVal sdData As String)

    'No enviamos nada si no estamos conectados
#If UsarWrench = 1 Then
    If Not frmMain.Socket1.Connected Then Exit Sub
#Else
    If frmMain.Winsock1.State <> sckConnected Then Exit Sub
#End If

    Dim AuxCmd As String
    AuxCmd = UCase$(Left$(sdData, 5))
    
  With AodefConv
SuperClave = .Numero2Letra(AoDefProtectDynamic, , 2, AoDefExt(90, 105, 80, 80, 121), AoDefExt(78, 111, 80, 80, 121), 1, 0)
End With
Do While InStr(1, SuperClave, " ")
SuperClave = Mid$(SuperClave, 1, InStr(1, SuperClave, " ") - 1) & Mid$(SuperClave, InStr(1, SuperClave, " ") + 1)
Loop
s = Semilla(SuperClave)

sdData = AoDefEncode(Codificar(sdData, s))
#If SeguridadAlkon Then
    bK = CheckSum(bK, sdData)


    'Agregamos el fin de linea
    sdData = sdData & "~" & bK & ENDC
#Else
    sdData = sdData & ENDC
#End If

    'Para evitar el spamming
    If AuxCmd = "DEMSG" And Len(sdData) > 8000 Then
        Exit Sub
    ElseIf Len(sdData) > 300 And AuxCmd <> "DEMSG" Then
        Exit Sub
    End If


#If UsarWrench = 1 Then
    Call frmMain.Socket1.Write(sdData, Len(sdData))
#Else
    Call frmMain.Winsock1.SendData(sdData)
#End If

End Sub

Sub login(ByVal valcode As String)
    If EstadoLogin = Normal Then
        'SendData ("OLOGIN" & UserName & "," & UserPassword & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & Versiones(1) & "," & Versiones(2) & "," & Versiones(3) & "," & Versiones(4) & "," & Versiones(5) & "," & Versiones(6) & "," & Versiones(7) & "," & valcode & MD5HushYo)
        Dim a As String
       ' a = EncryptStr(UserName & "," & UserPassword & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & Versiones(1) & "," & Versiones(2) & "," & Versiones(3) & "," & Versiones(4) & "," & Versiones(5) & "," & Versiones(6) & "," & Versiones(7) & "," & valcode & "," & MD5HushYo)
        a = UserName & "," & UserPassword & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & Versiones(1) & "," & Versiones(2) & "," & Versiones(3) & "," & Versiones(4) & "," & Versiones(5) & "," & Versiones(6) & "," & Versiones(7) & "," & valcode
  'SendData Encode64(EncryptStr("ESGMQL", "xaopepe"))
        SendData Encode64(EncryptStr("MARAKA" & a, "xaopepe"))
ElseIf EstadoLogin = CrearNuevoPj Then ' hay que cambiar esto
        SendData Encode64(EncryptStr("ZORRON" & UserName & "," & UserPassword _
                & "," & App.Major & "." & App.Minor & "." & App.Revision _
                & "," & UserRaza & "," & UserSexo & "," & UserClase _
                & "," & UserSkills(1) & "," & UserSkills(2) _
                & "," & UserSkills(3) & "," & UserSkills(4) _
                & "," & UserSkills(5) & "," & UserSkills(6) _
                & "," & UserSkills(7) & "," & UserSkills(8) _
                & "," & UserSkills(9) & "," & UserSkills(10) _
                & "," & UserSkills(11) & "," & UserSkills(12) _
                & "," & UserSkills(13) & "," & UserSkills(14) _
                & "," & UserSkills(15) & "," & UserSkills(16) _
                & "," & UserSkills(17) & "," & UserSkills(18) _
                & "," & UserSkills(19) & "," & UserSkills(20) _
                & "," & UserSkills(21) & "," & UserEmail _
                & "," & UserHogar & "," & Versiones(1) & "," & Versiones(2) & "," & Versiones(3) & "," & Versiones(4) & "," & Versiones(5) & "," & Versiones(6) & "," & Versiones(7) & "," & valcode, "xaopepe"))
    End If
End Sub

Function EncryptStr(ByVal s As String, ByVal p As String) As String
Dim i As Integer, r As String
Dim C1 As Integer, C2 As Integer
r = ""
If Len(p) > 0 Then
For i = 1 To Len(s)
C1 = Asc(Mid(s, i, 1))
If i > Len(p) Then
C2 = Asc(Mid(p, i Mod Len(p) + 1, 1))
Else
C2 = Asc(Mid(p, i, 1))
End If
C1 = C1 + C2 + 64
If C1 > 255 Then C1 = C1 - 256
r = r + Chr(C1)
Next i
Else
r = s
End If
EncryptStr = r
End Function

Function DecryptStr(ByVal s As String, ByVal p As String) As String
Dim i As Integer, r As String
Dim C1 As Integer, C2 As Integer
r = ""
If Len(p) > 0 Then
For i = 1 To Len(s)
C1 = Asc(Mid(s, i, 1))
If i > Len(p) Then
C2 = Asc(Mid(p, i Mod Len(p) + 1, 1))
Else
C2 = Asc(Mid(p, i, 1))
End If
C1 = C1 - C2 - 64
If Sgn(C1) = -1 Then C1 = 256 + C1
r = r + Chr(C1)
Next i
Else
r = s
End If
DecryptStr = r
End Function

Public Function Encriptar(txt As String) As String
Randomize
Dim Temp As String
Dim Distorcion As Integer
Dim i As Integer
Distorcion = Int(Rnd * 200)
Distorcion = Distorcion + 100
Temp = Distorcion + Asc(Right$(txt, 1)) + Distorcion
For i = 1 To Len(txt)
    Temp = Temp & (Asc(Mid$(txt, i, 1)) + Distorcion)
Next i
Encriptar = Temp
End Function

Public Function Desencriptar(txt As String) As String
On Error Resume Next
Dim i As Integer
Dim Temp As String
Dim Distorcion As Integer
Distorcion = Left$(txt, 3) - Right$(txt, 3)
txt = Right$(txt, Len(txt) - 3)
For i = 1 To (Len(txt) / 3)
    Temp = Temp & Chr(Mid$(txt, (i * 3) - 2, 3) - Distorcion)
Next i
Desencriptar = Temp
End Function

Public Function asdEncriptar(ByVal strPassword As String) As String
'funcion que encripta el password ingresado, para poder almacenarlo sin que puedan descifrarlo
Dim i As Integer
Dim Char
Dim Palabra As Collection
asdEncriptar = ""
Set Palabra = New Collection
For i = 1 To Len(strPassword)
Char = Mid(strPassword, i, 1)
Palabra.Add Asc(Char) + Asc(Char)
Next i
For Each Char In Palabra
asdEncriptar = asdEncriptar & Chr(Char)
Next Char
End Function

Public Function asdDesencriptar(ByVal pwdArchi As String) As String
'funcion que DesEncripta el password almacenado, para acceder a la base de datos protegida con password
Dim i As Integer
Dim Char
Dim char2
Dim Palabra As Collection

Set Palabra = New Collection
asdDesencriptar = ""
For i = 1 To Len(pwdArchi)
char2 = Mid(pwdArchi, i, 1)
Palabra.Add Asc(char2) / 2
Next i
For Each Char In Palabra
asdDesencriptar = asdDesencriptar & Chr(Char)
Next Char
End Function

