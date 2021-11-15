Attribute VB_Name = "GuerraRevival"
Option Explicit
Public Totalql As Integer
Private Cantban As Integer
Private Const MapaBan As Integer = 66
Private Const FortaDemon As Integer = 22
Private Const FortaDemony As Integer = 22
Private Const FortaAngel As Integer = 22
Private Const FortaAngely As Integer = 82
Private Const EsperaDemonio = 14
Private Const EsperaDemonioy = 56
Private Const EsperaAngel = 36
Private Const EsperaAngely = 56
Public Banac As Boolean
Public Banesp As Boolean
Public Bancantidad As Integer
Private Demonios As Integer
Private Angeles As Integer
Public CantidadGuerra As Integer
Private Ban_Luchadores() As Integer


Sub Ban_Entra(ByVal userindex)
On Error GoTo errordm:
Dim i As Integer

If Banac = False Then
 Call SendData(SendTarget.toindex, 0, 0, "||No hay ninguna Guerra RevivalAo!" & FONTTYPE_INFO)
 Exit Sub
End If
 
If Banesp = False Then
 Call SendData(SendTarget.toindex, 0, 0, "||La Guerra RevivalAo ya ha comenzado, te quedaste fuera!" & FONTTYPE_INFO)
 Exit Sub
End If
 
        For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)
                If (Ban_Luchadores(i) = userindex) Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas dentro!" & FONTTYPE_WARNING)
                        Exit Sub
                End If
        Next i

        For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)
            If (Ban_Luchadores(i) = -1) Then
                Ban_Luchadores(i) = userindex
                UserList(Ban_Luchadores(i)).flags.bandas = True
                 CantidadGuerra = CantidadGuerra + 1
                If Demonios <= Angeles Then
                  ' lo hago q es demonio
                   UserList(Ban_Luchadores(i)).flags.demonio = True
                   Demonios = Demonios + 1
                 ' convertir en demonio
                 Call Transforma(Ban_Luchadores(i))
                 ' lo teleporto donde los demonios
                  Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                  FuturePos.Map = MapaBan
                  FuturePos.x = EsperaDemonio: FuturePos.y = EsperaDemonioy
                  Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), NuevaPos.Map, NuevaPos.x, NuevaPos.y, True)
           
           Else
                     
                 UserList(Ban_Luchadores(i)).flags.angel = True
                 Angeles = Angeles + 1
                 ' convertir en angel
                  Call Transforma(Ban_Luchadores(i))
                 ' lo teleporto donde los angeles
                  Dim NuevaPoss As WorldPos
                  Dim FuturePoss As WorldPos
                  FuturePoss.Map = MapaBan
                  FuturePoss.x = EsperaAngel: FuturePoss.y = EsperaAngely
                  Call ClosestLegalPos(FuturePoss, NuevaPoss)
                    If NuevaPoss.x <> 0 And NuevaPoss.y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), NuevaPoss.Map, NuevaPoss.x, NuevaPoss.y, True)
                 
                End If
                 
                     Call SendData(SendTarget.toindex, userindex, 0, "||Estas dentro de la Guerra!" & FONTTYPE_INFO)
                
                    ' Call SendData(SendTarget.toall, 0, 0, "||Guerra RevivalAo: Entra el participante " & UserList(userindex).name & FONTTYPE_INFO)
                
                  If (i = UBound(Ban_Luchadores)) Then
                    
                    Banesp = False
                    Call Banauto_Empieza
                  End If
              
                  Exit Sub
          End If
        Next i
errordm:
End Sub
Sub Destransforma(ByVal userindex As Integer)
On Error GoTo errordm
If UserList(userindex).flags.bandas = True Then

    UserList(userindex).char.Body = UserList(userindex).CharMimetizado.Body
    UserList(userindex).char.Head = UserList(userindex).CharMimetizado.Head
    UserList(userindex).char.CascoAnim = UserList(userindex).CharMimetizado.CascoAnim
    UserList(userindex).char.ShieldAnim = UserList(userindex).CharMimetizado.ShieldAnim
    UserList(userindex).char.WeaponAnim = UserList(userindex).CharMimetizado.WeaponAnim
  
        
    UserList(userindex).Counters.Mimetismo = 0
    UserList(userindex).flags.Mimetizado = 0
     
    '[MaTeO 9]
     Call ChangeUserChar(SendTarget.ToMap, userindex, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
     '[/MaTeO 9]
 
End If
errordm:
End Sub
Sub Transforma(ByVal userindex As Integer)
On Error GoTo errordm:
If UserList(userindex).flags.demonio = True Then

With UserList(userindex)
        .CharMimetizado.Body = .char.Body
        .CharMimetizado.Head = .char.Head
        .CharMimetizado.CascoAnim = .char.CascoAnim
      
        .CharMimetizado.ShieldAnim = .char.ShieldAnim
        .CharMimetizado.WeaponAnim = .char.WeaponAnim
        
        .flags.Mimetizado = 1
        
        'ahora pongo local el del enemigo
        .char.Body = 292

    
               '[MaTeO 9]
        Call ChangeUserChar(SendTarget.ToMap, 0, .pos.Map, userindex, .char.Body, .char.Head, .char.Heading, .char.WeaponAnim, .char.ShieldAnim, .char.CascoAnim, .char.Alas)
        '[/MaTeO 9]
    End With
    End If
    
    If UserList(userindex).flags.angel = True Then

With UserList(userindex)
        .CharMimetizado.Body = .char.Body
        .CharMimetizado.Head = .char.Head
        .CharMimetizado.CascoAnim = .char.CascoAnim
      
        .CharMimetizado.ShieldAnim = .char.ShieldAnim
        .CharMimetizado.WeaponAnim = .char.WeaponAnim
        
        .flags.Mimetizado = 1
        
        'ahora pongo local el del enemigo
        .char.Body = 291

    
             '[MaTeO 9]
        Call ChangeUserChar(SendTarget.ToMap, 0, .pos.Map, userindex, .char.Body, .char.Head, .char.Heading, .char.WeaponAnim, .char.ShieldAnim, .char.CascoAnim, .char.Alas)
        '[/MaTeO 9]
    End With
    End If
errordm:
End Sub
Sub Ban_Comienza(ByVal giles As Integer)
On Error GoTo errordm
If Banac = True Then
 Call SendData(SendTarget.toindex, 0, 0, "||Ya hay una Guerra RevivalAo!!" & FONTTYPE_INFO)
 Exit Sub
 End If
 If Banesp = True Then
 Call SendData(SendTarget.toindex, 0, 0, "||La Guerra RevivalAo ya ha comenzado!" & FONTTYPE_INFO)
 Exit Sub
 End If
Cantban = giles

   Call SendData(SendTarget.toall, 0, 0, "||Guerra RevivalAo: Esta empezando Una nueva Guerra RevivalAo. Tienen 2 minutos para unirse a un bando, para Participar /REVIVALAO - (NO CAE INVENTARIO)" & FONTTYPE_GUERRA)
        Call SendData(SendTarget.toall, 0, 0, "TW48")
        Banac = True
        Banesp = True
         ReDim Ban_Luchadores(1 To Cantban) As Integer
        Dim i As Integer
        For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)
                Ban_Luchadores(i) = -1
        Next i
errordm:
End Sub

Sub Ban_Muere(ByVal userindex As Integer)
On Error GoTo errord
If UserList(userindex).flags.bandas = True Then
 If UserList(userindex).flags.demonio = True Then
 
                Dim NuevaPosDemon As WorldPos
                  Dim FuturePosDemon As WorldPos
                    FuturePosDemon.Map = MapaBan
                    FuturePosDemon.x = FortaDemon: FuturePosDemon.y = FortaDemony
                    Call ClosestLegalPos(FuturePosDemon, NuevaPosDemon)
                    If NuevaPosDemon.x <> 0 And NuevaPosDemon.y <> 0 Then Call WarpUserChar(userindex, NuevaPosDemon.Map, NuevaPosDemon.x, NuevaPosDemon.y, True)
                    End If
                    
                    If UserList(userindex).flags.angel = True Then
                Dim NuevaPosAngel As WorldPos
                  Dim FuturePosAngel As WorldPos
                    FuturePosAngel.Map = MapaBan
                    FuturePosAngel.x = FortaAngel: FuturePosAngel.y = FortaAngely
                    Call ClosestLegalPos(FuturePosAngel, NuevaPosAngel)
                    If NuevaPosAngel.x <> 0 And NuevaPosAngel.y <> 0 Then Call WarpUserChar(userindex, NuevaPosAngel.Map, NuevaPosAngel.x, NuevaPosAngel.y, True)
                    End If

End If
errord:
End Sub
Sub Ban_Desconecta(ByVal userindex As Integer)
On Error GoTo errordm
If UserList(userindex).flags.bandas = True Then
If UserList(userindex).flags.demonio = True Then
Demonios = Demonios - 1
End If
If UserList(userindex).flags.angel = True Then
Angeles = Angeles - 1
End If
Call Destransforma(userindex)
UserList(userindex).flags.bandas = False
UserList(userindex).flags.demonio = False
UserList(userindex).flags.angel = False

Call WarpUserChar(userindex, 1, 50, 50, True)
End If
errordm:
End Sub
Sub Ban_Cancela()
On Error GoTo errordm
If Banac = False And Banesp = False Then
Exit Sub
End If
    Banesp = False
    Banac = False
   
  ReDim Preserve Ban_Luchadores(1 To CantidadGuerra) As Integer
    Call SendData(SendTarget.toall, 0, 0, "||Guerra RevivalAo: Guerra RevivalAo Automatica Cancelada Por Game Master" & FONTTYPE_GUERRA)
    Dim i As Integer
    For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)
                If (Ban_Luchadores(i) <> -1) Then
                        Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                    FuturePos.Map = 1
                    FuturePos.x = 50: FuturePos.y = 50
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    
                If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), NuevaPos.Map, NuevaPos.x, NuevaPos.y, True)
                       Call Destransforma(Ban_Luchadores(i))
                    UserList(Ban_Luchadores(i)).flags.bandas = False
                    UserList(Ban_Luchadores(i)).flags.demonio = False
                    UserList(Ban_Luchadores(i)).flags.angel = False
                    Demonios = 0
                    Angeles = 0
                    CantidadGuerra = 0
                    Call RespGuerrasDemonio
                    Call RespGuerrasAngeles
                 
                End If
        Next i
errordm:
End Sub

Sub Banauto_Cancela()
On Error GoTo errordmm
If Banac = False And Banesp = False Then
Exit Sub
End If
    Banesp = False
    Banac = False
    
  ReDim Preserve Ban_Luchadores(1 To CantidadGuerra) As Integer
    Call SendData(SendTarget.toall, 0, 0, "||Guerra RevivalAo: Guerra RevivalAo Automatica cancelada por falta de participantes." & FONTTYPE_GUERRA)
    Dim i As Integer
    For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)
                If (Ban_Luchadores(i) <> -1) Then
                        Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                    FuturePos.Map = 1
                    FuturePos.x = 50: FuturePos.y = 50
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), NuevaPos.Map, NuevaPos.x, NuevaPos.y, True)
                    Call Destransforma(Ban_Luchadores(i))
                    UserList(Ban_Luchadores(i)).flags.bandas = False
                    UserList(Ban_Luchadores(i)).flags.demonio = False
                    UserList(Ban_Luchadores(i)).flags.angel = False
                    Demonios = 0
                    Angeles = 0
                    CantidadGuerra = 0
                    Call RespGuerrasDemonio
                    Call RespGuerrasAngeles
                    
                End If
        Next i
errordmm:
End Sub
Sub Reyes_Bandas()
On Error GoTo errordm:
Dim Npc3 As Integer
Dim Npc3Pos As WorldPos
Npc3 = 940
Npc3Pos.Map = 66
Npc3Pos.x = 77
Npc3Pos.y = 23

Dim Npc4 As Integer
Dim Npc4Pos As WorldPos
Npc4 = 941
Npc4Pos.Map = 66
Npc4Pos.x = 77
Npc4Pos.y = 77
Call SpawnNpc(val(Npc3), Npc3Pos, True, False)
        Call SpawnNpc(val(Npc4), Npc4Pos, True, False)
errordm:
End Sub
Sub Banauto_Empieza()
On Error GoTo errordm

  Banesp = False
 
 
  
  ReDim Preserve Ban_Luchadores(1 To CantidadGuerra) As Integer
 
   Call SendData(SendTarget.toall, 0, 0, "||Guerra RevivalAo: Empieza la Guerra!!" & FONTTYPE_GUERRA)
   Call Reyes_Bandas
    Dim i As Integer
    For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)
                If (Ban_Luchadores(i) <> -1) Then
                If UserList(Ban_Luchadores(i)).flags.demonio = True Then
                Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                    FuturePos.Map = MapaBan
                    FuturePos.x = FortaDemon: FuturePos.y = FortaDemony
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), NuevaPos.Map, NuevaPos.x, NuevaPos.y, True)
                    End If
                    
                    If UserList(Ban_Luchadores(i)).flags.angel = True Then
                Dim NuevaPoss As WorldPos
                  Dim FuturePoss As WorldPos
                    FuturePoss.Map = MapaBan
                    FuturePoss.x = FortaAngel: FuturePoss.y = FortaAngely
                    Call ClosestLegalPos(FuturePoss, NuevaPoss)
                    If NuevaPoss.x <> 0 And NuevaPoss.y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), NuevaPoss.Map, NuevaPoss.x, NuevaPoss.y, True)
                    End If
                End If
        Next i
errordm:
End Sub
Sub Ban_Demonios()
On Error GoTo errordm
    Dim i As Integer
    For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)
    If UserList(Ban_Luchadores(i)).flags.bandas = True Then
    
                If (Ban_Luchadores(i) <> -1) Then
                        Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                    FuturePos.Map = 1
                    FuturePos.x = 50: FuturePos.y = 50
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If UserList(Ban_Luchadores(i)).flags.demonio = True Then
                    UserList(Ban_Luchadores(i)).Stats.GLD = UserList(Ban_Luchadores(i)).Stats.GLD + 1000000
                    Call SendUserStatsBox(Ban_Luchadores(i))
                    End If
                     If UserList(Ban_Luchadores(i)).flags.bandas = True Then
                    If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), NuevaPos.Map, NuevaPos.x, NuevaPos.y, True)
                    End If
                   Call Destransforma(Ban_Luchadores(i))
                   UserList(Ban_Luchadores(i)).flags.bandas = False
                    UserList(Ban_Luchadores(i)).flags.demonio = False
                    UserList(Ban_Luchadores(i)).flags.angel = False
                    
                    Banac = False
                    Banesp = False
                    Demonios = 0
                    Angeles = 0
                    CantidadGuerra = 0
                End If
          
                End If
        Next i
errordm:
End Sub
Sub Ban_Angeles()
On Error GoTo errordm
    Dim i As Integer
    For i = LBound(Ban_Luchadores) To UBound(Ban_Luchadores)
    If UserList(Ban_Luchadores(i)).flags.bandas = True Then
  
                If (Ban_Luchadores(i) <> -1) Then
                        Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                    FuturePos.Map = 1
                    FuturePos.x = 50: FuturePos.y = 50
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If UserList(Ban_Luchadores(i)).flags.angel = True Then
                    UserList(Ban_Luchadores(i)).Stats.GLD = UserList(Ban_Luchadores(i)).Stats.GLD + 1000000
                    Call SendUserStatsBox(Ban_Luchadores(i))
                    End If
                     If UserList(Ban_Luchadores(i)).flags.bandas = True Then
                    If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then Call WarpUserChar(Ban_Luchadores(i), NuevaPos.Map, NuevaPos.x, NuevaPos.y, True)
                    End If
                     Call Destransforma(Ban_Luchadores(i))
                   UserList(Ban_Luchadores(i)).flags.bandas = False
                    UserList(Ban_Luchadores(i)).flags.demonio = False
                    UserList(Ban_Luchadores(i)).flags.angel = False
                   
                    Banac = False
                    Banesp = False
                    Demonios = 0
                    Angeles = 0
                    CantidadGuerra = 0
                End If
         
                End If
        Next i
errordm:
End Sub
