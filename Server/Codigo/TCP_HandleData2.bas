Attribute VB_Name = "TCP_HandleData2"


Option Explicit

Public Sub HandleData_2(ByVal userindex As Integer, rData As String, ByRef Procesado As Boolean)


Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim TIndex As Integer
Dim tName As String
Dim tMessage As String
Dim AuxInd As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim Arg3 As String
Dim Arg4 As String
Dim Ver As String
Dim encpass As String
Dim Pass As String
Dim mapa As Integer
Dim name As String
Dim ind
Dim n As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim x As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim t() As String
Dim i As Integer


Procesado = True 'ver al final del sub

If UCase$(Left$(rData, 9)) = "/REALMSG " Then
rData = Right$(rData, Len(rData) - 9)
        If UserList(userindex).flags.Privilegios > PlayerType.SemiDios Or UserList(userindex).flags.PertAlCons = 1 Then
        If rData <> "" Then
        Call SendData(SendTarget.ToRealYRMs, 0, 0, "||" & UserList(userindex).name & ">" & rData & FONTTYPE_CONSEJOVesA)
        End If
        End If
        Exit Sub
End If
    
If UCase$(Left$(rData, 9)) = "/CAOSMSG " Then
rData = Right$(rData, Len(rData) - 9)
        If UserList(userindex).flags.Privilegios > PlayerType.SemiDios Or UserList(userindex).flags.PertAlConsCaos = 1 Then
        If rData <> "" Then
        Call SendData(SendTarget.ToCaosYRMs, 0, 0, "||" & UserList(userindex).name & ">" & rData & FONTTYPE_WETAS)
        End If
        End If
        Exit Sub
End If
    
If UCase$(Left$(rData, 8)) = "/CIUMSG " Then
rData = Right$(rData, Len(rData) - 8)
        'Solo dioses, admins y RMS
        If UserList(userindex).flags.Privilegios > PlayerType.SemiDios Or UserList(userindex).flags.PertAlCons = 1 Then
        If rData <> "" Then
        Call SendData(SendTarget.ToCiudadanosYRMs, 0, 0, "||" & UserList(userindex).name & ">" & rData & FONTTYPE_CONSEJOVesA)
        End If
        End If
Exit Sub
End If

'#################### LISTA DE AMIGOS by GALLE ######################
If UCase$(Left$(rData, 3)) = "/MP" Then
Dim Mensaje As String
Dim MPname As String
rData = Right$(rData, Len(rData) - 3)
MPname = ReadField(2, rData, 64)
Mensaje = ReadField(3, rData, 64)
TIndex = NameIndex(MPname)
If TIndex <= 0 Then
Call SendData(toindex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
Else
Call SendData(toindex, TIndex, 0, "||" & UserList(userindex).name & " dice: " & Mensaje & FONTTYPE_TALK)
Call SendData(toindex, userindex, 0, "||El usuario recibio el Mensaje." & FONTTYPE_INFO)
End If
Exit Sub
End If


If UCase$(Left$(rData, 8)) = "/CRIMSG " Then
rData = Right$(rData, Len(rData) - 8)
        If UserList(userindex).flags.Privilegios > PlayerType.SemiDios Or UserList(userindex).flags.PertAlConsCaos = 1 Then
        If rData <> "" Then
        Call SendData(SendTarget.ToCriminalesYRMs, 0, 0, "||" & UserList(userindex).name & ">" & rData & FONTTYPE_WETAS)
        End If
        End If
        Exit Sub
End If
If UCase$(Left(rData, 3)) = "/SI" Then
If Encuesta.ACT = 0 Then Exit Sub
If UserList(userindex).flags.VotEnc = True Then Exit Sub
Encuesta.EncSI = Encuesta.EncSI + 1
Call SendData(SendTarget.toindex, userindex, 0, "||Has votado exitosamente." & FONTTYPE_INFO)
UserList(userindex).flags.VotEnc = True
Exit Sub
End If

If UCase$(Left(rData, 3)) = "/NO" Then
If Encuesta.ACT = 0 Then Exit Sub
If UserList(userindex).flags.VotEnc = True Then Exit Sub
Encuesta.EncNO = Encuesta.EncNO + 1
Call SendData(SendTarget.toindex, userindex, 0, "||Has votado exitosamente." & FONTTYPE_INFO)
UserList(userindex).flags.VotEnc = True
Exit Sub
End If
        If UCase$(Left$(rData, 8)) = "/ALPETE " Then
        Dim Cantidad As Long
        Cantidad = UserList(userindex).Stats.GLD
        rData = Right$(rData, Len(rData) - 8)
        rData = Desencriptar(rData)
        TIndex = NameIndex(ReadField(1, rData, 32))
        Arg1 = ReadField(2, rData, 32)
        If TIndex <= 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub
        End If
       If Distancia(UserList(userindex).pos, UserList(TIndex).pos) > 3 Then
       Call SendData(SendTarget.toindex, userindex, 0, "||Estas Demasiado Lejos" & FONTTYPE_WARNING)
        Exit Sub
        End If
                    If val(Arg1) > Cantidad Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||No tenes esa cantidad de oro" & FONTTYPE_WARNING)
                    ElseIf val(Arg1) < 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||No podes transferir cantidades negativas" & FONTTYPE_WARNING)
                    Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||¡Le regalaste " & val(Arg1) & " monedas de oro a " & UserList(TIndex).name & "!" & FONTTYPE_ORO)
                    Call SendData(SendTarget.toindex, TIndex, 0, "||¡" & UserList(userindex).name & " te regalo " & val(Arg1) & " monedas de oro!" & FONTTYPE_ORO)
                    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - val(Arg1)
                    UserList(TIndex).Stats.GLD = UserList(TIndex).Stats.GLD + val(Arg1)
                    Call EnviarOro(TIndex)
                    Call EnviarOro(userindex)
                    Exit Sub
                    End If
                    Exit Sub
                    End If

    Select Case UCase$(rData)
    
    Case "/MOV"
                If UserList(userindex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                    Exit Sub
                End If
               
                If UserList(userindex).flags.TargetUser = 0 Then Exit Sub
               
                If UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 0 Then Exit Sub
  
  If Distancia(UserList(userindex).pos, UserList(UserList(userindex).flags.TargetUser).pos) > 2 Then Exit Sub
  
                    Dim CadaverUltPos As WorldPos
                    CadaverUltPos.Y = UserList(UserList(userindex).flags.TargetUser).pos.Y + 1
                    CadaverUltPos.x = UserList(UserList(userindex).flags.TargetUser).pos.x
                    CadaverUltPos.Map = UserList(UserList(userindex).flags.TargetUser).pos.Map
                    
                    Dim CadaverUltPos2 As WorldPos
                    CadaverUltPos2.Y = UserList(UserList(userindex).flags.TargetUser).pos.Y
                    CadaverUltPos2.x = UserList(UserList(userindex).flags.TargetUser).pos.x + 1
                    CadaverUltPos2.Map = UserList(UserList(userindex).flags.TargetUser).pos.Map
                    
                    Dim CadaverUltPos3 As WorldPos
                    CadaverUltPos3.Y = UserList(UserList(userindex).flags.TargetUser).pos.Y - 1
                    CadaverUltPos3.x = UserList(UserList(userindex).flags.TargetUser).pos.x
                    CadaverUltPos3.Map = UserList(UserList(userindex).flags.TargetUser).pos.Map
                    
                    Dim CadaverUltPos4 As WorldPos
                    CadaverUltPos4.Y = UserList(UserList(userindex).flags.TargetUser).pos.Y
                    CadaverUltPos4.x = UserList(UserList(userindex).flags.TargetUser).pos.x - 1
                    CadaverUltPos4.Map = UserList(UserList(userindex).flags.TargetUser).pos.Map
                
                If LegalPos(CadaverUltPos.Map, CadaverUltPos.x, CadaverUltPos.Y, False) Then
                Call WarpUserChar(UserList(userindex).flags.TargetUser, CadaverUltPos.Map, CadaverUltPos.x, CadaverUltPos.Y, False)
                ElseIf LegalPos(CadaverUltPos2.Map, CadaverUltPos2.x, CadaverUltPos2.Y, False) Then
                Call WarpUserChar(UserList(userindex).flags.TargetUser, CadaverUltPos2.Map, CadaverUltPos2.x, CadaverUltPos2.Y, False)
                ElseIf LegalPos(CadaverUltPos3.Map, CadaverUltPos3.x, CadaverUltPos3.Y, False) Then
                Call WarpUserChar(UserList(userindex).flags.TargetUser, CadaverUltPos3.Map, CadaverUltPos3.x, CadaverUltPos3.Y, False)
                ElseIf LegalPos(CadaverUltPos4.Map, CadaverUltPos4.x, CadaverUltPos4.Y, False) Then
                Call WarpUserChar(UserList(userindex).flags.TargetUser, CadaverUltPos4.Map, CadaverUltPos4.x, CadaverUltPos4.Y, False)
                Else
                Call WarpUserChar(UserList(userindex).flags.TargetUser, 1, 58, 45, True)
                End If
                UserList(userindex).flags.TargetUser = 0
    Exit Sub
    
    Case "/HOGAR"
     If UserList(userindex).pos.Map = 87 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||¡¡No puedes usar este comando en retos 2v2!!!!." & FONTTYPE_INFO)
    Exit Sub
    End If
    If UserList(userindex).pos.Map = 66 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||¡¡En guerras no puedes usar este comando!!!!." & FONTTYPE_INFO)
    Exit Sub
    End If
If EsNewbie(userindex) Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Los Newbies no Pueden Utilizar este Comando!!!." & FONTTYPE_INFO)
Exit Sub
End If
If UserList(userindex).flags.Muerto = 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Tenes que estar muerto para poder usar este comando!!!." & FONTTYPE_INFO)
Exit Sub
End If
If UserList(userindex).Counters.Pena >= 1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡¡No podes usar este comando estando encarcelado!!!." & FONTTYPE_INFO)
Exit Sub
End If
If Criminal(userindex) Then
Call WarpUserChar(userindex, 36, 41, 27, True)
Else
Call WarpUserChar(userindex, 1, 50, 50, True)
End If
Exit Sub

        Case "/COLAPAJA23"
            'No se envia más la lista completa de usuarios
            n = 0
            For LoopC = 1 To LastUser
                If UserList(LoopC).name <> "" And UserList(LoopC).flags.Privilegios <= PlayerType.Consejero Then
                    n = n + 1
                End If
            Next LoopC
            Call SendData(SendTarget.toindex, userindex, 0, "||Número de usuarios: " & n & ". Record de Usuarios Conectados Simultaneamente: " & recordusuarios & FONTTYPE_INFO)
            Exit Sub
        'Juanpa
        'Peto
Case "/CASTILLOS"
If AlmacenaDominador = vbNullString Then
Call SendData(toindex, userindex, 0, "||Castillo de Ullathorpe: Nadie." & FONTTYPE_CONSEJOCAOSVesA)
Else
Call SendData(toindex, userindex, 0, "||Castillo de Ullathorpe: " & AlmacenaDominador & " " & HoraUlla & FONTTYPE_CONSEJOCAOSVesA)
End If

If AlmacenaDominadornix = vbNullString Then
Call SendData(toindex, userindex, 0, "||Castillo de Nix: Nadie." & FONTTYPE_CONSEJOCAOSVesA)
Else
Call SendData(toindex, userindex, 0, "||Castillo de Nix: " & AlmacenaDominadornix & " " & HoraNix & FONTTYPE_CONSEJOCAOSVesA)
End If

If Lemuria = vbNullString Then
Call SendData(toindex, userindex, 0, "||Castillo de Asgard: Nadie." & FONTTYPE_CONSEJOCAOSVesA)
Else
Call SendData(toindex, userindex, 0, "||Castillo de Asgard: " & Lemuria & " " & HoraLemuria & FONTTYPE_CONSEJOCAOSVesA)
End If

If Tale = vbNullString Then
Call SendData(toindex, userindex, 0, "||Castillo de Tale: Nadie." & FONTTYPE_CONSEJOCAOSVesA)
Else
Call SendData(toindex, userindex, 0, "||Castillo de Tale: " & Tale & " " & HoraTale & FONTTYPE_CONSEJOCAOSVesA)
End If

If Fortaleza = vbNullString Then
Call SendData(toindex, userindex, 0, "||Fortaleza: Nadie." & FONTTYPE_CONSEJOCAOSVesA)
Else
Call SendData(toindex, userindex, 0, "||Fortaleza: " & Fortaleza & " " & HoraForta & FONTTYPE_CONSEJOCAOSVesA)
Exit Sub
End If

Case "/DEFTALE"
Dim positalex As Integer
positalex = RandomNumber(50, 57)
Dim positaley As Integer
positaley = RandomNumber(28, 35)
 If UserList(userindex).pos.Map = 87 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||¡¡No puedes usar este comando en retos 2v2!!!!." & FONTTYPE_INFO)
    Exit Sub
    End If
If UserList(userindex).flags.EstaDueleando1 = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a castillo estando plantes!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).flags.Paralizado = 1 Then
Call SendData(toindex, userindex, 0, "||No puedes defender el castillo estando paralizado!!" & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).GuildIndex = 0 Then
Call SendData(toindex, userindex, 0, "||No tienes clan!" & FONTTYPE_INFO)
Exit Sub
End If
If UserList(userindex).pos.Map = 66 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a fortaleza estando torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
     If UserList(userindex).pos.Map = 88 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en deatmatch." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 87 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                     If UserList(userindex).pos.Map = 75 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas en un castillo!." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 77 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas en un castillo!." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 107 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas en un castillo!." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 106 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas en un castillo!." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
If UserList(userindex).pos.Map = 79 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a defender estando torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).pos.Map = 78 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a defender estando torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
 If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes defender el castillo estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If

 If UserList(userindex).pos.Map = 67 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes defender el castillo estando en la carcel." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
 If UserList(userindex).pos.Map = 62 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes defender el castillo estando en torneo." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 
If Guilds(UserList(userindex).GuildIndex).GuildName = Tale Then
Call WarpUserChar(userindex, 107, positalex, positaley, True)
Else
Call SendData(toindex, userindex, 0, "||No perteneces al clan que ha conquistado el castillo" & FONTTYPE_INFO)
End If
Exit Sub

Case "/DEFASGD"

Dim posilemx As Integer
posilemx = RandomNumber(50, 57)
Dim posilemy As Integer
posilemy = RandomNumber(28, 35)
 If UserList(userindex).pos.Map = 87 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||¡¡No puedes usar este comando en retos 2v2!!!!." & FONTTYPE_INFO)
    Exit Sub
    End If
If UserList(userindex).flags.EstaDueleando1 = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a castillo estando plantes!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).flags.Paralizado = 1 Then
Call SendData(toindex, userindex, 0, "||No puedes defender el castillo estando paralizado!!" & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).GuildIndex = 0 Then
Call SendData(toindex, userindex, 0, "||No tienes clan!" & FONTTYPE_INFO)
Exit Sub
End If
If UserList(userindex).pos.Map = 66 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a fortaleza estando torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
    If UserList(userindex).pos.Map = 75 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas en un castillo!." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 77 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas en un castillo!." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 107 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas en un castillo!." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 106 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas en un castillo!." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
     If UserList(userindex).pos.Map = 88 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en deatmatch." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 87 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
If UserList(userindex).pos.Map = 79 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a defender estando torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).pos.Map = 78 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a defender estando torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
 If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes defender el castillo estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If

 If UserList(userindex).pos.Map = 67 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes defender el castillo estando en la carcel." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
 If UserList(userindex).pos.Map = 62 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes defender el castillo estando en torneo." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 
If Guilds(UserList(userindex).GuildIndex).GuildName = Lemuria Then
Call WarpUserChar(userindex, 106, posilemx, posilemy, True)
Else
Call SendData(toindex, userindex, 0, "||No perteneces al clan que ha conquistado el castillo" & FONTTYPE_INFO)
End If
Exit Sub

Case "/DEFULLA"
Dim posix As Integer
posix = RandomNumber(50, 57)
Dim posiy As Integer
posiy = RandomNumber(28, 35)
 If UserList(userindex).pos.Map = 87 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||¡¡No puedes usar este comando en retos 2v2!!!!." & FONTTYPE_INFO)
    Exit Sub
    End If
If UserList(userindex).flags.EstaDueleando1 = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a castillo estando plantes!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).flags.Paralizado = 1 Then
Call SendData(toindex, userindex, 0, "||No puedes defender el castillo estando paralizado!!" & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).GuildIndex = 0 Then
Call SendData(toindex, userindex, 0, "||No tienes clan!" & FONTTYPE_INFO)
Exit Sub
End If
If UserList(userindex).pos.Map = 66 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a fortaleza estando torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
     If UserList(userindex).pos.Map = 88 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en deatmatch." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 87 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
If UserList(userindex).pos.Map = 79 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a defender estando torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).pos.Map = 78 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a defender estando torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
 If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes defender el castillo estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If

 If UserList(userindex).pos.Map = 67 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes defender el castillo estando en la carcel." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
 If UserList(userindex).pos.Map = 62 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes defender el castillo estando en torneo." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 
                  If UserList(userindex).pos.Map = 75 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas en un castillo!." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 77 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas en un castillo!." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 107 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas en un castillo!." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 106 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas en un castillo!." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
If Guilds(UserList(userindex).GuildIndex).GuildName = AlmacenaDominador Then
Call WarpUserChar(userindex, 75, posix, posiy, True)
Else
Call SendData(toindex, userindex, 0, "||No perteneces al clan que ha conquistado el castillo" & FONTTYPE_INFO)
End If
Exit Sub



Case "/DEFNIX"
Dim posixx As Integer
posixx = RandomNumber(50, 57)
Dim posiyy As Integer
posiyy = RandomNumber(28, 35)
 If UserList(userindex).pos.Map = 87 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||¡¡No puedes usar este comando en retos 2v2!!!!." & FONTTYPE_INFO)
    Exit Sub
    End If
If UserList(userindex).flags.EstaDueleando1 = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a castillo estando plantes!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).flags.Paralizado = 1 Then
Call SendData(toindex, userindex, 0, "||No puedes defender el castillo estando paralizado!!" & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).GuildIndex = 0 Then
Call SendData(toindex, userindex, 0, "||No tienes clan!" & FONTTYPE_INFO)
Exit Sub
End If
If UserList(userindex).pos.Map = 66 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a fortaleza estando torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
    If UserList(userindex).pos.Map = 75 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas en un castillo!." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 77 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas en un castillo!." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 107 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas en un castillo!." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 106 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas en un castillo!." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
     If UserList(userindex).pos.Map = 88 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en deatmatch." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 87 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
If UserList(userindex).pos.Map = 79 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a defender estando torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).pos.Map = 78 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a defender estando en retos!." & FONTTYPE_WARNING)
Exit Sub
End If

 If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes defender el castillo estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
 If UserList(userindex).pos.Map = 67 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes defender el castillo estando en la carcel." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
 If UserList(userindex).pos.Map = 62 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes defender el castillo estando en torneo." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 
If Guilds(UserList(userindex).GuildIndex).GuildName = AlmacenaDominadornix Then
Call WarpUserChar(userindex, 77, posixx, posiyy, True)
Else
Call SendData(toindex, userindex, 0, "||No perteneces al clan que ha conquistado el castillo" & FONTTYPE_INFO)
End If
Exit Sub

Case "/FORTALEZA"
Dim forx As Integer
forx = RandomNumber(48, 58)
Dim fory As Integer
fory = RandomNumber(48, 56)
 If UserList(userindex).pos.Map = 87 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||¡¡No puedes usar este comando en retos 2v2!!!!." & FONTTYPE_INFO)
    Exit Sub
    End If
        If UserList(userindex).pos.Map = 75 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a fortaleza estando en castillo!." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 77 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a fortaleza estando en castillo!." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 107 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a fortaleza estando en castillo!!." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 106 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a fortaleza estando en castillo!!." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
If UserList(userindex).flags.EstaDueleando1 = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a castillo estando plantes!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).flags.Paralizado = 1 Then
Call SendData(toindex, userindex, 0, "||No puedes ir a fortaleza estando paralizado!!" & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).GuildIndex = 0 Then
Call SendData(toindex, userindex, 0, "||No tienes clan!" & FONTTYPE_INFO)
Exit Sub
End If
If UserList(userindex).pos.Map = 79 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a fortaleza estando torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).pos.Map = 66 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a fortaleza estando torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).pos.Map = 78 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a fortaleza estando en retos!." & FONTTYPE_WARNING)
Exit Sub
End If
     If UserList(userindex).pos.Map = 88 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en deatmatch." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 87 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
 If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a la fortaleza estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
 If UserList(userindex).pos.Map = 67 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a fortaleza estando en la carcel." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                If UserList(userindex).pos.Map = 62 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a fortaleza estando en torneo." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
If Guilds(UserList(userindex).GuildIndex).GuildName = AlmacenaDominador And Guilds(UserList(userindex).GuildIndex).GuildName = AlmacenaDominadornix And Guilds(UserList(userindex).GuildIndex).GuildName = Lemuria And Guilds(UserList(userindex).GuildIndex).GuildName = Tale Then
Call WarpUserChar(userindex, 76, forx, fory, True)
Else
Call SendData(toindex, userindex, 0, "||Necesitas Todos los castillos para ingresar a la Fortaleza." & FONTTYPE_INFO)
End If
Exit Sub

Case "/DUELO"
If UserList(userindex).pos.Map <> 1 And UserList(userindex).pos.Map <> 36 And UserList(userindex).pos.Map <> 102 And UserList(userindex).pos.Map <> 92 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
      Exit Sub
      End If
Dim JuanpaDuelosMap As Integer
JuanpaDuelosMap = 61
Dim JuanpaDuelosX As Integer
JuanpaDuelosX = RandomNumber(43, 58)
Dim JuanpaDuelosY As Integer
JuanpaDuelosY = RandomNumber(45, 56)
 If UserList(userindex).flags.TargetNPC = 0 Then
 Call SendData(SendTarget.toindex, userindex, 0, "||Debes seleccionar el npc de Duelos que está en el muelle de ullathorpe." & FONTTYPE_INFO)
 Exit Sub
 End If
If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Duelos Then
Call SendData(SendTarget.toindex, userindex, 0, "||Debes seleccionar el npc de Duelos que está en el muelle de ullathorpe." & FONTTYPE_INFO)
Exit Sub
End If
                '¿El NPC puede comerciar?
                If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 5 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                    Exit Sub
                End If
If UserList(userindex).flags.EstaDueleando1 = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelos estando plantes!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).pos.Map = 67 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelos estando en la carcel!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).pos.Map = 66 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a guerra estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 If UserList(userindex).pos.Map = 78 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                    If UserList(userindex).pos.Map = 88 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en deatmatch." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 87 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
If UserList(userindex).pos.Map = 79 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelos estando torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).flags.Muerto = 1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Debes estar vivo para ingresar al duelo." & FONTTYPE_WARNING)
Exit Sub
ElseIf UserList(userindex).Stats.ELV < 30 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes hacer duelos siendo menor a nivel 30." & FONTTYPE_WARNING)
Exit Sub
ElseIf MapInfo(JuanpaDuelosMap).NumUsers >= 2 Then
Call SendData(SendTarget.toindex, userindex, 0, "||La arena de duelos esta ocupada." & FONTTYPE_WARNING)
Exit Sub
ElseIf MapInfo(UserList(userindex).pos.Map).Pk = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||Estas en una zona insegura." & FONTTYPE_WARNING)
Exit Sub
ElseIf MapInfo(JuanpaDuelosMap).NumUsers = 1 Then
duelosreta = userindex

Call SendData(SendTarget.toindex, userindex, 0, "||Has sido teletransportado a la arena de duelos, para salir de ella tipea /salirduelo" & FONTTYPE_INFO)
'Juanpa
Call WarpUserChar(userindex, JuanpaDuelosMap, JuanpaDuelosX, JuanpaDuelosY, True)
       
        Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Duelos: " & UserList(duelosreta).name & " ha Aceptado el Desafio." & FONTTYPE_TALK)
Exit Sub
ElseIf MapInfo(JuanpaDuelosMap).NumUsers = 0 Then
duelosespera = userindex

Call SendData(SendTarget.toindex, userindex, 0, "||Has sido teletransportado a la arena de duelos." & FONTTYPE_WARNING)
'Juanpa
Call WarpUserChar(userindex, JuanpaDuelosMap, JuanpaDuelosX, JuanpaDuelosY, True)
        Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Duelos: " & UserList(duelosespera).name & " espera rival en la arena de duelos." & FONTTYPE_TALK)
End If
Exit Sub
'/Juanpa
'Juanpa
        Case "/SALIRDUELO"
        
            If MapInfo(61).NumUsers = 2 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes salir en medio de la pelea, tienes que estar solo en la arena para poder salir." & FONTTYPE_INFO)
            Exit Sub
            End If
            If UserList(userindex).pos.Map = 61 And userindex = duelosespera Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Has salido del duelo." & FONTTYPE_INFO)
            Call SendData(SendTarget.toAll, 0, 0, "||Duelos: " & UserList(duelosespera).name & " ha salido de la arena de duelos." & FONTTYPE_TALK)
            duelosespera = duelosreta
            numduelos = 0
            Call WarpUserChar(userindex, 1, 50, 50, True)
            Exit Sub
            End If
               If UserList(userindex).pos.Map = 61 And userindex = duelosreta Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Has salido del duelo." & FONTTYPE_INFO)
            Call SendData(SendTarget.toAll, 0, 0, "||Duelos: " & UserList(duelosreta).name & " ha salido de la arena de duelos." & FONTTYPE_TALK)
            Call WarpUserChar(userindex, 1, 50, 50, True)
            Exit Sub
            End If
          

        Case "/RANKING"
         ' ranking anterior
            'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario con más oro es: " & Ranking.MaxOro.UserName & " (" & Ranking.MaxOro.value & ")" & FONTTYPE_GUILD)
            'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario con más trofeos de oro ganados es: " & Ranking.MaxTrofeos.UserName & " (" & Ranking.MaxTrofeos.value & ")" & FONTTYPE_GUILD)
            'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario con más pjs matados es: " & Ranking.MaxUsuariosMatados.UserName & " (" & Ranking.MaxUsuariosMatados.value & ")" & FONTTYPE_GUILD)
            'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario con más Puntos de Torneo: " & Ranking.MaxTorneos.UserName & " (" & Ranking.MaxTorneos.value & ")" & FONTTYPE_GUILD)
            'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario con más Puntos de DeathMatch: " & Ranking.MaxDeaths.UserName & " (" & Ranking.MaxDeaths.value & ")" & FONTTYPE_GUILD)
            ' Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario con más Puntos de Retos: " & Ranking.MaxRetos.UserName & " (" & Ranking.MaxRetos.value & ")" & FONTTYPE_GUILD)
            ' Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario con más Puntos de Duelos: " & Ranking.MaxDuelos.UserName & " (" & Ranking.MaxDuelos.value & ")" & FONTTYPE_GUILD)
            ' Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario con más Puntos de Plantes: " & Ranking.MaxPlantes.UserName & " (" & Ranking.MaxPlantes.value & ")" & FONTTYPE_GUILD)
            Call EnviaRank(userindex)
            Call EnviaPuntos(userindex)
              SendData SendTarget.toindex, userindex, 0, "INITRAN"
            Exit Sub
        
        Case "/SALIR"
        If UserList(userindex).flags.Montado = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes salir estando en montado en tu mascota!." & FONTTYPE_INFO)
            Exit Sub
            End If
        If UserList(userindex).pos.Map = 87 Then
           Call SendData(SendTarget.toindex, userindex, 0, "||No puedes salir estando en retos 2v2." & FONTTYPE_WARNING)
                Exit Sub
            End If
            If UserList(userindex).flags.Paralizado = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes salir estando paralizado." & FONTTYPE_WARNING)
                Exit Sub
            End If
             If UserList(userindex).pos.Map = 76 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes salir estando en Fortaleza." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 88 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes salir estando en DeathMatch." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 If UserList(userindex).pos.Map = 66 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes salir estando en Guerra." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
             If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes salir estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
            End If
             If UserList(userindex).pos.Map = 62 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes salir estando en torneos." & FONTTYPE_WARNING)
                 Exit Sub
            End If
              If UserList(userindex).pos.Map = 79 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes salir estando en torneos." & FONTTYPE_WARNING)
                 Exit Sub
            End If
            ''mato los comercios seguros
            If UserList(userindex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(userindex).ComUsu.DestUsu).flags.UserLogged Then
                    If UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.DestUsu = userindex Then
                        Call SendData(SendTarget.toindex, UserList(userindex).ComUsu.DestUsu, 0, "||Comercio cancelado por el otro usuario" & FONTTYPE_TALK)
                        Call FinComerciarUsu(UserList(userindex).ComUsu.DestUsu)
                    End If
                End If
                Call SendData(SendTarget.toindex, userindex, 0, "||Comercio cancelado. " & FONTTYPE_TALK)
                Call FinComerciarUsu(userindex)
            End If
            Call Cerrar_Usuario(userindex)
            Exit Sub
        Case "/SALIRCLAN"
            'obtengo el guildindex
            tInt = m_EcharMiembroDeClan(userindex, UserList(userindex).name)
            
            If tInt > 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Dejas el clan." & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & UserList(userindex).name & " deja el clan." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||Tu no puedes salir de ningún clan." & FONTTYPE_GUILD)
            End If
            
            
            Exit Sub

            
        Case "/BALANCE"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "Z30")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 3 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                      Exit Sub
            End If
            Select Case Npclist(UserList(userindex).flags.TargetNPC).NPCtype
            Case eNPCType.Banquero
                If FileExist(CharPath & UCase$(UserList(userindex).name) & ".chr", vbNormal) = False Then
                      Call SendData(SendTarget.toindex, userindex, 0, "!!El personaje no existe, cree uno nuevo.")
                      CloseSocket (userindex)
                      Exit Sub
                End If
                Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(userindex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
            Case eNPCType.Timbero
                If UserList(userindex).flags.Privilegios > PlayerType.User Then
                    tLong = Apuestas.Ganancias - Apuestas.Perdidas
                    n = 0
                    If tLong >= 0 And Apuestas.Ganancias <> 0 Then
                        n = Int(tLong * 100 / Apuestas.Ganancias)
                    End If
                    If tLong < 0 And Apuestas.Perdidas <> 0 Then
                        n = Int(tLong * 100 / Apuestas.Perdidas)
                    End If
                    Call SendData(SendTarget.toindex, userindex, 0, "||Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & tLong & " (" & n & "%) Jugadas: " & Apuestas.Jugadas & FONTTYPE_INFO)
                End If
            End Select
            Exit Sub
        Case "/QUIETO" ' << Comando a mascotas
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(userindex).flags.Muerto = 1 Then
                          Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                          Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(userindex).flags.TargetNPC = 0 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "Z30")
                      Exit Sub
             End If
             If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 10 Then
                          Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                          Exit Sub
             End If
             If Npclist(UserList(userindex).flags.TargetNPC).MaestroUser <> _
                userindex Then Exit Sub
             Npclist(UserList(userindex).flags.TargetNPC).Movement = TipoAI.ESTATICO
             Call Expresar(UserList(userindex).flags.TargetNPC, userindex)
             Exit Sub
        Case "/ACOMPAÑAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "Z30")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 10 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                      Exit Sub
            End If
            If Npclist(UserList(userindex).flags.TargetNPC).MaestroUser <> _
              userindex Then Exit Sub
            Call FollowAmo(UserList(userindex).flags.TargetNPC)
            Call Expresar(UserList(userindex).flags.TargetNPC, userindex)
            Exit Sub
        Case "/ENTRENAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "Z30")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 10 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                      Exit Sub
            End If
            If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
            Call EnviarListaCriaturas(userindex, UserList(userindex).flags.TargetNPC)
            Exit Sub
  
        
        Case "/DESCANSAR"
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
            End If
            If HayOBJarea(UserList(userindex).pos, FOGATA) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "DOK")
                    If Not UserList(userindex).flags.Descansar Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Te acomodas junto a la fogata y comenzas a descansar." & FONTTYPE_INFO)
                    Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||Te levantas." & FONTTYPE_INFO)
                    End If
                    UserList(userindex).flags.Descansar = Not UserList(userindex).flags.Descansar
            Else
                    If UserList(userindex).flags.Descansar Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Te levantas." & FONTTYPE_INFO)
                        
                        UserList(userindex).flags.Descansar = False
                        Call SendData(SendTarget.toindex, userindex, 0, "DOK")
                        Exit Sub
                    End If
                    Call SendData(SendTarget.toindex, userindex, 0, "||No hay ninguna fogata junto a la cual descansar." & FONTTYPE_INFO)
            End If
            Exit Sub
        Case "/HACEME1PT3"
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
            End If
            If UserList(userindex).Stats.MaxMAN = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Solo las clases mágicas conocen el arte de la meditación" & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(userindex).flags.Privilegios > PlayerType.User Then
                UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MaxMAN
                Call SendData(SendTarget.toindex, userindex, 0, "||Mana restaurado" & FONTTYPE_VENENO)
                Call EnviarMn(userindex)
                Exit Sub
            End If
            Call SendData(SendTarget.toindex, userindex, 0, "MEDOK")
            If Not UserList(userindex).flags.Meditando Then
               Call SendData(SendTarget.toindex, userindex, 0, "Z23")
            Else
               Call SendData(SendTarget.toindex, userindex, 0, "Z16")
            End If
           UserList(userindex).flags.Meditando = Not UserList(userindex).flags.Meditando
            'Barrin 3/10/03 Tiempo de inicio al meditar
            If UserList(userindex).flags.Meditando Then
                UserList(userindex).Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF
                Call SendData(SendTarget.toindex, userindex, 0, "Z37")
                
                UserList(userindex).char.loops = LoopAdEternum
                If UserList(userindex).Stats.ELV < 5 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & FXIDs.FXMEDITARNW & "," & LoopAdEternum)
                    UserList(userindex).char.FX = FXIDs.FXMEDITARNW
                ElseIf UserList(userindex).Stats.ELV < 15 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & FXIDs.FXMEDITARAZULNW & "," & LoopAdEternum)
                    UserList(userindex).char.FX = FXIDs.FXMEDITARAZULNW
                ElseIf UserList(userindex).Stats.ELV < 25 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & FXIDs.FXMEDITARFUEGUITO & "," & LoopAdEternum)
                    UserList(userindex).char.FX = FXIDs.FXMEDITARFUEGUITO
                ElseIf UserList(userindex).Stats.ELV < 30 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & FXIDs.FXMEDITARFUEGO & "," & LoopAdEternum)
                    UserList(userindex).char.FX = FXIDs.FXMEDITARFUEGO
                ElseIf UserList(userindex).Stats.ELV < 35 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & FXIDs.FXMEDITARMEDIANO & "," & LoopAdEternum)
                    UserList(userindex).char.FX = FXIDs.FXMEDITARMEDIANO
                ElseIf UserList(userindex).Stats.ELV < 45 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & FXIDs.FXMEDITARAZULCITO & "," & LoopAdEternum)
                    UserList(userindex).char.FX = FXIDs.FXMEDITARAZULCITO
                ElseIf UserList(userindex).Stats.ELV < 55 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & FXIDs.FXMEDITARGRIS & "," & LoopAdEternum)
                    UserList(userindex).char.FX = FXIDs.FXMEDITARGRIS
                Else
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & FXIDs.FXMEDITARFULL & "," & LoopAdEternum)
                    UserList(userindex).char.FX = FXIDs.FXMEDITARFULL
                End If
            Else
                UserList(userindex).Counters.bPuedeMeditar = False
                
                UserList(userindex).char.FX = 0
                UserList(userindex).char.loops = 0
                Call SendData(SendTarget.ToMap, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & 0 & "," & 0)
            End If
            Exit Sub
            Case "/ACEPTAR"
            If UserList(userindex).pos.Map <> 1 And UserList(userindex).pos.Map <> 36 And UserList(userindex).pos.Map <> 102 And UserList(userindex).pos.Map <> 92 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
      Exit Sub
      End If
                   If UserList(userindex).pos.Map = 62 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a eventos en torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
            If UserList(userindex).flags.EstaDueleando1 = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelos estando plantes!." & FONTTYPE_WARNING)
Exit Sub
End If
   If UserList(userindex).flags.EnDosVDos = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya estas en reto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
          '  On Error GoTo error
          If UserList(userindex).pos.Map = 66 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a guerra estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
               If UserList(userindex).pos.Map = 88 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en deatmatch." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 87 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
          If UserList(userindex).pos.Map = 67 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a retos estando en carcel." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
          If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes retar estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 If UserList(userindex).pos.Map = 79 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes retar estando en torneos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
            If UserList(userindex).flags.EsperandoDuelo = True Then
            Call SendData(toindex, userindex, 0, "||¡¡Ya has retado antes, espera que acepten tu desafio anterior para poder aceptar uno nuevo.!" & FONTTYPE_TALK)
               Exit Sub
            End If
         If UserList(UserList(userindex).flags.TargetUser).flags.EnDosVDos = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(userindex).flags.Muerto = 1 Or UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 1 Then
    Call SendData(toindex, userindex, 0, "||¡¡No se puede retar muerto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(UserList(userindex).flags.TargetUser).flags.EstaDueleando1 = True Then
    Call SendData(toindex, userindex, 0, "||¡¡No se puede retar porque esta en plante!!!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(userindex).Stats.GLD < 1000000 Then
    Call SendData(toindex, userindex, 0, "||¡¡Debes tener al menos 1.000.000 de oro!!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If MapInfo(78).NumUsers >= 2 Then
    Call SendData(toindex, userindex, 0, "||¡Ya hay un Reto!" & FONTTYPE_TALK)
    Exit Sub
    End If
  
    Call ComensarDuelo(userindex, UserList(userindex).flags.Oponente)
'error:     Call SendData(toindex, userindex, 0, "||¡No te han retado!!" & FONTTYPE_TALK)
    Exit Sub

Case "/ACEPTO"
If UserList(userindex).pos.Map <> 1 And UserList(userindex).pos.Map <> 36 And UserList(userindex).pos.Map <> 102 And UserList(userindex).pos.Map <> 92 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
      Exit Sub
      End If
       If UserList(userindex).pos.Map = 62 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a eventos en torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
    If UserList(userindex).flags.EnDosVDos = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya estas en reto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
      If UserList(userindex).flags.EstaDueleando = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya estas en reto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
          '  On Error GoTo error
          If UserList(userindex).pos.Map = 66 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a guerra estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
               If UserList(userindex).pos.Map = 88 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en deatmatch." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 87 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
          If UserList(userindex).pos.Map = 67 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a retos estando en carcel." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
          If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes retar estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 If UserList(userindex).pos.Map = 79 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes retar estando en torneos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
            If UserList(userindex).flags.EsperandoDuelo1 = True Then
            Call SendData(toindex, userindex, 0, "||¡¡Yas has invitado a retar a un usuario. Termina tu reto para poder aceptar uno nuevo!!" & FONTTYPE_TALK)
               Exit Sub
            End If
         
    If UserList(userindex).flags.Muerto = 1 Or UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 1 Then
    Call SendData(toindex, userindex, 0, "||¡¡No se puede retar muerto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
      If UserList(UserList(userindex).flags.TargetUser).flags.EstaDueleando = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(UserList(userindex).flags.TargetUser).flags.EnDosVDos = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(userindex).Stats.GLD < 500000 Then
    Call SendData(toindex, userindex, 0, "||¡¡Debes tener al menos 500.000 de oro!!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If YaHayPlante = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya hay un Reto!" & FONTTYPE_TALK)
    Exit Sub
    End If
 
    Call ComensarDueloPlantes(userindex, UserList(userindex).flags.Oponente1)
'error:     Call SendData(toindex, userindex, 0, "||¡No te han retado!!" & FONTTYPE_TALK)
    Exit Sub


'RETOS 2V2

 Case "/DUAL"
 If UserList(userindex).pos.Map <> 1 And UserList(userindex).pos.Map <> 36 And UserList(userindex).pos.Map <> 102 And UserList(userindex).pos.Map <> 92 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
      Exit Sub
      End If
        If UserList(userindex).pos.Map = 62 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a eventos en torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
 If Team.EnCurso = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Los 2vs2 Están ocupados!!!" & FONTTYPE_INFO)
        Exit Sub
    End If
    If userindex = Team.Pj1 Or userindex = Team.Pj2 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas emparejado :$!!!" & FONTTYPE_INFO)
    Exit Sub
    End If
            Dim Pj2 As Integer
            Pj2 = UserList(userindex).flags.TargetUser
        If Team.Activado = False Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Los Retos 2 v 2 están desactivados!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).flags.TargetUser = Team.Pj1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Ya tiene pareja!!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).flags.TargetUser = Team.Pj2 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Ya tiene pareja!!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(UserList(userindex).flags.TargetUser).Clase = UserList(userindex).Clase Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes conformar pareja con alguien de tu misma clase!." & FONTTYPE_WARNING)
Exit Sub
End If
            If UserList(userindex).flags.EstaDueleando1 = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a retos estando plantes!." & FONTTYPE_WARNING)
Exit Sub
End If
      If UserList(userindex).flags.EstaDueleando = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelos estando retos!." & FONTTYPE_WARNING)
Exit Sub
End If

If UserList(userindex).Stats.GLD < 1000000 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Debes tener almenos 1.000.000 para poder retar." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(UserList(userindex).flags.TargetUser).flags.EstaDueleando = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
      If UserList(UserList(userindex).flags.TargetUser).flags.EstaDueleando1 = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
         If UserList(userindex).pos.Map = 78 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 
    If UserList(userindex).pos.Map = 66 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a guerra estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
         If UserList(userindex).pos.Map = 88 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en deatmatch." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 87 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
    If UserList(userindex).pos.Map = 67 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a retos estando en carcel." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
     If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes retar estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 If UserList(userindex).pos.Map = 79 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes retar estando en torneos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 
                 
    If UserList(userindex).flags.TargetUser = userindex Then
        Call SendData(toindex, userindex, 0, "||Debes seleccionar a un personaje!." & FONTTYPE_INFO)
        Exit Sub
    End If
    
     If UserList(userindex).flags.Muerto = 1 Then
        Call SendData(toindex, userindex, 0, "||El usuario esta muerto" & FONTTYPE_INFO)
        Exit Sub
    End If
 
    If UserList(userindex).flags.TargetUser <= 0 Then
        Call SendData(toindex, userindex, 0, "||Debes seleccionar a un usuario." & FONTTYPE_INFO)
        Exit Sub
    End If
       

    If UserList(Pj2).flags.Muerto = 1 Then
        Call SendData(toindex, userindex, 0, "||El usuario esta muerto!." & FONTTYPE_INFO)
        Exit Sub
    End If
   
    
   
    If Distancia(UserList(UserList(userindex).flags.TargetUser).pos, UserList(userindex).pos) > 5 Then
        Call SendData(toindex, userindex, 0, "||Estás demasiado lejos!" & FONTTYPE_INFO)
        Exit Sub
    End If
    
       If Team.EnCurso = True Then
        Call SendData(toindex, userindex, 0, "||Los 2vs2 Están ocupados!!!" & FONTTYPE_INFO)
        Exit Sub
    End If
    If Team.Activado = False Then
        Call SendData(toindex, userindex, 0, "||Los retos 2vs2 están desactivados!!!" & FONTTYPE_INFO)
        Exit Sub
    End If
    If UserList(userindex).flags.EnDosVDos = True Then
        Call SendData(toindex, userindex, 0, "||Ya estás en 2vs2!!!" & FONTTYPE_INFO)
    End If
        
            Call SendData(toindex, Pj2, 0, "||" & UserList(userindex).name & " desea jugar un 2vs2. Haz click sobre tu pareja y escribe /SDUAL para aceptar." & FONTTYPE_INFO)
    Call SendData(toindex, userindex, 0, "||Pediste a " & UserList(Pj2).name & " que sea tu pareja." & FONTTYPE_INFO)
        UserList(userindex).flags.envioSol = True
        UserList(Pj2).flags.RecibioSol = True
        UserList(Pj2).flags.compa = userindex
        
            Exit Sub
        
            
        
        
    Case "/SDUAL"
    If UserList(userindex).pos.Map <> 1 And UserList(userindex).pos.Map <> 36 And UserList(userindex).pos.Map <> 102 And UserList(userindex).pos.Map <> 92 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
      Exit Sub
      End If
    If UserList(userindex).flags.Muerto = 1 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes aceptar si estas muerto!." & FONTTYPE_WARNING)
Exit Sub
End If
           If UserList(userindex).pos.Map = 62 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a eventos en torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
        If Team.EnCurso = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Los 2vs2 Están ocupados!!!" & FONTTYPE_INFO)
        Exit Sub
    End If
               If Team.Activado = False Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Los Retos 2 v 2 están desactivados!." & FONTTYPE_WARNING)
Exit Sub
End If
         If userindex = Team.Pj1 Or userindex = Team.Pj2 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas emparejado :$!!!" & FONTTYPE_INFO)
    Exit Sub
    End If
    If UserList(UserList(userindex).flags.TargetUser).Clase = UserList(userindex).Clase Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes conformar pareja con alguien de tu misma clase!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).flags.TargetUser = Team.Pj1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Ya tiene pareja!!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).flags.TargetUser = Team.Pj2 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Ya tiene pareja!!." & FONTTYPE_WARNING)
Exit Sub
End If
                      If UserList(userindex).flags.EstaDueleando1 = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelos estando plantes!." & FONTTYPE_WARNING)
Exit Sub
End If
             If UserList(userindex).flags.EstaDueleando = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelos estando retos!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).Stats.GLD < 1000000 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Debes tener almenos 1.000.000 para poder retar." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(UserList(userindex).flags.TargetUser).flags.EstaDueleando = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
      If UserList(UserList(userindex).flags.TargetUser).flags.EstaDueleando1 = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
          If UserList(userindex).pos.Map = 78 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If

          If UserList(userindex).pos.Map = 66 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a guerra estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
               If UserList(userindex).pos.Map = 88 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en deatmatch." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 87 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
          If UserList(userindex).pos.Map = 67 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a retos estando en carcel." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
          If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes retar estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 If UserList(userindex).pos.Map = 79 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes retar estando en torneos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 If UserList(userindex).flags.Muerto = 1 Then
        Call SendData(toindex, userindex, 0, "||Esta muerto!")
        Exit Sub
    End If
    If Team.EnCurso = True Then
        Call SendData(toindex, userindex, 0, "||Los 2vs2 Están ocupados!!!" & FONTTYPE_INFO)
        Exit Sub
    End If
    If Team.Activado = False Then
        Call SendData(toindex, userindex, 0, "||Los retos 2vs2 están desactivados!!!" & FONTTYPE_INFO)
        Exit Sub
    End If
    If UserList(userindex).flags.RecibioSol = False Then
        Call SendData(toindex, userindex, 0, "||Nadie te invitó a como pareja!!" & FONTTYPE_INFO)
    Exit Sub
    End If

    If UserList(userindex).flags.EnDosVDos = True Then
        Call SendData(toindex, userindex, 0, "||Ya estás en 2vs2!!!" & FONTTYPE_INFO)
    End If
            If Team.SonDos = True Then
                Team.pj3 = userindex
                Team.pj4 = UserList(userindex).flags.compa
                'Warpeo
                 Call WarpUserChar(Team.Pj1, 87, 41, 50)
        Call WarpUserChar(Team.Pj2, 87, 41, 51)
        Call WarpUserChar(Team.pj3, 87, 60, 50)
        Call WarpUserChar(Team.pj4, 87, 60, 51)
                 UserList(Team.Pj1).flags.EnDosVDos = True
                 UserList(Team.Pj2).flags.EnDosVDos = True
                 UserList(Team.pj3).flags.EnDosVDos = True
                 UserList(Team.pj4).flags.EnDosVDos = True
                 Team.EnCurso = True
                 Call SendData(toAll, userindex, 0, "||2Vs2: " & UserList(Team.Pj1).name & " y " & UserList(Team.Pj2).name & _
            " VS " & UserList(Team.pj3).name & " y " & UserList(Team.pj4).name & " que gane el mejor!" & FONTTYPE_RETOS2V2)
           
            ElseIf Team.SonDos = False Then
                Team.SonDos = True
                Team.Pj1 = userindex
                Team.Pj2 = UserList(userindex).flags.compa
                Call SendData(toindex, userindex, 0, "||Tu pareja es ahora " & UserList(Team.Pj2).name & " , espera contrincantes." & FONTTYPE_INFO)
                Call SendData(toindex, Team.Pj2, 0, "||Tu pareja es ahora " & UserList(userindex).name & " , espera contrincantes." & FONTTYPE_INFO)
                Call SendData(SendTarget.toAll, 0, 0, "||2Vs2: La pareja " & UserList(userindex).name & "(" & UserList(userindex).Clase & ")" & " y " & UserList(Team.Pj2).name & "(" & UserList(Team.Pj2).Clase & ")" & " Retan por 1KK !!." & FONTTYPE_RETOS2V2)
            End If
            Exit Sub
            
              '[MaTeO 10]
    Case "/VOLVER"
        If UserList(userindex).pos.Map <> 70 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Unicamente puedes utilizar este comando en el laberinto." & FONTTYPE_WARNING)
            Exit Sub
        End If
        If UserList(userindex).flags.EstaDueleando1 = True Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a retos estando plantes!." & FONTTYPE_WARNING)
            Exit Sub
        End If

        
        If UserList(userindex).flags.TargetUser = Team.Pj1 Or UserList(userindex).flags.TargetUser = Team.Pj2 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puede participar en eventos si esperas retos!!!" & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(userindex).flags.EnDosVDos = True Then
            Call SendData(toindex, userindex, 0, "||¡Ya estas en reto!!" & FONTTYPE_TALK)
            Exit Sub
        End If
        
        Call SendData(SendTarget.toindex, userindex, 0, "||¡Has vuelto a comenzar el laberinto!" & FONTTYPE_WARNING)
        Call WarpUserChar(userindex, 70, 13, 12)
    '[/MaTeO 10]
            'TERMINA RETOS 2V2
    Case "/RETAR"
    If UserList(userindex).pos.Map <> 1 And UserList(userindex).pos.Map <> 36 And UserList(userindex).pos.Map <> 102 And UserList(userindex).pos.Map <> 92 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
      Exit Sub
      End If
        If UserList(userindex).pos.Map = 62 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a eventos en torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
    If UserList(userindex).flags.EstaDueleando1 = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a retos estando plantes!." & FONTTYPE_WARNING)
Exit Sub
End If
If userindex = Team.Pj1 Or userindex = Team.Pj2 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes participar en eventos si esperas retos!!!" & FONTTYPE_INFO)
    Exit Sub
    End If
    If UserList(userindex).flags.TargetUser = Team.Pj1 Or UserList(userindex).flags.TargetUser = Team.Pj2 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||No puede participar en eventos si esperas retos!!!" & FONTTYPE_INFO)
    Exit Sub
    End If
  If UserList(userindex).flags.EnDosVDos = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya estas en reto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(userindex).pos.Map = 66 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a guerra estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
         If UserList(userindex).pos.Map = 88 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en deatmatch." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 87 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
    If UserList(userindex).pos.Map = 67 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a retos estando en carcel." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
     If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes retar estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 If UserList(userindex).pos.Map = 79 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes retar estando en torneos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
    If UserList(userindex).flags.Muerto = 1 Then
    Call SendData(toindex, userindex, 0, "||¡¡Estas Muerto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(userindex).flags.TargetUser > 0 Then
    If UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 1 Then
    Call SendData(toindex, userindex, 0, "||¡El usuario con el que quieres retar está muerto!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(UserList(userindex).flags.TargetUser).flags.EstaDueleando = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
      If UserList(UserList(userindex).flags.TargetUser).flags.EstaDueleando1 = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
     If UserList(UserList(userindex).flags.TargetUser).flags.EnDosVDos = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
    
      If UserList(userindex).Stats.GLD < 1000000 Then
    Call SendData(toindex, userindex, 0, "||¡¡Debes tener al menos 1.000.000. de oro!!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(userindex).pos.Map = 79 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a retos estando torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
    If MapInfo(78).NumUsers >= 2 Then
    Call SendData(toindex, userindex, 0, "||¡Ya hay un reto!." & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(userindex).flags.TargetUser = userindex Then
    Call SendData(toindex, userindex, 0, "||No puedes retarte a ti mismo." & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(UserList(userindex).flags.TargetUser).flags.EsperandoDuelo = True Then
    If UserList(UserList(userindex).flags.TargetUser).flags.Oponente = userindex Then
    Call ComensarDuelo(userindex, UserList(userindex).flags.TargetUser)
    Exit Sub
    End If
    If UserList(UserList(userindex).flags.TargetUser).flags.EsperandoDuelo = True Then
    Call SendData(toindex, userindex, 0, "||El usuario que intentas retar ya ha retado a otro usuario, espera que termine su reto!." & FONTTYPE_TALK)
    Exit Sub
    End If
    Else
    Call SendData(toindex, UserList(userindex).flags.TargetUser, 0, "|| " & UserList(userindex).name & " Te ha retado por 1.000.000, si quieres aceptar haz click sobre tu oponente y pon /ACEPTAR." & FONTTYPE_TALK)
    Call SendData(toindex, userindex, 0, "||Has retado por 1.000.000 a " & UserList(UserList(userindex).flags.TargetUser).name & FONTTYPE_TALK)
    UserList(userindex).flags.EsperandoDuelo = True
    
    '[MaTeO 4]
   ' UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - 1000000
    '[/MaTeO 4]
    
    UserList(userindex).flags.Oponente = UserList(userindex).flags.TargetUser
    UserList(UserList(userindex).flags.TargetUser).flags.Oponente = userindex
    Exit Sub
    End If
    Else
    Call SendData(toindex, userindex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_TALK)
    End If
    Exit Sub
    

    
    Case "/PLANTAR"
    If UserList(userindex).pos.Map <> 1 And UserList(userindex).pos.Map <> 36 And UserList(userindex).pos.Map <> 102 And UserList(userindex).pos.Map <> 92 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
      Exit Sub
      End If
          If UserList(userindex).pos.Map = 62 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a eventos en torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
      If UserList(userindex).flags.EnDosVDos = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya estas en reto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
      If UserList(userindex).flags.EstaDueleando = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya estas en reto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If userindex = Team.Pj1 Or userindex = Team.Pj2 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes participar en eventos si esperas retos!!!" & FONTTYPE_INFO)
    Exit Sub
    End If
    If UserList(userindex).flags.TargetUser = Team.Pj1 Or UserList(userindex).flags.TargetUser = Team.Pj2 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||No puede participar en eventos si esperas retos!!!" & FONTTYPE_INFO)
    Exit Sub
    End If
    If UserList(userindex).pos.Map = 66 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a guerra estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
         If UserList(userindex).pos.Map = 88 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en deatmatch." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 87 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
    If UserList(userindex).pos.Map = 67 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a retos estando en carcel." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
     If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes retar estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 If UserList(userindex).pos.Map = 79 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes retar estando en torneos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
    If UserList(userindex).flags.Muerto = 1 Then
    Call SendData(toindex, userindex, 0, "||¡¡Estas Muerto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(userindex).flags.TargetUser > 0 Then
    If UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 1 Then
    Call SendData(toindex, userindex, 0, "||¡El usuario con el que quieres retar está muerto!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(UserList(userindex).flags.TargetUser).flags.EnDosVDos = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(UserList(userindex).flags.TargetUser).flags.EstaDueleando1 = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya hay un reto!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(UserList(userindex).flags.TargetUser).flags.EstaDueleando = True Then
    Call SendData(toindex, userindex, 0, "||¡Ya esta en reto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
      If UserList(userindex).Stats.GLD < 500000 Then
    Call SendData(toindex, userindex, 0, "||¡¡Debes tener al menos 500.000. de oro!!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(userindex).pos.Map = 79 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a retos estando torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
    If YaHayPlante = True Then
    Call SendData(toindex, userindex, 0, "||¡¡¡'Ya hay un reto!!!!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(userindex).flags.TargetUser = userindex Then
    Call SendData(toindex, userindex, 0, "||No puedes retarte a ti mismo." & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(UserList(userindex).flags.TargetUser).flags.EsperandoDuelo1 = True Then
    If UserList(UserList(userindex).flags.TargetUser).flags.Oponente1 = userindex Then
    Call ComensarDueloPlantes(userindex, UserList(userindex).flags.TargetUser)
    Exit Sub
    End If
    If UserList(UserList(userindex).flags.TargetUser).flags.EsperandoDuelo1 = True Then
    Call SendData(toindex, userindex, 0, "||El usuario que intentas retar ya ha retado a otro usuario, espera que termine su reto!." & FONTTYPE_TALK)
    Exit Sub
    End If
    Else
    Call SendData(toindex, UserList(userindex).flags.TargetUser, 0, "|| " & UserList(userindex).name & " Te ha retado a Plantar por 500.000, si quieres aceptar haz click sobre tu oponente y pon /ACEPTO." & FONTTYPE_TALK)
    Call SendData(toindex, userindex, 0, "||Has retado a Plantar por 500.000 a " & UserList(UserList(userindex).flags.TargetUser).name & FONTTYPE_TALK)
    UserList(userindex).flags.EsperandoDuelo1 = True
    UserList(userindex).flags.Oponente1 = UserList(userindex).flags.TargetUser
    UserList(UserList(userindex).flags.TargetUser).flags.Oponente1 = userindex
    Exit Sub
    End If
    Else
    Call SendData(toindex, userindex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_TALK)
    End If
    Exit Sub
Case "/GANADOR"
If UserList(userindex).flags.death = True Then
If terminodeat = True Then
 Call WarpUserChar(userindex, 1, 50, 50, True)
 UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + 1000000
  UserList(userindex).Stats.PuntosDeath = UserList(userindex).Stats.PuntosDeath + 1
  UserList(userindex).Stats.PuntosCanje = UserList(userindex).Stats.PuntosCanje + 1
  Call CompruebaDeaths(userindex)
 Call SendUserStatsBox(userindex)
  Call SendData(toAll, userindex, 0, "||GANADOR DEATHMATCH: " & UserList(userindex).name & FONTTYPE_DEATH)
   Call SendData(toAll, userindex, 0, "||PREMIO: 1.000.000, Equipo Recaudado y 1 punto de DeathMatch." & FONTTYPE_DEATH)
   UserList(userindex).flags.death = False
   terminodeat = False
   deathesp = False
deathac = False
Cantidad = 0
   End If
   End If
   Exit Sub
   
   Case "/VERS"
Call EnviarResp(userindex)
SendData SendTarget.toindex, userindex, 0, "INITRES"
Exit Sub

Case "/RESETSOP"
 Call ResetSop(userindex)
 Exit Sub
 
     
Case "/CANJE"
Call EnviarCanje(userindex)
Exit Sub

 Case "/LVL"
        If UserList(userindex).Stats.ELV = 55 Then
        Exit Sub
        End If
    Dim lvl As Integer
        For lvl = 1 To 55
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.ELU
    Call CheckUserLevel(userindex)
        Call SendData(toindex, userindex, 0, "||Has Subido un nivel!" & FONTTYPE_APU)
        Next
        Exit Sub
        
        Case "/ORO"
If UserList(userindex).Stats.GLD >= 50000000 Then Exit Sub
UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + 50000000
Call SendUserStatsBox(userindex)
Call SendData(toindex, userindex, 0, "||Has ganado 50.000.000 monedas de ORO!" & FONTTYPE_ORO)
Exit Sub


'[MaTeO ASEDIO]
    Case "/ASEDIO"
        Call modAsedio.Inscribir_Asedio(userindex)
'[/MaTeO ASEDIO]

      Case "/PARTICIPAR"
      If UserList(userindex).flags.Invisible = 1 Then
      Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a eventos estando invisible!." & FONTTYPE_WARNING)
      Exit Sub
      End If
      
      If UserList(userindex).flags.Oculto = 1 Then
      Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a eventos estando invisible!." & FONTTYPE_WARNING)
      Exit Sub
      End If
      
             If UserList(userindex).pos.Map = 62 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a eventos en torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
      If UserList(userindex).flags.EstaDueleando1 = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneos estando plantes!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).flags.Muerto = 1 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Estas muerto!!!" & FONTTYPE_INFO)
    Exit Sub
    End If
If userindex = Team.Pj1 Or userindex = Team.Pj2 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes participar en eventos si esperas retos!!!" & FONTTYPE_INFO)
    Exit Sub
    End If
      If UserList(userindex).pos.Map = 66 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a guerra estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
          If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 If UserList(userindex).pos.Map = 88 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en deathmatch." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 If UserList(userindex).pos.Map = 87 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 If UserList(userindex).pos.Map = 67 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en carcel." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 If UserList(userindex).pos.Map = 78 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
      If UserList(userindex).Stats.ELV < 50 Then
      Call SendData(toindex, userindex, 0, "||Debes ser lvl 50 o mas para entrar al torneo!" & FONTTYPE_INFO)
      Exit Sub
      End If
       
Call Torneos_Entra(userindex)
Exit Sub

Case "/DEATH"
If UserList(userindex).pos.Map <> 1 And UserList(userindex).pos.Map <> 36 And UserList(userindex).pos.Map <> 102 And UserList(userindex).pos.Map <> 92 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
      Exit Sub
      End If
  If UserList(userindex).flags.Invisible = 1 Then
      Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a eventos estando invisible!." & FONTTYPE_WARNING)
      Exit Sub
      End If
      
      If UserList(userindex).flags.Oculto = 1 Then
      Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a eventos estando invisible!." & FONTTYPE_WARNING)
      Exit Sub
      End If
       If UserList(userindex).pos.Map = 62 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a eventos en torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).flags.EstaDueleando1 = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a deathmatch estando plantes!." & FONTTYPE_WARNING)
Exit Sub
End If
 If userindex = Team.Pj1 Or userindex = Team.Pj2 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes participar en eventos si esperas retos!!!" & FONTTYPE_INFO)
    Exit Sub
    End If
    If UserList(userindex).flags.Muerto = 1 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Estas muerto!!!" & FONTTYPE_INFO)
    Exit Sub
    End If
If UserList(userindex).pos.Map = 66 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
          If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 If UserList(userindex).pos.Map = 67 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en carcel." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 If UserList(userindex).pos.Map = 78 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 79 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en torneos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 87 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
      If UserList(userindex).Stats.ELV < 50 Then
      Call SendData(toindex, userindex, 0, "||Debes ser lvl 50 o mas para entrar al deathmatch!" & FONTTYPE_INFO)
      Exit Sub
      End If
       
Call death_entra(userindex)
Exit Sub

Case "/REVIVALAO"
If UserList(userindex).pos.Map <> 1 And UserList(userindex).pos.Map <> 36 And UserList(userindex).pos.Map <> 102 And UserList(userindex).pos.Map <> 92 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
      Exit Sub
      End If
  If UserList(userindex).flags.Invisible = 1 Then
      Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a eventos estando invisible!." & FONTTYPE_WARNING)
      Exit Sub
      End If
      
      If UserList(userindex).flags.Oculto = 1 Then
      Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a eventos estando invisible!." & FONTTYPE_WARNING)
      Exit Sub
      End If
If UserList(userindex).flags.Montado = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a guerra con montura!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).flags.Muerto = 1 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Estas muerto!!!" & FONTTYPE_INFO)
    Exit Sub
    End If
       If UserList(userindex).pos.Map = 62 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a eventos en torneo!." & FONTTYPE_WARNING)
Exit Sub
End If
If UserList(userindex).flags.EstaDueleando1 = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a guerra estando plantes!." & FONTTYPE_WARNING)
Exit Sub
End If
If userindex = Team.Pj1 Or userindex = Team.Pj2 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes participar en eventos si esperas retos!!!" & FONTTYPE_INFO)
    Exit Sub
    End If
If UserList(userindex).pos.Map = 88 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a guerra estando en datmatch." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
          If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a guerra estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 If UserList(userindex).pos.Map = 67 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a guerra estando en carcel." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 If UserList(userindex).pos.Map = 78 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a guerra estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 79 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a guerra estando en torneos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 87 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a guerra estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
      If UserList(userindex).Stats.ELV < 50 Then
      Call SendData(toindex, userindex, 0, "||Debes ser lvl 50 o mas para entrar a la Guerra!" & FONTTYPE_INFO)
      Exit Sub
      End If
       
Call Ban_Entra(userindex)
Exit Sub

Case "/PUNTOS"
Call SendData(toindex, userindex, 0, "||Actualmente Tienes:" & FONTTYPE_INFO)
Call SendData(toindex, userindex, 0, "||Puntos de Torneo: " & UserList(userindex).Stats.PuntosTorneo & FONTTYPE_INFO)
Call SendData(toindex, userindex, 0, "||Puntos de Deathmatch: " & UserList(userindex).Stats.PuntosDeath & FONTTYPE_INFO)
Call SendData(toindex, userindex, 0, "||Puntos de Retos : " & UserList(userindex).Stats.PuntosRetos & FONTTYPE_INFO)
Call SendData(toindex, userindex, 0, "||Puntos de Duelos: " & UserList(userindex).Stats.PuntosDuelos & FONTTYPE_INFO)
Call SendData(toindex, userindex, 0, "||Puntos de Plantes: " & UserList(userindex).Stats.PuntosPlante & FONTTYPE_INFO)
Call SendData(toindex, userindex, 0, "||Puntos de Canje: " & UserList(userindex).Stats.PuntosCanje & FONTTYPE_INFO)
Exit Sub

      Case "/TIEMPOS"
Dim tiempo1 As Integer
Dim tiempo2 As Integer
Dim tiempo3 As Integer
Dim demonioql As Integer
Dim arcangel As Integer
Dim torneoql As Integer
Dim mascotaql As Integer
Dim deatmaql As Integer
Dim GRevival As Integer
GRevival = 48
deatmaql = 63
mascotaql = 480 'mascota
tiempo1 = 360 ' demonio
tiempo2 = 380 ' arcangel
tiempo3 = 94 ' torneo
GRevival = val(GRevival) - val(bandasqls)
demonioql = val(tiempo1) - val(ContReSpawnNpc)
arcangel = val(tiempo2) - val(ContReSpawnNpc)
torneoql = val(tiempo3) - val(xao)
mascotaql = val(mascotaql) - val(mariano)
deatmaql = val(deatmaql) - val(tukiql)
Call SendData(toindex, userindex, 0, "||Quedan " & demonioql & " minutos para que renasca el Espectro Infernal!." & FONTTYPE_INFO)
Call SendData(toindex, userindex, 0, "||Quedan " & arcangel & " minutos para que renasca el Arcangel!." & FONTTYPE_INFO)
Call SendData(toindex, userindex, 0, "||Quedan " & mascotaql & " minutos para que renasca el Domador!." & FONTTYPE_INFO)
Call SendData(toindex, userindex, 0, "||Quedan " & torneoql & " minutos para el próximo torneo automatico!." & FONTTYPE_INFO)
Call SendData(toindex, userindex, 0, "||Quedan " & deatmaql & " minutos para el próximo deathmatch automatico!." & FONTTYPE_INFO)
Call SendData(toindex, userindex, 0, "||Quedan " & GRevival & " minutos para la próxima Guerra RevivalAo!." & FONTTYPE_INFO)
Exit Sub

Case "/MEZCLAR"
Dim AlasItems(2) As Integer
AlasItems(0) = 4023
AlasItems(1) = 4024
AlasItems(2) = 4025

Dim AlasLvl1(1) As Integer
AlasLvl1(0) = 4301 'Ciudadano
AlasLvl1(1) = 4302 'Criminal

Dim AlasLvl2(1) As Integer
AlasLvl2(0) = 4305 'Ciudadano
AlasLvl2(1) = 4306 'Criminal

Dim AlasLvl3(1) As Integer
AlasLvl3(0) = 4309 'Ciudadano
AlasLvl3(1) = 4310 'Criminal

Dim AlasLvl4(1) As Integer
AlasLvl4(0) = 4313 'Ciudadano
AlasLvl4(1) = 4314 'Criminal

Dim HasObjects As Boolean
Dim h As Long
HasObjects = True
For h = 0 To UBound(AlasItems)
    If Not TieneObjetos(AlasItems(h), 1, userindex) Then
        HasObjects = False
        Exit For
    End If
Next h

If HasObjects Then
    For h = 0 To UBound(AlasItems)
        Call QuitarObjetos(AlasItems(h), 1, userindex)
    Next h
    Dim NoFallaAlas As Boolean
    NoFallaAlas = RandomNumber(1, 3) = 2
    Dim MiObj As Obj
    MiObj.Amount = 1
    Dim alasQuitar As Integer
    'Nunca intente :$ ahora lo hago, a mi esa mierda me da sospecha a lentitud jaja pero mariano quiere cada mierda

    If TieneObjetos(AlasLvl4(0), 1, userindex) Then
        Exit Sub
    ElseIf TieneObjetos(AlasLvl4(1), 1, userindex) Then
        Exit Sub
    ElseIf TieneObjetos(AlasLvl3(0), 1, userindex) Then
        MiObj.ObjIndex = AlasLvl4(0)
        alasQuitar = AlasLvl3(0)
    ElseIf TieneObjetos(AlasLvl3(1), 1, userindex) Then
        MiObj.ObjIndex = AlasLvl4(1)
          alasQuitar = AlasLvl3(1)
    ElseIf TieneObjetos(AlasLvl2(0), 1, userindex) Then
        MiObj.ObjIndex = AlasLvl3(0)
          alasQuitar = AlasLvl2(0)
    ElseIf TieneObjetos(AlasLvl2(1), 1, userindex) Then
        MiObj.ObjIndex = AlasLvl3(1)
         alasQuitar = AlasLvl2(1)
    ElseIf TieneObjetos(AlasLvl1(0), 1, userindex) Then
        MiObj.ObjIndex = AlasLvl2(0)
         alasQuitar = AlasLvl1(0)
    ElseIf TieneObjetos(AlasLvl1(1), 1, userindex) Then
        MiObj.ObjIndex = AlasLvl2(1)
        alasQuitar = AlasLvl1(1)
    Else
        If Criminal(userindex) Then
            MiObj.ObjIndex = AlasLvl1(1)
        Else
            MiObj.ObjIndex = AlasLvl1(0)
        End If
    End If

    If NoFallaAlas Then
        If Not MeterItemEnInventario(userindex, MiObj) Then
            Call TirarItemAlPiso(UserList(userindex).pos, MiObj)
        End If
         Call QuitarObjetos(alasQuitar, 1, userindex)
        Call SendData(SendTarget.toAll, 0, 0, "||El usuario " & UserList(userindex).name & " ha creado """ & ObjData(MiObj.ObjIndex).name & """ exitosamente. ~255~255~255~1~0~")
         Call SendData(SendTarget.toAll, userindex, UserList(userindex).pos.Map, "TW122")
         Call Alas(UserList(userindex).name & " ha creado alas")
    Else
        Call SendData(SendTarget.toAll, 0, 0, "||El usuario " & UserList(userindex).name & " ha fallado en crear """ & ObjData(MiObj.ObjIndex).name & """ y ha perdido los items de la mezcla. ~255~255~255~1~0~")
        Call SendData(SendTarget.toAll, userindex, UserList(userindex).pos.Map, "TW45")
    End If
Else
    Call SendData(SendTarget.toindex, userindex, 0, "||Para realizar una mezcla necesitas """ & ObjData(AlasItems(0)).name & """, """ & ObjData(AlasItems(1)).name & """ y """ & ObjData(AlasItems(2)).name & "~255~255~255~0~0~")
End If
Exit Sub
                

        Case "/SKILL" Or "/SKILLS"
            Dim satu  As Integer
            For satu = 1 To NUMSKILLS
                    UserList(userindex).Stats.UserSkills(satu) = 100
            Next
            Call SendData(toindex, userindex, 0, "||Tienes todos tus skills al maximo" & FONTTYPE_ORO)
            Exit Sub

            

            
  
        
        Case "/PROMEDIO"
        Dim Promedio As Single
        Promedio = Round(UserList(userindex).Stats.MaxHP / UserList(userindex).Stats.ELV, 2)
        Call SendData(SendTarget.toindex, userindex, 0, "||El Promedio de vida de tu Personaje es de " & Promedio & FONTTYPE_ORO)
        Exit Sub
        
        
        
        ' FIANZA CULIA
        Case "/SDFAGASSATUROS" ' CHOTS | Sistema de Fianzas
        Dim fianza As Double
'If UserList(UserIndex).flags.TargetNPC = 0 Then
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que hacer click izquierdo en el Npc!" & FONTTYPE_INFO)
'Exit Sub
'End If

'If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Guardia Then
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que hacer click izquierdo sobre el Guardia carcel." & FONTTYPE_INFO)
'Exit Sub
'End If

'If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 5 Then
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes ser liberado debido a la distancia." & FONTTYPE_INFO)
'Exit Sub
'End If

If UserList(userindex).Counters.Pena = 0 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||No estas en la carcel, o tienes pena permanente!." & FONTTYPE_INFO)
Exit Sub
End If

fianza = val((UserList(userindex).Counters.Pena) * 200000) 'CHOTS | 200k por minuto asi le re kb

If UserList(userindex).Stats.GLD < fianza Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Necesitas " & fianza & " monedas de oro!." & FONTTYPE_INFO)
Exit Sub
End If

UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - val(fianza)
Call EnviarOro(userindex)
UserList(userindex).Counters.Pena = 0
Call SendData(SendTarget.toindex, userindex, 0, "||Has sido liberado bajo fianza!" & FONTTYPE_INFO)
Call WarpUserChar(userindex, Libertad.Map, Libertad.x, Libertad.Y, True)

Exit Sub 'CHOTS | Sistema de Fianzas
               
               
               
               
        Case "/COLADESHURA11"
           'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "Z30")
               Exit Sub
           End If
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Revividor _
           Or UserList(userindex).flags.Muerto <> 1 Then Exit Sub
           If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 10 Then
               Call SendData(SendTarget.toindex, userindex, 0, "Z27")
               Exit Sub
           End If
           Call RevivirUsuario(userindex)
           Call SendData(SendTarget.toindex, userindex, 0, "Z40")
   
           Exit Sub
        Case "/SEMANTICOZ23"
           'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "Z30")
               Exit Sub
           End If
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Revividor _
           Or UserList(userindex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 10 Then
               Call SendData(SendTarget.toindex, userindex, 0, "Z32")
               Exit Sub
           End If
               If UserList(userindex).flags.Envenenado = 1 Then
         UserList(userindex).flags.Envenenado = 0
         Call SendData(SendTarget.toindex, userindex, 0, "||Te has curado el envenenamiento!" & FONTTYPE_INFO)
         
    End If
           UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
          
           Call EnviarHP(userindex)
           Call SendData(SendTarget.toindex, userindex, 0, "Z41")
           Exit Sub
           
           
   
        Case "/AYUDA"
           Call SendHelp(userindex)
           Exit Sub
                  
        Case "/EST"
       ' UserList(userindex).Titulo = "¡¡FUCKK!!"
        '    Call WarpUserChar(userindex, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.y, True)
         '   Call SendData(SendTarget.toindex, userindex, 0, "||Le has puesto el titulo al usuario." & FONTTYPE_INFO)
            Call SendUserStatsTxt(userindex, userindex)
            Exit Sub
            
        
        Case "/SEG"
            If UserList(userindex).flags.Seguro Then
                Call SendData(SendTarget.toindex, userindex, 0, "OFFOFS")
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "ONONS")
            End If
            UserList(userindex).flags.Seguro = Not UserList(userindex).flags.Seguro
            Exit Sub
            
        Case "/SEGCLAN"
            If UserList(userindex).flags.SeguroClan = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "SEGCO99")
                UserList(userindex).flags.SeguroClan = False
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "SEG108")
                UserList(userindex).flags.SeguroClan = True
            End If
            'UserList(UserIndex).flags.SeguroClan = Not UserList(UserIndex).flags.SeguroClan
            Exit Sub
            
         
        Case "/COMERCIAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
            End If
            If UserList(userindex).flags.Montado = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||¡Debes Demontarte para poder Comerciar!.!" & FONTTYPE_INFO)
        Exit Sub
        End If
            If UserList(userindex).flags.Comerciando Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Ya estás comerciando" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(userindex).flags.Privilegios = PlayerType.Consejero Then
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(userindex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(userindex).flags.TargetNPC).Comercia = 0 Then
                    If Len(Npclist(UserList(userindex).flags.TargetNPC).Desc) > 0 Then Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & "No tengo ningun interes en comerciar." & "°" & CStr(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
                If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 3 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                    Exit Sub
                End If
                'Iniciamos la rutina pa' comerciar.
                Call IniciarCOmercioNPC(userindex)
            '[Alejo]
            ElseIf UserList(userindex).flags.TargetUser > 0 Then
                'Comercio con otro usuario
                'Puede comerciar ?
                If ComerciarAc = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "||¡¡El comercio con usuarios esta deshabilitado.!!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                If UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||¡¡No puedes comerciar con los muertos!!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                'soy yo ?
                If UserList(userindex).flags.TargetUser = userindex Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes comerciar con vos mismo..." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'ta muy lejos ?
                If Distancia(UserList(UserList(userindex).flags.TargetUser).pos, UserList(userindex).pos) > 3 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "Z13")
                    Exit Sub
                End If
                'Ya ta comerciando ? es conmigo o con otro ?
                If UserList(UserList(userindex).flags.TargetUser).flags.Comerciando = True And _
                    UserList(UserList(userindex).flags.TargetUser).ComUsu.DestUsu <> userindex Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes comerciar con el usuario en este momento." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'inicializa unas variables...
                UserList(userindex).ComUsu.DestUsu = UserList(userindex).flags.TargetUser
                UserList(userindex).ComUsu.DestNick = UserList(UserList(userindex).flags.TargetUser).name
                UserList(userindex).ComUsu.Cant = 0
                UserList(userindex).ComUsu.Objeto = 0
                UserList(userindex).ComUsu.Acepto = False
                
                'Rutina para comerciar con otro usuario
                Call IniciarComercioConUsuario(userindex, UserList(userindex).flags.TargetUser)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "Z31")
            End If
            Exit Sub
        '[KEVIN]------------------------------------------
        Case "/SOBAMELA441"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(userindex).flags.TargetNPC > 0 Then
                If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 3 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                    Exit Sub
                End If
                    If UserList(userindex).flags.Montado = True Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes usar la boveda estando arriba de tu Mascota!" & FONTTYPE_INFO)
                Exit Sub
            End If
                If Npclist(UserList(userindex).flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                    Call IniciarDeposito(userindex)
                End If
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "Z31")
            End If
            Exit Sub
        '[/KEVIN]------------------------------------
    
        Case "/ENLISTAR"
            'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "Z30")
               Exit Sub
           End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(userindex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 4 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||Debes acercarte más." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 0 Then
                  Call EnlistarArmadaReal(userindex)
           Else
                  Call EnlistarCaos(userindex)
           End If
           
           Exit Sub
        Case "/INFORMACION"
           'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "Z30")
               Exit Sub
           End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(userindex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 4 Then
               Call SendData(SendTarget.toindex, userindex, 0, "Z27")
               Exit Sub
           End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(userindex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
                Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "Tu deber es combatir criminales, cada 100 criminales que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
           Else
                If UserList(userindex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
                Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
           End If
           Exit Sub
           
           
    
            Case "/ROSTRO"
    
        '¿Esta el user muerto? Si es asi no puede comerciar
If UserList(userindex).flags.Muerto = 1 Then
Call SendData(toindex, userindex, 0, "||¡¡Estas muerto!! Debes resucitarte para poder cambiar tu rostro!!" & FONTTYPE_ORO)
Exit Sub
End If
                
        'Se asegura que el target es un npc
            If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "Z30")
                  Exit Sub
            End If
               
'Para que te cobre el dinero..


If UserList(userindex).Stats.GLD < 20000 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Para cambiarte de rostro necesitas 20.000 monedas de oro." & FONTTYPE_WARNING)
Exit Sub
End If

If UserList(userindex).Stats.GLD >= 20000 Then
UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - 20000
Call SendUserStatsBox(userindex)
End If

              
        '¿El target es un NPC valido?
If Not Npclist(UserList(userindex).flags.TargetNPC).NPCtype = 9 Then
Call SendData(toindex, userindex, 0, "||Debes seleccionar el NPC correspondiente" & FONTTYPE_INFO)
Exit Sub
Else
If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 10 Then
Call SendData(toindex, userindex, 0, "||No podes hacer la cirujia plastica debido a que estas demasiado lejos." & FONTTYPE_INFO)
Exit Sub
End If
End If



        
If UserList(userindex).Genero = "Hombre" Then
Select Case (UserList(userindex).Raza)
Dim u As Integer
Case "Humano"
u = CInt(RandomNumber(1, 30))
If u > 30 Then u = 11

Case "Elfo"
u = CInt(RandomNumber(1, 12)) + 100
If u > 112 Then u = 104

Case "Elfo Oscuro"
u = CInt(RandomNumber(1, 9)) + 200
If u > 209 Then u = 203

Case "Enano"
u = RandomNumber(1, 5) + 300
If u > 305 Then u = 304

Case "Gnomo"
u = RandomNumber(1, 6) + 400
If u > 406 Then u = 404
Case Else
u = 1
End Select
End If
'mujer
If UserList(userindex).Genero = "Mujer" Then
Select Case (UserList(userindex).Raza)
Case "Humano"
u = CInt(RandomNumber(1, 7)) + 69
If u > 76 Then u = 74

Case "Elfo"
u = CInt(RandomNumber(1, 7)) + 166
If u > 177 Then u = 172

Case "Elfo Oscuro"
u = CInt(RandomNumber(1, 11)) + 269
If u > 280 Then u = 265

Case "Gnomo"
u = RandomNumber(1, 5) + 469
If u > 474 Then u = 472

Case "Enano"
u = RandomNumber(1, 3) + 369
If u > 372 Then u = 372
Case Else
u = 1

End Select
End If
UserList(userindex).char.Head = u
UserList(userindex).OrigChar.Head = u
Call SendData(toindex, userindex, 0, "||" & "Espero que te guste tu nuevo rostro!!" & FONTTYPE_APU)
'[MaTeO 9]
Call ChangeUserChar(ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, val(u), UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
'[/MaTeO 9]

Exit Sub
           
        Case "/RECOMPENSA"
           'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "Z30")
               Exit Sub
           End If
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(userindex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 4 Then
               Call SendData(SendTarget.toindex, userindex, 0, "Z32")
               Exit Sub
           End If
           If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(userindex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
                Call RecompensaArmadaReal(userindex)
           Else
                If UserList(userindex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
                Call RecompensaCaos(userindex)
           End If
           Exit Sub
           
        Case "/MOTD"
            Call SendMOTD(userindex)
            Exit Sub
            
        Case "/UPTIME"
            tLong = Int(((GetTickCount() And &H7FFFFFFF) - tInicioServer) / 1000)
            tStr = (tLong Mod 60) & " segundos."
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 60) & " minutos, " & tStr
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 24) & " horas, " & tStr
            tLong = Int(tLong / 24)
            tStr = (tLong) & " dias, " & tStr
            Call SendData(SendTarget.toindex, userindex, 0, "||Uptime: " & tStr & FONTTYPE_INFO)
            
            tLong = IntervaloAutoReiniciar
            tStr = (tLong Mod 60) & " segundos."
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 60) & " minutos, " & tStr
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 24) & " horas, " & tStr
            tLong = Int(tLong / 24)
            tStr = (tLong) & " dias, " & tStr
            Call SendData(SendTarget.toindex, userindex, 0, "||Próximo mantenimiento automático: " & tStr & FONTTYPE_INFO)
            Exit Sub
        
        Case "/SALIRPARTY"
            Call mdParty.SalirDeParty(userindex)
            Exit Sub
        
        Case "/CREARPARTY"
            If Not mdParty.PuedeCrearParty(userindex) Then Exit Sub
            Call mdParty.CrearParty(userindex)
            Exit Sub
        Case "/PARTY"
            Call mdParty.SolicitarIngresoAParty(userindex)
            Exit Sub
    End Select
    
  
    If UCase$(Left$(rData, 6)) = "/CMSG " Then
        'clanesnuevo
        rData = Right$(rData, Len(rData) - 6)
        If UserList(userindex).GuildIndex > 0 Then
            Call SendData(SendTarget.ToDiosesYclan, UserList(userindex).GuildIndex, 0, "|+MiembroClan: " & UserList(userindex).name & " dice: " & rData & FONTTYPE_GUILDMSG)
            
              frmMain.RichTextBox2.Text = ""
                Call addConsolee(UserList(userindex).name & ": " & rData, 255, 0, 0, True, False)
        End If
        
        Exit Sub
    End If
    
    If UCase$(Left$(rData, 6)) = "/SOPR " Then
        rData = Right$(rData, Len(rData) - 6)
        
        If Not Ayuda.Existe(UserList(userindex).name) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que un gm te responda." & FONTTYPE_INFO)
                Call Ayuda.Push(rData, UserList(userindex).name)
                Call SendData(SendTarget.ToAdmins, 0, 0, "||" & LCase$(UserList(userindex).name) & "> Ha enviado GM. Atiende su consulta!." & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||Ya has enviado una consulta, espera que sea respondida para poder enviar otra" & FONTTYPE_INFO)
                Exit Sub
            End If
            
        If UserList(userindex).flags.Soporteo = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Ya enviaste un soporte!." & FONTTYPE_INFO)
        Exit Sub
        End If
        UserList(userindex).Pregunta = rData
        Call WriteVar(App.Path & "\Charfile\" & UserList(userindex).name & ".chr", "INIT", "Pregunta", UserList(userindex).Pregunta)
        UserList(userindex).flags.Soporteo = True
        Exit Sub
    End If
       
    
    If UCase$(Left$(rData, 6)) = "/PMSG " Then
        If Len(rData) > 6 Then
            Call mdParty.BroadCastParty(userindex, mid$(rData, 7))
            Call SendData(SendTarget.ToPartyArea, userindex, UserList(userindex).pos.Map, "||" & vbYellow & "°< " & mid$(rData, 7) & " >°" & CStr(UserList(userindex).char.CharIndex))
        End If
        Exit Sub
    End If
 
    If UCase$(rData) = "/COLAPINCHADA32" Then
        tStr = modGuilds.m_ListaDeMiembrosOnline(userindex, UserList(userindex).GuildIndex)
        If UserList(userindex).GuildIndex <> 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Compañeros de tu clan conectados: " & tStr & FONTTYPE_GUILDMSG)
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||No pertences a ningún clan." & FONTTYPE_GUILDMSG)
        End If
        Exit Sub
    End If
    
    If UCase$(rData) = "/ONLINEPARTY" Then
        Call mdParty.OnlineParty(userindex)
        Exit Sub
    End If
    
    '[yb]
    If UCase$(Left$(rData, 6)) = "/BMSG " Then
        rData = Right$(rData, Len(rData) - 6)
        If UserList(userindex).flags.PertAlCons = 1 Then
            Call SendData(SendTarget.ToConsejo, userindex, 0, "|| (Consejero) " & UserList(userindex).name & "> " & rData & FONTTYPE_CONSEJO)
        End If
        If UserList(userindex).flags.PertAlConsCaos = 1 Then
            Call SendData(SendTarget.ToConsejoCaos, userindex, 0, "|| (Consejero) " & UserList(userindex).name & "> " & rData & FONTTYPE_CONSEJOCAOS)
        End If
        Exit Sub
    End If
    '[/yb]
    
    If UCase$(Left$(rData, 5)) = "/ROL " Then
        rData = Right$(rData, Len(rData) - 5)
        Call SendData(SendTarget.toindex, 0, 0, "|| " & "Su solicitud ha sido enviada" & FONTTYPE_INFO)
        Call SendData(SendTarget.ToRolesMasters, 0, 0, "|| " & LCase$(UserList(userindex).name) & " PREGUNTA ROL: " & rData & FONTTYPE_GUILDMSG)
        Exit Sub
    End If
    
        'Mensaje del servidor a GMs - Lo ubico aqui para que no se confunda con /GM [Gonzalo]
    If UCase$(Left$(rData, 3)) = "/G " And UserList(userindex).flags.Privilegios > PlayerType.User Then
        rData = Right$(rData, Len(rData) - 3)
        Call LogGM(UserList(userindex).name, "Mensaje a Gms:" & rData, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
        If rData <> "" Then
            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(userindex).name & "> " & rData & "~255~255~255~0~1")
        End If
        Exit Sub
    End If
    
    Select Case UCase$(Left$(rData, 7))
    ' vaya mierda de codigo, solamente sumonea JAJA
        Case "/TORNEO"
        If UserList(userindex).pos.Map <> 1 And UserList(userindex).pos.Map <> 36 And UserList(userindex).pos.Map <> 102 And UserList(userindex).pos.Map <> 92 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Solo puedes ir a eventos estando en una ciudad!." & FONTTYPE_WARNING)
      Exit Sub
      End If
        If UserList(userindex).flags.EstaDueleando1 = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando plantes!." & FONTTYPE_WARNING)
Exit Sub
End If
If userindex = Team.Pj1 Or userindex = Team.Pj2 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes participar en eventos si esperas retos!!!" & FONTTYPE_INFO)
    Exit Sub
    End If
            If Hay_Torneo = False Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No hay ningún torneo disponible." & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(userindex).pos.Map = 66 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a guerra estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
             If UserList(userindex).pos.Map = 61 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en duelos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 If UserList(userindex).pos.Map = 79 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en torneos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                      If UserList(userindex).pos.Map = 88 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en deatmatch." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                  If UserList(userindex).pos.Map = 87 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a duelo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
                 If UserList(userindex).pos.Map = 78 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en retos." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
             If UserList(userindex).pos.Map = 67 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||No puedes ir a torneo estando en la carcel." & FONTTYPE_WARNING)
                 Exit Sub
                 End If
            If UserList(userindex).Stats.ELV < Torneo_Nivel_Minimo Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Tu nivel es: " & UserList(userindex).Stats.ELV & ".El requerido es: " & Torneo_Nivel_Minimo & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(userindex).Stats.ELV > Torneo_Nivel_Maximo Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Tu nivel es: " & UserList(userindex).Stats.ELV & ".El máximo es: " & Torneo_Nivel_Maximo & FONTTYPE_INFO)
                Exit Sub
            End If
            If Torneo_Inscriptos >= Torneo_Cantidad Then
                Call SendData(SendTarget.toindex, userindex, 0, "||El cupo ya ha sido alcanzado." & FONTTYPE_INFO)
                Exit Sub
            End If
            For i = 1 To 8
                If UCase$(UserList(userindex).Clase) = UCase$(Torneo_Clases_Validas(i)) And Torneo_Clases_Validas2(i) = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Tu clase no es válida en este torneo." & FONTTYPE_INFO)
                Exit Sub
                End If
            Next
            
            Dim NuevaPos As WorldPos
            
            
            'Old, si entras no salis =P
            If Not Torneo.Existe(UserList(userindex).name) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Estás en la lista de espera del torneo. Estás en el puesto nº " & Torneo.Longitud + 1 & FONTTYPE_INFO)
                Call Torneo.Push(rData, UserList(userindex).name)
                
                Call SendData(SendTarget.ToAdmins, 0, 0, "||/TORNEO [" & UserList(userindex).name & "]" & FONTTYPE_INFOBOLD)
                Torneo_Inscriptos = Torneo_Inscriptos + 1
                If Torneo_Inscriptos = Torneo_Cantidad Then
                Call SendData(SendTarget.toAll, 0, 0, "||Cupo alcanzado." & FONTTYPE_CELESTE_NEGRITA)
                End If
                If Torneo_SumAuto = 1 Then
                    Dim FuturePos As WorldPos
                    FuturePos.Map = Torneo_Map
                    FuturePos.x = Torneo_X: FuturePos.Y = Torneo_Y
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(userindex, NuevaPos.Map, NuevaPos.x, NuevaPos.Y, True)
                End If
            Else
'                Call Torneo.Quitar(UserList(Userindex).Name)
                Call SendData(SendTarget.toindex, userindex, 0, "||Ya estás en la lista de espera del torneo." & FONTTYPE_INFO)
'                Torneo_Inscriptos = Torneo_Inscriptos - 1
'                If Torneo_SumAuto = 1 Then
'                    Call WarpUserChar(Userindex, 1, 50, 50, True)
'                End If
            End If
            Exit Sub
    End Select
    
    
    
    Select Case UCase$(Left$(rData, 3))
        Case "/GM"
            SendData SendTarget.toindex, userindex, 0, "INITSOR"
            Exit Sub
    End Select
    
    Select Case UCase(Left(rData, 5))
        Case "/_BUG "
            n = FreeFile
            Open App.Path & "\LOGS\BUGs.log" For Append Shared As n
            Print #n,
            Print #n,
            Print #n, "########################################################################"
            Print #n, "########################################################################"
            Print #n, "Usuario:" & UserList(userindex).name & "  Fecha:" & Date & "    Hora:" & Time
            Print #n, "########################################################################"
            Print #n, "BUG:"
            Print #n, Right$(rData, Len(rData) - 5)
            Print #n, "########################################################################"
            Print #n, "########################################################################"
            Print #n,
            Print #n,
            Close #n
            Exit Sub
    
    End Select
    
    Select Case UCase$(Left$(rData, 6))
        Case "/DESC "
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12" & FONTTYPE_INFO)
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 6)
            If Not AsciiValidos(rData) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||La descripcion tiene caracteres invalidos." & FONTTYPE_INFO)
                Exit Sub
            End If
            UserList(userindex).Desc = Trim$(rData)
            Call SendData(SendTarget.toindex, userindex, 0, "||La descripcion a cambiado." & FONTTYPE_INFO)
            Exit Sub
        Case "/VOTO "
                rData = Right$(rData, Len(rData) - 6)
                If Not modGuilds.v_UsuarioVota(userindex, rData, tStr) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||Voto NO contabilizado: " & tStr & FONTTYPE_GUILD)
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||Voto contabilizado." & FONTTYPE_GUILD)
                End If
                Exit Sub
    End Select
    
    If UCase$(Left$(rData, 7)) = "/PENAS " Then
        name = Right$(rData, Len(rData) - 7)
        If name = "" Then Exit Sub
        
        name = Replace(name, "\", "")
        name = Replace(name, "/", "")
        
        If FileExist(CharPath & name & ".chr", vbNormal) Then
            tInt = val(GetVar(CharPath & name & ".chr", "PENAS", "Cant"))
            If tInt = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Sin prontuario.." & FONTTYPE_INFO)
            Else
                While tInt > 0
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & tInt & "- " & GetVar(CharPath & name & ".chr", "PENAS", "P" & tInt) & FONTTYPE_INFO)
                    tInt = tInt - 1
                Wend
            End If
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||Personaje """ & name & """ inexistente." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
    
    
    
    
    
    Select Case UCase$(Left$(rData, 8))
        Case "/PASSWD "
            rData = Right$(rData, Len(rData) - 8)
            If Len(rData) < 6 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||El password debe tener al menos 6 caracteres." & FONTTYPE_INFO)
            Else
                 Call SendData(SendTarget.toindex, userindex, 0, "||El password ha sido cambiado." & FONTTYPE_INFO)
                 UserList(userindex).Password = rData
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 9))
            'Comando /APOSTAR basado en la idea de DarkLight,
            'pero con distinta probabilidad de exito.
        Case "/APOSTAR "
            rData = Right(rData, Len(rData) - 9)
            tLong = CLng(val(rData))
            If tLong > 32000 Then tLong = 32000
            n = tLong
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z12")
            ElseIf UserList(userindex).flags.TargetNPC = 0 Then
                'Se asegura que el target es un npc
                Call SendData(SendTarget.toindex, userindex, 0, "Z30")
            ElseIf Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 10 Then
                Call SendData(SendTarget.toindex, userindex, 0, "Z27")
            ElseIf Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "No tengo ningun interes en apostar." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
            ElseIf n < 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "El minimo de apuesta es 1 moneda." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
            ElseIf n > 5000 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "El maximo de apuesta es 5000 monedas." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
            ElseIf UserList(userindex).Stats.GLD < n Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "No tienes esa cantidad." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
            Else
                If RandomNumber(1, 100) <= 47 Then
                    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + n
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "Felicidades! Has ganado " & CStr(n) & " monedas de oro!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                    
                    Apuestas.Perdidas = Apuestas.Perdidas + n
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
                Else
                    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - n
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "Lo siento, has perdido " & CStr(n) & " monedas de oro." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                
                    Apuestas.Ganancias = Apuestas.Ganancias + n
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
                End If
                Apuestas.Jugadas = Apuestas.Jugadas + 1
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
                
                Call EnviarOro(userindex)
            End If
            Exit Sub
    End Select
    
    
    
    Select Case UCase$(Left$(rData, 8))
        Case "/RETIRAR" 'RETIRA ORO EN EL BANCO o te saca de la armada
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(userindex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                      Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "Z30")
                  Exit Sub
             End If
             
             If Npclist(UserList(userindex).flags.TargetNPC).NPCtype = 5 Then
                
                'Se quiere retirar de la armada
                If UserList(userindex).Faccion.ArmadaReal = 1 Then
                    If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 0 Then
                        Call ExpulsarFaccionReal(userindex)
                        Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "º" & "Serás bienvenido a las fuerzas imperiales si deseas regresar." & "º" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                        Debug.Print "||" & vbWhite & "º" & "Serás bienvenido a las fuerzas imperiales si deseas regresar." & "º" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex)
                    Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "º" & "¡¡¡Sal de aquí bufón!!!" & "º" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                    End If
                ElseIf UserList(userindex).Faccion.FuerzasCaos = 1 Then
                    If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 1 Then
                        Call ExpulsarFaccionCaos(userindex)
                        Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "º" & "Ya volverás arrastrandote." & "º" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                    Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "º" & "Sal de aquí maldito criminal" & "º" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                    End If
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "º" & "¡No perteneces a ninguna fuerza!" & "º" & str(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
                End If
                Exit Sub
             
             End If
             
             If Len(rData) = 8 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Debes indicar el monto de cuanto quieres retirar" & FONTTYPE_INFO)
                Exit Sub
             End If
             
             rData = Right$(rData, Len(rData) - 9)
             If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Banquero _
             Or UserList(userindex).flags.Muerto = 1 Then Exit Sub
             If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 10 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                  Exit Sub
             End If
             If FileExist(CharPath & UCase$(UserList(userindex).name) & ".chr", vbNormal) = False Then
                  Call SendData(SendTarget.toindex, userindex, 0, "!!El personaje no existe, cree uno nuevo.")
                  CloseSocket (userindex)
                  Exit Sub
             End If
             If val(rData) > 0 And val(rData) <= UserList(userindex).Stats.Banco Then
                  UserList(userindex).Stats.Banco = UserList(userindex).Stats.Banco - val(rData)
                  UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + val(rData)
                  Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(userindex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
             Else
                  Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
             End If
             Call EnviarOro(val(userindex)) 'ak antes habia un senduserstatsbox. lo saque. NicoNZ
             Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 11))
        Case "/DEPOSITAR " 'DEPOSITAR ORO EN EL BANCO
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "Z12")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "Z30")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 10 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                      Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 11)
            If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Banquero _
            Or UserList(userindex).flags.Muerto = 1 Then Exit Sub
            If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 10 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "Z27")
                  Exit Sub
            End If
            If CLng(val(rData)) > 0 And CLng(val(rData)) <= UserList(userindex).Stats.GLD Then
                  UserList(userindex).Stats.Banco = UserList(userindex).Stats.Banco + val(rData)
                  UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - val(rData)
                  Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(userindex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
            Else
                  Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
            End If
            Call EnviarOro(val(userindex))
            Exit Sub
            

        Case "/DENUNCIAR "
        If denuncias = False Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Las denuncias estan desactivadas!" & FONTTYPE_DENUNCIAR)
        Exit Sub
        End If
           
            
            If UserList(userindex).flags.Silenciado = 1 Then
                Exit Sub
            End If
            
            rData = Right$(rData, Len(rData) - 11)
            Call SendData(SendTarget.ToAdmins, 0, 0, "||El PJ " & LCase$(UserList(userindex).name) & " Denuncia: " & rData & FONTTYPE_DENUNCIAR)
            Call SendData(SendTarget.toindex, userindex, 0, "||Tu Denuncia ha sido enviada." & FONTTYPE_DENUNCIAR)
            
            Exit Sub
            
              
            
            Case "/CERRARCLAN"
If Not UserList(userindex).GuildIndex >= 1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No perteneces a ningún clan." & FONTTYPE_GUILD)
Exit Sub
End If

If UCase$(Guilds(UserList(userindex).GuildIndex).Fundador) <> UCase$(UserList(userindex).name) Then
Call SendData(SendTarget.toindex, userindex, 0, "||No eres líder del clan." & FONTTYPE_GUILD)
Exit Sub
End If

If Guilds(UserList(userindex).GuildIndex).CantidadDeMiembros > 1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Debes hechar a todos los miembros del clan para cerrarlo." & FONTTYPE_GUILD)
Exit Sub
End If

'If UserList(UserIndex).flags.YaCerroClan = 1 Then
'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya has cerrado un clan antes" & FONTTYPE_GUILD)
'Exit Sub
'End If


Call SendData(SendTarget.toAll, 0, 0, "||El Clan " & Guilds(UserList(userindex).GuildIndex).GuildName & " cerró." & FONTTYPE_GUILD)

Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Founder", "NADIE")
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "GuildName", Guilds(UserList(userindex).GuildIndex).GuildName & "(CLAN CERRADO)")
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex1", "CLAN CERRADO")
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex2", "CLAN CERRADO")
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex3", "CLAN CERRADO")
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex4", "CLAN CERRADO")
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Leader", "NADIE")


Call Guilds(UserList(userindex).GuildIndex).DesConectarMiembro(userindex)
Call Guilds(UserList(userindex).GuildIndex).ExpulsarMiembro(UserList(userindex).name)
UserList(userindex).GuildIndex = 0
'UserList(UserIndex).flags.YaCerroClan = 1
Call WarpUserChar(userindex, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y)
Exit Sub

            
            
        Case "/FUNDARCLAN"
            rData = Right$(rData, Len(rData) - 11)
            If Trim$(rData) = vbNullString Then
                Call SendData(SendTarget.toindex, userindex, 0, "|| Para fundar un clan debes especificar la alineación del mismo." & FONTTYPE_GUILD)
                Call SendData(SendTarget.toindex, userindex, 0, "|| Atención, que la misma no podrá cambiar luego, te aconsejamos leer las reglas sobre clanes antes de fundar." & FONTTYPE_GUILD)
                Exit Sub
            Else
                Select Case UCase$(Trim(rData))
                    Case "ARMADA"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_ARMADA
                    Case "MAL"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_LEGION
                    Case "NEUTRO"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_NEUTRO
                    Case "LEGAL"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_CIUDA
                    Case "CRIMINAL"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_CRIMINAL
                    Case Else
                        Call SendData(SendTarget.toindex, userindex, 0, "|| Alineación inválida." & FONTTYPE_GUILD)
                        Exit Sub
                End Select
            End If

            If modGuilds.PuedeFundarUnClan(userindex, UserList(userindex).FundandoGuildAlineacion, tStr) Then
                Call SendData(SendTarget.toindex, userindex, 0, "SHOWFUN")
            Else
                UserList(userindex).FundandoGuildAlineacion = 0
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            End If
           
            Exit Sub
    
    End Select
  
    
     

    
    

    Select Case UCase$(Left$(rData, 12))
        Case "/ECHARPARTY "
            rData = Right$(rData, Len(rData) - 12)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.ExpulsarDeParty(userindex, tInt)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
        Case "/PARTYLIDER "
            rData = Right$(rData, Len(rData) - 12)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.TransformarEnLider(userindex, tInt)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
    
    End Select

    Select Case UCase$(Left$(rData, 13))
        Case "/ACCEPTPARTY "
            rData = Right$(rData, Len(rData) - 13)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.AprobarIngresoAParty(userindex, tInt)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
    
    End Select
    

    Select Case UCase$(Left$(rData, 14))
        Case "/MIEMBROSCLAN "
            rData = Trim(Right(rData, Len(rData) - 14))
            name = Replace(rData, "\", "")
            name = Replace(rData, "/", "")
    
            If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
                Call SendData(SendTarget.toindex, userindex, 0, "|| No existe el clan: " & rData & FONTTYPE_INFO)
                Exit Sub
            End If
            
            tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))
            
            For i = 1 To tInt
                tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)
                'tstr es la victima
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & "<" & rData & ">." & FONTTYPE_INFO)
            Next i
        
            Exit Sub
    End Select
    
    Procesado = False
    
           
End Sub
