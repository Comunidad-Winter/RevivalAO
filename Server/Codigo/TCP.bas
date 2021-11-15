Attribute VB_Name = "TCP"

'Pablo Ignacio Márquez

Option Explicit

'RUTAS DE ENVIO DE DATOS
Public Enum SendTarget
    toIndex = 0         'Envia a un solo User
    toAll = 1           'A todos los Users
    ToMap = 2           'Todos los Usuarios en el mapa
    ToPCArea = 3        'Todos los Users en el area de un user determinado
    ToNone = 4          'Ninguno
    ToAllButIndex = 5   'Todos menos el index
    ToMapButIndex = 6   'Todos en el mapa menos el indice
    ToGM = 7
    ToNPCArea = 8       'Todos los Users en el area de un user determinado
    ToGuildMembers = 9
    ToAdmins = 10
    ToPCAreaButIndex = 11
    ToAdminsAreaButConsejeros = 12
    ToDiosesYclan = 13
    ToConsejo = 14
    ToClanArea = 15
    ToConsejoCaos = 16
    ToRolesMasters = 17
    ToDeadArea = 18
    ToCiudadanos = 19
    ToCriminales = 20
    ToPartyArea = 21
    ToReal = 22
    ToCaos = 23
    ToCiudadanosYRMs = 24
    ToCriminalesYRMs = 25
    ToRealYRMs = 26
    ToCaosYRMs = 27
End Enum


#If UsarQueSocket = 0 Then
' General constants used with most of the controls
Public Const INVALID_HANDLE As Integer = -1
Public Const CONTROL_ERRIGNORE As Integer = 0
Public Const CONTROL_ERRDISPLAY As Integer = 1


' SocietWrench Control Actions
Public Const SOCKET_OPEN As Integer = 1
Public Const SOCKET_CONNECT As Integer = 2
Public Const SOCKET_LISTEN As Integer = 3
Public Const SOCKET_ACCEPT As Integer = 4
Public Const SOCKET_CANCEL As Integer = 5
Public Const SOCKET_FLUSH As Integer = 6
Public Const SOCKET_CLOSE As Integer = 7
Public Const SOCKET_DISCONNECT As Integer = 7
Public Const SOCKET_ABORT As Integer = 8

' SocketWrench Control States
Public Const SOCKET_NONE As Integer = 0
Public Const SOCKET_IDLE As Integer = 1
Public Const SOCKET_LISTENING As Integer = 2
Public Const SOCKET_CONNECTING As Integer = 3
Public Const SOCKET_ACCEPTING As Integer = 4
Public Const SOCKET_RECEIVING As Integer = 5
Public Const SOCKET_SENDING As Integer = 6
Public Const SOCKET_CLOSING As Integer = 7

' Societ Address Families
Public Const AF_UNSPEC As Integer = 0
Public Const AF_UNIX As Integer = 1
Public Const AF_INET As Integer = 2

' Societ Types
Public Const SOCK_STREAM As Integer = 1
Public Const SOCK_DGRAM As Integer = 2
Public Const SOCK_RAW As Integer = 3
Public Const SOCK_RDM As Integer = 4
Public Const SOCK_SEQPACKET As Integer = 5

' Protocol Types
Public Const IPPROTO_IP As Integer = 0
Public Const IPPROTO_ICMP As Integer = 1
Public Const IPPROTO_GGP As Integer = 2
Public Const IPPROTO_TCP As Integer = 6
Public Const IPPROTO_PUP As Integer = 12
Public Const IPPROTO_UDP As Integer = 17
Public Const IPPROTO_IDP As Integer = 22
Public Const IPPROTO_ND As Integer = 77
Public Const IPPROTO_RAW As Integer = 255
Public Const IPPROTO_MAX As Integer = 256


' Network Addpesses
Public Const INADDR_ANY As String = "0.0.0.0"
Public Const INADDR_LOOPBACK As String = "127.0.0.1"
Public Const INADDR_NONE As String = "255.055.255.255"

' Shutdown Values
Public Const SOCKET_READ As Integer = 0
Public Const SOCKET_WRITE As Integer = 1
Public Const SOCKET_READWRITE As Integer = 2

' SocketWrench Error Pesponse
Public Const SOCKET_ERRIGNORE As Integer = 0
Public Const SOCKET_ERRDISPLAY As Integer = 1

' SocketWrench Error Codes
Public Const WSABASEERR As Integer = 24000
Public Const WSAEINTR As Integer = 24004
Public Const WSAEBADF As Integer = 24009
Public Const WSAEACCES As Integer = 24013
Public Const WSAEFAULT As Integer = 24014
Public Const WSAEINVAL As Integer = 24022
Public Const WSAEMFILE As Integer = 24024
Public Const WSAEWOULDBLOCK As Integer = 24035
Public Const WSAEINPROGRESS As Integer = 24036
Public Const WSAEALREADY As Integer = 24037
Public Const WSAENOTSOCK As Integer = 24038
Public Const WSAEDESTADDRREQ As Integer = 24039
Public Const WSAEMSGSIZE As Integer = 24040
Public Const WSAEPROTOTYPE As Integer = 24041
Public Const WSAENOPROTOOPT As Integer = 24042
Public Const WSAEPROTONOSUPPORT As Integer = 24043
Public Const WSAESOCKTNOSUPPORT As Integer = 24044
Public Const WSAEOPNOTSUPP As Integer = 24045
Public Const WSAEPFNOSUPPORT As Integer = 24046
Public Const WSAEAFNOSUPPORT As Integer = 24047
Public Const WSAEADDRINUSE As Integer = 24048
Public Const WSAEADDRNOTAVAIL As Integer = 24049
Public Const WSAENETDOWN As Integer = 24050
Public Const WSAENETUNREACH As Integer = 24051
Public Const WSAENETRESET As Integer = 24052
Public Const WSAECONNABORTED As Integer = 24053
Public Const WSAECONNRESET As Integer = 24054
Public Const WSAENOBUFS As Integer = 24055
Public Const WSAEISCONN As Integer = 24056
Public Const WSAENOTCONN As Integer = 24057
Public Const WSAESHUTDOWN As Integer = 24058
Public Const WSAETOOMANYREFS As Integer = 24059
Public Const WSAETIMEDOUT As Integer = 24060
Public Const WSAECONNREFUSED As Integer = 24061
Public Const WSAELOOP As Integer = 24062
Public Const WSAENAMETOOLONG As Integer = 24063
Public Const WSAEHOSTDOWN As Integer = 24064
Public Const WSAEHOSTUNREACH As Integer = 24065
Public Const WSAENOTEMPTY As Integer = 24066
Public Const WSAEPROCLIM As Integer = 24067
Public Const WSAEUSERS As Integer = 24068
Public Const WSAEDQUOT As Integer = 24069
Public Const WSAESTALE As Integer = 24070
Public Const WSAEREMOTE As Integer = 24071
Public Const WSASYSNOTREADY As Integer = 24091
Public Const WSAVERNOTSUPPORTED As Integer = 24092
Public Const WSANOTINITIALISED As Integer = 24093
Public Const WSAHOST_NOT_FOUND As Integer = 25001
Public Const WSATRY_AGAIN As Integer = 25002
Public Const WSANO_RECOVERY As Integer = 25003
Public Const WSANO_DATA As Integer = 25004
Public Const WSANO_ADDRESS As Integer = 2500
#End If

Sub DarCuerpoYCabeza(ByRef UserBody As Integer, ByRef UserHead As Integer, ByVal Raza As String, ByVal Gen As String)
'TODO: Poner las heads en arrays, así se acceden por índices
'y no hay problemas de discontinuidad de los índices.
'También se debe usar enums para raza y sexo
Select Case Gen
   Case "Hombre"
        Select Case Raza
            Case "Humano"
                UserHead = RandomNumber(1, 30)
                UserBody = 1
            Case "Elfo"
                UserHead = RandomNumber(1, 13) + 100
                If UserHead = 113 Then UserHead = 201       'Un índice no es continuo.... :S muy feo
                UserBody = 2
            Case "Elfo Oscuro"
                UserHead = RandomNumber(1, 8) + 201
                UserBody = 3
            Case "Enano"
                UserHead = RandomNumber(1, 5) + 300
                UserBody = 52
            Case "Gnomo"
                UserHead = RandomNumber(1, 6) + 400
                UserBody = 52
            Case Else
                UserHead = 1
                UserBody = 1
        End Select
   Case "Mujer"
        Select Case Raza
            Case "Humano"
                UserHead = RandomNumber(1, 7) + 69
                UserBody = 1
            Case "Elfo"
                UserHead = RandomNumber(1, 7) + 169
                UserBody = 2
            Case "Elfo Oscuro"
                UserHead = RandomNumber(1, 11) + 269
                UserBody = 3
            Case "Gnomo"
                UserHead = RandomNumber(1, 5) + 469
                UserBody = 52
            Case "Enano"
                UserHead = RandomNumber(1, 3) + 369
                UserBody = 52
            Case Else
                UserHead = 70
                UserBody = 1
        End Select
End Select

End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True

End Function

Function Numeric(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 48 Or car > 57) Then
        Numeric = False
        Exit Function
    End If
    
Next i

Numeric = True

End Function


Function NombrePermitido(ByVal nombre As String) As Boolean
Dim i As Integer

For i = 1 To UBound(ForbidenNames)
    If InStr(nombre, ForbidenNames(i)) Then
            NombrePermitido = False
            Exit Function
    End If
Next i

NombrePermitido = True

End Function

Function ValidateSkills(ByVal userindex As Integer) As Boolean

Dim LoopC As Integer

For LoopC = 1 To NUMSKILLS
    If UserList(userindex).Stats.UserSkills(LoopC) < 0 Then
        Exit Function
        If UserList(userindex).Stats.UserSkills(LoopC) > 100 Then UserList(userindex).Stats.UserSkills(LoopC) = 100
    End If
Next LoopC

ValidateSkills = True
    
End Function

'Barrin 3/3/03
'Agregué PadrinoName y Padrino password como opcionales, que se les da un valor siempre y cuando el servidor esté usando el sistema
Sub ConnectNewUser(userindex As Integer, name As String, Password As String, UserRaza As String, UserSexo As String, UserClase As String, _
                    US1 As String, US2 As String, US3 As String, US4 As String, US5 As String, _
                    US6 As String, US7 As String, US8 As String, US9 As String, US10 As String, _
                    US11 As String, US12 As String, US13 As String, US14 As String, US15 As String, _
                    US16 As String, US17 As String, US18 As String, US19 As String, US20 As String, _
                    US21 As String, UserEmail As String, Hogar As String)

If Not AsciiValidos(name) Then
    Call SendData(SendTarget.toIndex, userindex, 0, "ERRNombre invalido.")
    Exit Sub
End If

Dim LoopC As Integer
Dim totalskpts As Long

'¿Existe el personaje?
If FileExist(CharPath & UCase$(name) & ".chr", vbNormal) = True Then
    Call SendData(SendTarget.toIndex, userindex, 0, "ERRYa existe el personaje.")
    Exit Sub
End If

'Tiró los dados antes de llegar acá??
If UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = 0 Then
    Call SendData(SendTarget.toIndex, userindex, 0, "ERRDebe tirar los dados antes de poder crear un personaje.")
    Exit Sub
End If

UserList(userindex).flags.Muerto = 0
UserList(userindex).flags.Escondido = 0



UserList(userindex).Reputacion.AsesinoRep = 0
UserList(userindex).Reputacion.BandidoRep = 0
UserList(userindex).Reputacion.BurguesRep = 0
UserList(userindex).Reputacion.LadronesRep = 0
UserList(userindex).Reputacion.NobleRep = 1000
UserList(userindex).Reputacion.PlebeRep = 30

UserList(userindex).Reputacion.Promedio = 30 / 6


UserList(userindex).name = name
UserList(userindex).Clase = UserClase
UserList(userindex).Raza = UserRaza
UserList(userindex).Genero = UserSexo
UserList(userindex).email = UserEmail
UserList(userindex).Hogar = Hogar

Select Case UCase$(UserRaza)
    Case "HUMANO"
        UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) + 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) + 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) + 2
        UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) + 1
    Case "ELFO"
        UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) + 4
        UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) + 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) + 2
        UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) + 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) + 2
    Case "ELFO OSCURO"
        UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) + 2
        UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) + 2
        UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) + 2
        UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) + 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) - 3
    Case "ENANO"
        UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) + 3
        UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) + 3
        UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) - 5
        UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) + 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) - 2
    Case "GNOMO"
        UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) + 1
        UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) + 3
        UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) + 3
        UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) + 1
End Select



UserList(userindex).Stats.UserSkills(1) = val(US1)
UserList(userindex).Stats.UserSkills(2) = val(US2)
UserList(userindex).Stats.UserSkills(3) = val(US3)
UserList(userindex).Stats.UserSkills(4) = val(US4)
UserList(userindex).Stats.UserSkills(5) = val(US5)
UserList(userindex).Stats.UserSkills(6) = val(US6)
UserList(userindex).Stats.UserSkills(7) = val(US7)
UserList(userindex).Stats.UserSkills(8) = val(US8)
UserList(userindex).Stats.UserSkills(9) = val(US9)
UserList(userindex).Stats.UserSkills(10) = val(US10)
UserList(userindex).Stats.UserSkills(11) = val(US11)
UserList(userindex).Stats.UserSkills(12) = val(US12)
UserList(userindex).Stats.UserSkills(13) = val(US13)
UserList(userindex).Stats.UserSkills(14) = val(US14)
UserList(userindex).Stats.UserSkills(15) = val(US15)
UserList(userindex).Stats.UserSkills(16) = val(US16)
UserList(userindex).Stats.UserSkills(17) = val(US17)
UserList(userindex).Stats.UserSkills(18) = val(US18)
UserList(userindex).Stats.UserSkills(19) = val(US19)
UserList(userindex).Stats.UserSkills(20) = val(US20)
UserList(userindex).Stats.UserSkills(21) = val(US21)

totalskpts = 0

'Abs PREVINENE EL HACKEO DE LOS SKILLS %%%%%%%%%%%%%
For LoopC = 1 To NUMSKILLS
    totalskpts = totalskpts + Abs(UserList(userindex).Stats.UserSkills(LoopC))
Next LoopC



If totalskpts > 10 Then
    Call LogHackAttemp(UserList(userindex).name & " intento hackear los skills.")
    Call BorrarUsuario(UserList(userindex).name)
    Call CloseSocket(userindex)
    Exit Sub
End If
'%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%

UserList(userindex).Password = Password
UserList(userindex).char.Heading = eHeading.SOUTH

Call DarCuerpoYCabeza(UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).Raza, UserList(userindex).Genero)
UserList(userindex).OrigChar = UserList(userindex).char
   
 
UserList(userindex).char.WeaponAnim = NingunArma
UserList(userindex).char.ShieldAnim = NingunEscudo
UserList(userindex).char.CascoAnim = NingunCasco
'[MaTeO 9]
UserList(userindex).char.Alas = NingunAlas
'[/MaTeO 9]

UserList(userindex).Stats.MET = 1
Dim MiInt As Long
MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 3)

UserList(userindex).Stats.MaxHP = 15 + MiInt
UserList(userindex).Stats.MinHP = 15 + MiInt

MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) \ 6)
If MiInt = 1 Then MiInt = 2

UserList(userindex).Stats.MaxSta = 20 * MiInt
UserList(userindex).Stats.MinSta = 20 * MiInt


UserList(userindex).Stats.MaxAGU = 100
UserList(userindex).Stats.MinAGU = 100
UserList(userindex).Stats.TrofOro = 0
UserList(userindex).Stats.TrofPlata = 0
UserList(userindex).Stats.TrofBronce = 0
UserList(userindex).Stats.MaxHam = 100
UserList(userindex).Stats.MinHam = 100

' puntos
UserList(userindex).Stats.PuntosDeath = 0
UserList(userindex).Stats.PuntosDuelos = 0
UserList(userindex).Stats.PuntosTorneo = 0
UserList(userindex).Stats.PuntosRetos = 0
UserList(userindex).Stats.PuntosPlante = 0
UserList(userindex).Stats.PuntosCanje = 0

' soporte
UserList(userindex).Pregunta = "Ninguna"
UserList(userindex).Respuesta = "Ninguna"
'<-----------------MANA----------------------->
If UCase$(UserClase) = "MAGO" Then
    MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)) / 3
    UserList(userindex).Stats.MaxMAN = 100 + MiInt
    UserList(userindex).Stats.MinMAN = 100 + MiInt
ElseIf UCase$(UserClase) = "CLERIGO" Or UCase$(UserClase) = "DRUIDA" _
    Or UCase$(UserClase) = "BARDO" Or UCase$(UserClase) = "ASESINO" Then
        MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)) / 4
        UserList(userindex).Stats.MaxMAN = 50
        UserList(userindex).Stats.MinMAN = 50
Else
    UserList(userindex).Stats.MaxMAN = 0
    UserList(userindex).Stats.MinMAN = 0
End If

If UCase$(UserClase) = "MAGO" Or UCase$(UserClase) = "CLERIGO" Or _
   UCase$(UserClase) = "DRUIDA" Or UCase$(UserClase) = "BARDO" Or _
   UCase$(UserClase) = "ASESINO" Then
        UserList(userindex).Stats.UserHechizos(1) = 2
End If

UserList(userindex).Stats.MaxHIT = 2
UserList(userindex).Stats.MinHIT = 1

UserList(userindex).Stats.GLD = 0

Dim Skills


UserList(userindex).Stats.Exp = 0
UserList(userindex).Stats.ELU = 300
UserList(userindex).Stats.ELV = 1

'???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿
UserList(userindex).Invent.NroItems = 4

UserList(userindex).Invent.Object(1).ObjIndex = 467
UserList(userindex).Invent.Object(1).Amount = 100

UserList(userindex).Invent.Object(2).ObjIndex = 468
UserList(userindex).Invent.Object(2).Amount = 100

UserList(userindex).Invent.Object(3).ObjIndex = 460
UserList(userindex).Invent.Object(3).Amount = 1
UserList(userindex).Invent.Object(3).Equipped = 1

Select Case UserRaza
    Case "Humano"
        UserList(userindex).Invent.Object(4).ObjIndex = 463
    Case "Elfo"
        UserList(userindex).Invent.Object(4).ObjIndex = 464
    Case "Elfo Oscuro"
        UserList(userindex).Invent.Object(4).ObjIndex = 465
    Case "Enano"
        UserList(userindex).Invent.Object(4).ObjIndex = 466
    Case "Gnomo"
        UserList(userindex).Invent.Object(4).ObjIndex = 466
End Select

UserList(userindex).Invent.Object(4).Amount = 1
UserList(userindex).Invent.Object(4).Equipped = 1

UserList(userindex).Invent.ArmourEqpSlot = 4
UserList(userindex).Invent.ArmourEqpObjIndex = UserList(userindex).Invent.Object(4).ObjIndex

UserList(userindex).Invent.WeaponEqpObjIndex = UserList(userindex).Invent.Object(3).ObjIndex
UserList(userindex).Invent.WeaponEqpSlot = 3



Call SaveUser(userindex, CharPath & UCase$(name) & ".chr")
  
'Open User
Call ConnectUser(userindex, name, Password)
  
End Sub

#If UsarQueSocket = 1 Or UsarQueSocket = 2 Then

Sub CloseSocket(ByVal userindex As Integer, Optional ByVal cerrarlo As Boolean = True)
Dim LoopC As Integer

On Error GoTo errhandler

    If userindex = LastUser Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser < 1 Then Exit Do
        Loop
    End If
    
If UserList(userindex).flags.automatico = True Then
Call Rondas_UsuarioDesconecta(userindex)
End If
If UserList(userindex).pos.Map = 79 And UserList(userindex).flags.automatico = False Then
Call WarpUserChar(userindex, 1, 45, 49, True)
End If
If UserList(userindex).flags.death = True Then
Call death_muere(userindex)
End If
If UserList(userindex).flags.bandas = True Then
Call Ban_Desconecta(userindex)
End If
If UserList(userindex).flags.EnDosVDos = True Then
    Call CerroEnDuelo(userindex)
End If
If UserList(userindex).flags.Montado = True Then
UserList(userindex).char.Body = UserList(userindex).flags.NumeroMont
'[MaTeO 9]
Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).char.Body, UserList(userindex).char.Head, UserList(userindex).char.Heading, UserList(userindex).char.WeaponAnim, UserList(userindex).char.ShieldAnim, UserList(userindex).char.CascoAnim, UserList(userindex).char.Alas)
'[/MaTeO 9]
UserList(userindex).flags.NumeroMont = 0
UserList(userindex).flags.Montado = False
End If

    
     If UserList(userindex).pos.Map = 61 And userindex = duelosespera Then
Call WarpUserChar(userindex, 1, 50, 50, True)
Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Duelos: " & UserList(duelosespera).name & " Ha salido de la sala de duelos" & FONTTYPE_TALK)
duelosespera = duelosreta
numduelos = 0
End If

If UserList(userindex).pos.Map = 61 And userindex = duelosreta Then
Call WarpUserChar(userindex, 1, 50, 50, True)
Call SendData(SendTarget.ToAllButIndex, 0, 0, "||Duelos: " & UserList(duelosreta).name & " Ha salido de la sala de duelos" & FONTTYPE_TALK)
End If

    If UserList(userindex).pos.Map = 76 Then
Call WarpUserChar(userindex, 1, 50, 50, True)
End If
  If UserList(userindex).pos.Map = 117 Then
Call WarpUserChar(userindex, 1, 50, 50, True)
UserList(userindex).Counters.maparql = 0
End If
If userindex = Team.Pj1 Or userindex = Team.Pj2 Then
    Team.SonDos = False
    Team.Pj1 = 0
    Team.Pj2 = 0
End If

    If UserList(userindex).flags.EstaDueleando = True Then
    Call DesconectarDuelo(UserList(userindex).flags.Oponente, userindex)
    End If
     If UserList(userindex).flags.EstaDueleando1 = True Then
    Call DesconectarDueloPlantes(UserList(userindex).flags.Oponente1, userindex)
    End If
'////////////////////////////////////////////////////////////////////////////////////////
    'Call SecurityIp.IpRestarConexion(GetLongIp(UserList(UserIndex).ip))
    
    If UserList(userindex).ConnID <> -1 Then
        Call CloseSocketSL(userindex)
    End If
    
    'Es el mismo user al que está revisando el centinela??
    'IMPORTANTE!!! hacerlo antes de resetear así todavía sabemos el nombre del user
    ' y lo podemos loguear
    
    'mato los comercios seguros
    If UserList(userindex).ComUsu.DestUsu > 0 Then
        If UserList(UserList(userindex).ComUsu.DestUsu).flags.UserLogged Then
            If UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.DestUsu = userindex Then
                Call SendData(SendTarget.toIndex, UserList(userindex).ComUsu.DestUsu, 0, "||Comercio cancelado por el otro usuario" & FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(userindex).ComUsu.DestUsu)
            End If
        End If
    End If
    
    If UserList(userindex).flags.UserLogged Then
        If NumUsers > 0 Then NumUsers = NumUsers - 1
        Call CloseUser(userindex)
        Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    Else
        Call ResetUserSlot(userindex)
    End If
    Call SendData(toAll, 0, 0, "³" & NumUsers)
    UserList(userindex).flags.EnDosVDos = False
    UserList(userindex).flags.envioSol = False
    UserList(userindex).flags.RecibioSol = False
    UserList(userindex).flags.ParejaMuerta = False
    UserList(userindex).flags.EsperandoDuelo1 = False
    UserList(userindex).flags.Oponente1 = 0
    UserList(userindex).flags.EstaDueleando1 = False
    UserList(userindex).flags.EsperandoDuelo = False
    UserList(userindex).flags.Oponente = 0
    UserList(userindex).flags.EstaDueleando = False
    UserList(userindex).ConnID = -1
    UserList(userindex).ConnIDValida = False
    UserList(userindex).NumeroPaquetesPorMiliSec = 0
    
Exit Sub

errhandler:
    UserList(userindex).ConnID = -1
    UserList(userindex).ConnIDValida = False
    UserList(userindex).NumeroPaquetesPorMiliSec = 0
    Call ResetUserSlot(userindex)
    
#If UsarQueSocket = 1 Then
    If UserList(userindex).ConnID <> -1 Then
        Call CloseSocketSL(userindex)
    End If
#End If

    Call LogError("CloseSocket - Error = " & Err.Number & " - Descripción = " & Err.Description & " - UserIndex = " & userindex)
End Sub

#ElseIf UsarQueSocket = 0 Then

Sub CloseSocket(ByVal userindex As Integer)
On Error GoTo errhandler
    
    
    
    UserList(userindex).ConnID = -1
    UserList(userindex).NumeroPaquetesPorMiliSec = 0

    If userindex = LastUser And LastUser > 1 Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop
    End If

If UserList(userindex).flags.UserLogged Then
    If NumUsers <> 0 Then NumUsers = NumUsers - 1
    Call CloseUser(userindex)
    Call SendData(toAll, 0, 0, "³" & NumUsers)
    End If
Call SendData(toAll, 0, 0, "³" & NumUsers)
    frmMain.Socket2(userindex).Cleanup
    Unload frmMain.Socket2(userindex)
    Call ResetUserSlot(userindex)

Exit Sub

errhandler:
    UserList(userindex).ConnID = -1
    UserList(userindex).NumeroPaquetesPorMiliSec = 0
    Call ResetUserSlot(userindex)
End Sub







#ElseIf UsarQueSocket = 3 Then

Sub CloseSocket(ByVal userindex As Integer, Optional ByVal cerrarlo As Boolean = True)

On Error GoTo errhandler

Dim NURestados As Boolean
Dim CoNnEcTiOnId As Long


    NURestados = False
    CoNnEcTiOnId = UserList(userindex).ConnID
    
    'call logindex(UserIndex, "******> Sub CloseSocket. ConnId: " & CoNnEcTiOnId & " Cerrarlo: " & Cerrarlo)
    
    
  
    UserList(userindex).ConnID = -1 'inabilitamos operaciones en socket
    UserList(userindex).NumeroPaquetesPorMiliSec = 0

    If userindex = LastUser And LastUser > 1 Then
        Do
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop While UserList(LastUser).ConnID = -1
    End If

   If UserList(userindex).flags.UserLogged Then
   If NumUsers <> 0 Then NumUsers = NumUsers - 1
   NURestados = True
   Call CloseUser(userindex)
    Call SendData(toAll, 0, 0, "³" & NumUsers)
   End If
    Call SendData(toAll, 0, 0, "³" & NumUsers)
    Call ResetUserSlot(userindex)
    Call SendData(toAll, 0, 0, "³" & NumUsers)
    'limpiada la userlist... reseteo el socket, si me lo piden
    'Me lo piden desde: cerrada intecional del servidor (casi todas
    'las llamadas a CloseSocket del codigo)
    'No me lo piden desde: disconnect remoto (el on_close del control
    'de alejo realiza la desconexion automaticamente). Esto puede pasar
    'por ejemplo, si el cliente cierra el AO.
    If cerrarlo Then Call frmMain.TCPServ.CerrarSocket(CoNnEcTiOnId)

Exit Sub

errhandler:
    UserList(userindex).NumeroPaquetesPorMiliSec = 0
    Call LogError("CLOSESOCKETERR: " & Err.Description & " UI:" & userindex)
    
    If Not NURestados Then
        If UserList(userindex).flags.UserLogged Then
            If NumUsers > 0 Then
                NumUsers = NumUsers - 1
            End If
            Call LogError("Cerre sin grabar a: " & UserList(userindex).name)
        End If
    End If
    
    Call LogError("El usuario no guardado tenia connid " & CoNnEcTiOnId & ". Socket no liberado.")
    Call ResetUserSlot(userindex)

End Sub


#End If

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal userindex As Integer)

#If UsarQueSocket = 1 Then

If UserList(userindex).ConnID <> -1 And UserList(userindex).ConnIDValida Then
    Call BorraSlotSock(UserList(userindex).ConnID)
    Call WSApiCloseSocket(UserList(userindex).ConnID)
    UserList(userindex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 0 Then

If UserList(userindex).ConnID <> -1 And UserList(userindex).ConnIDValida Then
    frmMain.Socket2(userindex).Cleanup
    Unload frmMain.Socket2(userindex)
    UserList(userindex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 2 Then

If UserList(userindex).ConnID <> -1 And UserList(userindex).ConnIDValida Then
    Call frmMain.Serv.CerrarSocket(UserList(userindex).ConnID)
    UserList(userindex).ConnIDValida = False
End If

#End If
End Sub

Public Function EnviarDatosASlot(ByVal userindex As Integer, Datos As String) As Long

#If UsarQueSocket = 1 Then '**********************************************
    On Error GoTo Err
    
    Dim Ret As Long
    
   ' Datos = Encode64(EncryptStr(Datos, "xaopepe"))
    
    
    Ret = WsApiEnviar(userindex, Datos)
    
    If Ret <> 0 And Ret <> WSAEWOULDBLOCK Then
        Call CloseSocketSL(userindex)
        Call Cerrar_Usuario(userindex)
    End If
    EnviarDatosASlot = Ret
    Exit Function
    
Err:
        'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("EnviarDatosASlot:: ERR Handler. userindex=" & UserIndex & " datos=" & Datos & " UL?/CId/CIdV?=" & UserList(UserIndex).flags.UserLogged & "/" & UserList(UserIndex).ConnID & "/" & UserList(UserIndex).ConnIDValida & " ERR: " & Err.Description)

#ElseIf UsarQueSocket = 0 Then '**********************************************

    Dim Encolar As Boolean
    Encolar = False
    
    EnviarDatosASlot = 0
    
    If UserList(userindex).ColaSalida.Count <= 0 Then
        If frmMain.Socket2(userindex).Write(Datos, Len(Datos)) < 0 Then
            If frmMain.Socket2(userindex).LastError = WSAEWOULDBLOCK Then
                UserList(userindex).SockPuedoEnviar = False
                Encolar = True
            Else
                Call Cerrar_Usuario(userindex)
            End If
        End If
    Else
        Encolar = True
    End If
    
    If Encolar Then
        Debug.Print "Encolando..."
        UserList(userindex).ColaSalida.Add Datos
    End If

#ElseIf UsarQueSocket = 2 Then '**********************************************

Dim Encolar As Boolean
Dim Ret As Long
    
    Encolar = False
    
    '//
    '// Valores de retorno:
    '//                     0: Todo OK
    '//                     1: WSAEWOULDBLOCK
    '//                     2: Error critico
    '//
    If UserList(userindex).ColaSalida.Count <= 0 Then
        Ret = frmMain.Serv.Enviar(UserList(userindex).ConnID, Datos, Len(Datos))
        If Ret = 1 Then
            Encolar = True
        ElseIf Ret = 2 Then
            Call CloseSocketSL(userindex)
            Call Cerrar_Usuario(userindex)
        End If
    Else
        Encolar = True
    End If
    
    If Encolar Then
        Debug.Print "Encolando..."
        UserList(userindex).ColaSalida.Add Datos
    End If

#ElseIf UsarQueSocket = 3 Then
    Dim rv As Long
    'al carajo, esto encola solo!!! che, me aprobará los
    'parciales también?, este control hace todo solo!!!!
    On Error GoTo ErrorHandler
        
        If UserList(userindex).ConnID = -1 Then
            Call LogError("TCP::EnviardatosASlot, se intento enviar datos a un userIndex con ConnId=-1")
            Exit Function
        End If
        
        If frmMain.TCPServ.Enviar(UserList(userindex).ConnID, Datos, Len(Datos)) = 2 Then Call CloseSocket(userindex, True)

Exit Function
ErrorHandler:
    Call LogError("TCP::EnviarDatosASlot. UI/ConnId/Datos: " & userindex & "/" & UserList(userindex).ConnID & "/" & Datos)
#End If '**********************************************

End Function

Sub SendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal sndData As String)

On Error Resume Next

Dim LoopC As Integer
Dim x As Integer
Dim Y As Integer
sndData = AoDefEncode(AoDefServEncrypt(sndData))
sndData = sndData & ENDC

Select Case sndRoute

    Case SendTarget.ToPCArea
        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If MapData(sndMap, x, Y).userindex > 0 Then
                       If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, sndData)
                       End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub
    
    Case SendTarget.toIndex
        If UserList(sndIndex).ConnID <> -1 Then
            Call EnviarDatosASlot(sndIndex, sndData)
            Exit Sub
        End If


    Case SendTarget.ToNone
        Exit Sub
        
        
    Case SendTarget.ToAdmins
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID <> -1 Then
                If UserList(LoopC).flags.Privilegios > 0 Then
                    Call EnviarDatosASlot(LoopC, sndData)
               End If
            End If
        Next LoopC
        Exit Sub
        
    Case SendTarget.toAll
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID <> -1 Then
                If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToAllButIndex
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) And (LoopC <> sndIndex) Then
                If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToMap
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).pos.Map = sndMap Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            End If
        Next LoopC
        Exit Sub
      
    Case SendTarget.ToMapButIndex
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) And LoopC <> sndIndex Then
                If UserList(LoopC).pos.Map = sndMap Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
            
    Case SendTarget.ToGuildMembers
        
        LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        While LoopC > 0
            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, sndData)
            End If
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        Wend
        
        Exit Sub


    Case SendTarget.ToDeadArea
        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If MapData(sndMap, x, Y).userindex > 0 Then
                        If UserList(MapData(sndMap, x, Y).userindex).flags.Muerto = 1 Or UserList(MapData(sndMap, x, Y).userindex).flags.Privilegios >= 1 Then
                           If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                                Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, sndData)
                           End If
                        End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub

    '[Alejo-18-5]
    Case SendTarget.ToPCAreaButIndex
        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If (MapData(sndMap, x, Y).userindex > 0) And (MapData(sndMap, x, Y).userindex <> sndIndex) Then
                       If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, sndData)
                       End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub
       
    Case SendTarget.ToClanArea
        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If (MapData(sndMap, x, Y).userindex > 0) Then
                        If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                            If UserList(sndIndex).GuildIndex > 0 And UserList(MapData(sndMap, x, Y).userindex).GuildIndex = UserList(sndIndex).GuildIndex Then
                                Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, sndData)
                            End If
                        End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub



    Case SendTarget.ToPartyArea
        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If (MapData(sndMap, x, Y).userindex > 0) Then
                        If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                            If UserList(sndIndex).PartyIndex > 0 And UserList(MapData(sndMap, x, Y).userindex).PartyIndex = UserList(sndIndex).PartyIndex Then
                                Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, sndData)
                            End If
                        End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub
        
    '[CDT 17-02-2004]
    Case SendTarget.ToAdminsAreaButConsejeros
        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If (MapData(sndMap, x, Y).userindex > 0) And (MapData(sndMap, x, Y).userindex <> sndIndex) Then
                       If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                            If UserList(MapData(sndMap, x, Y).userindex).flags.Privilegios > 1 Then
                                Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, sndData)
                            End If
                       End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub
    '[/CDT]

    Case SendTarget.ToNPCArea
        For Y = Npclist(sndIndex).pos.Y - MinYBorder + 1 To Npclist(sndIndex).pos.Y + MinYBorder - 1
            For x = Npclist(sndIndex).pos.x - MinXBorder + 1 To Npclist(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If MapData(sndMap, x, Y).userindex > 0 Then
                       If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, sndData)
                       End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub

    Case SendTarget.ToDiosesYclan
        LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        While LoopC > 0
            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, sndData)
            End If
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        Wend

        LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
        While LoopC > 0
            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, sndData)
            End If
            LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
        Wend

        Exit Sub

    Case SendTarget.ToConsejo
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.PertAlCons > 0 Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    Case SendTarget.ToConsejoCaos
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.PertAlConsCaos > 0 Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    Case SendTarget.ToRolesMasters
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToCiudadanos
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If Not Criminal(LoopC) Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToCriminales
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If Criminal(LoopC) Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToReal
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).Faccion.ArmadaReal = 1 Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToCaos
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).Faccion.FuerzasCaos = 1 Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
        
    Case ToCiudadanosYRMs
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If Not Criminal(LoopC) Or UserList(LoopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case ToCriminalesYRMs
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If Criminal(LoopC) Or UserList(LoopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case ToRealYRMs
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).Faccion.ArmadaReal = 1 Or UserList(LoopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case ToCaosYRMs
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).Faccion.FuerzasCaos = 1 Or UserList(LoopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
End Select

End Sub

#If SeguridadAlkon Then

Sub SendCryptedMoveChar(ByVal Map As Integer, ByVal userindex As Integer, ByVal x As Integer, ByVal Y As Integer)
Dim LoopC As Integer

    For LoopC = 1 To LastUser
        If UserList(LoopC).pos.Map = Map Then
            If LoopC <> userindex Then
                If (UserList(LoopC).ConnID <> -1) Then
                    Call EnviarDatosASlot(LoopC, "+" & Encriptacion.MoveCharCrypt(LoopC, UserList(userindex).char.CharIndex, x, Y) & ENDC)
                End If
            End If
        End If
    Next LoopC
    Exit Sub
    

End Sub

Sub SendCryptedData(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal sndData As String)
'No puse un optional parameter en senddata porque no estoy seguro
'como afecta la performance un parametro opcional
'Prefiero 1K mas de exe que arriesgar performance
On Error Resume Next

Dim LoopC As Integer
Dim x As Integer
Dim Y As Integer


Select Case sndRoute


    Case SendTarget.ToNone
        Exit Sub
        
    Case SendTarget.ToAdmins
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID <> -1 Then
'               If EsDios(UserList(LoopC).Name) Or EsSemiDios(UserList(LoopC).Name) Then
                If UserList(LoopC).flags.Privilegios > 0 Then
                    Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
               End If
            End If
        Next LoopC
        Exit Sub
        
    Case SendTarget.toAll
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID <> -1 Then
                If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                    Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToAllButIndex
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) And (LoopC <> sndIndex) Then
                If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                    Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToMap
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).pos.Map = sndMap Then
                        Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
                    End If
                End If
            End If
        Next LoopC
        Exit Sub
      
    Case SendTarget.ToMapButIndex
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) And LoopC <> sndIndex Then
                If UserList(LoopC).pos.Map = sndMap Then
                    Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToGuildMembers
    
        LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        While LoopC > 0
            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
            End If
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        Wend
        
        Exit Sub
    
    Case SendTarget.ToPCArea
        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If MapData(sndMap, x, Y).userindex > 0 Then
                       If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, ProtoCrypt(sndData, MapData(sndMap, x, Y).userindex) & ENDC)
                       End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub

    '[Alejo-18-5]
    Case SendTarget.ToPCAreaButIndex
        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If (MapData(sndMap, x, Y).userindex > 0) And (MapData(sndMap, x, Y).userindex <> sndIndex) Then
                       If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, ProtoCrypt(sndData, MapData(sndMap, x, Y).userindex) & ENDC)
                       End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub

    '[CDT 17-02-2004]
    Case SendTarget.ToAdminsAreaButConsejeros
        For Y = UserList(sndIndex).pos.Y - MinYBorder + 1 To UserList(sndIndex).pos.Y + MinYBorder - 1
            For x = UserList(sndIndex).pos.x - MinXBorder + 1 To UserList(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If (MapData(sndMap, x, Y).userindex > 0) And (MapData(sndMap, x, Y).userindex <> sndIndex) Then
                       If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                            If UserList(MapData(sndMap, x, Y).userindex).flags.Privilegios > 1 Then
                                Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, ProtoCrypt(sndData, MapData(sndMap, x, Y).userindex) & ENDC)
                            End If
                       End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub
    '[/CDT]

    Case SendTarget.ToNPCArea
        For Y = Npclist(sndIndex).pos.Y - MinYBorder + 1 To Npclist(sndIndex).pos.Y + MinYBorder - 1
            For x = Npclist(sndIndex).pos.x - MinXBorder + 1 To Npclist(sndIndex).pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, Y) Then
                    If MapData(sndMap, x, Y).userindex > 0 Then
                       If UserList(MapData(sndMap, x, Y).userindex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, x, Y).userindex, ProtoCrypt(sndData, MapData(sndMap, x, Y).userindex) & ENDC)
                       End If
                    End If
               End If
            Next x
        Next Y
        Exit Sub

    Case SendTarget.toIndex
        If UserList(sndIndex).ConnID <> -1 Then
             Call EnviarDatosASlot(sndIndex, ProtoCrypt(sndData, sndIndex) & ENDC)
             Exit Sub
        End If
    Case SendTarget.ToDiosesYclan
        
        LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        While LoopC > 0
            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
            End If
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        Wend

        LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
        While LoopC > 0
            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
            End If
            LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
        Wend

        Exit Sub
        

End Select

End Sub

#End If

Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean


Dim x As Integer, Y As Integer
For Y = UserList(Index).pos.Y - MinYBorder + 1 To UserList(Index).pos.Y + MinYBorder - 1
        For x = UserList(Index).pos.x - MinXBorder + 1 To UserList(Index).pos.x + MinXBorder - 1

            If MapData(UserList(Index).pos.Map, x, Y).userindex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
        Next x
Next Y
EstaPCarea = False
End Function

Function HayPCarea(pos As WorldPos) As Boolean


Dim x As Integer, Y As Integer
For Y = pos.Y - MinYBorder + 1 To pos.Y + MinYBorder - 1
        For x = pos.x - MinXBorder + 1 To pos.x + MinXBorder - 1
            If x > 0 And Y > 0 And x < 101 And Y < 101 Then
                If MapData(pos.Map, x, Y).userindex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next x
Next Y
HayPCarea = False
End Function

Function HayOBJarea(pos As WorldPos, ObjIndex As Integer) As Boolean


Dim x As Integer, Y As Integer
For Y = pos.Y - MinYBorder + 1 To pos.Y + MinYBorder - 1
        For x = pos.x - MinXBorder + 1 To pos.x + MinXBorder - 1
            If MapData(pos.Map, x, Y).OBJInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function
            End If
        
        Next x
Next Y
HayOBJarea = False
End Function

Function ValidateChr(ByVal userindex As Integer) As Boolean

ValidateChr = UserList(userindex).char.Head <> 0 _
                And UserList(userindex).char.Body <> 0 _
                And ValidateSkills(userindex)

End Function

Sub ConnectUser(ByVal userindex As Integer, name As String, Password As String)
Dim n As Integer
Dim tStr As String





'Reseteamos los FLAGS
UserList(userindex).flags.Escondido = 0
UserList(userindex).flags.TargetNPC = 0
UserList(userindex).flags.TargetNpcTipo = eNPCType.Comun
UserList(userindex).flags.TargetObj = 0
UserList(userindex).flags.TargetUser = 0
UserList(userindex).char.FX = 0

Call SendData(toIndex, userindex, 0, "||RevivalAo> Presiona - AYUDA - para saber donde entrenar!." & FONTTYPE_WARNING)

If AlmacenaDominador = vbNullString Then
Call SendData(toIndex, userindex, 0, "||Castillo de Ullathorpe: Nadie." & FONTTYPE_CONSEJOCAOSVesA)
Else
Call SendData(toIndex, userindex, 0, "||Castillo de Ullathorpe: " & AlmacenaDominador & " " & HoraUlla & FONTTYPE_CONSEJOCAOSVesA)
End If
If AlmacenaDominadornix = vbNullString Then
Call SendData(toIndex, userindex, 0, "||Castillo de Nix: Nadie." & FONTTYPE_CONSEJOCAOSVesA)
Else
Call SendData(toIndex, userindex, 0, "||Castillo de Nix: " & AlmacenaDominadornix & " " & HoraNix & FONTTYPE_CONSEJOCAOSVesA)
End If
If Lemuria = vbNullString Then
Call SendData(toIndex, userindex, 0, "||Castillo de Asgard: Nadie." & FONTTYPE_CONSEJOCAOSVesA)
Else
Call SendData(toIndex, userindex, 0, "||Castillo de Asgard: " & Lemuria & " " & HoraLemuria & FONTTYPE_CONSEJOCAOSVesA)
End If
If Tale = vbNullString Then
Call SendData(toIndex, userindex, 0, "||Castillo de Tale: Nadie." & FONTTYPE_CONSEJOCAOSVesA)
Else
Call SendData(toIndex, userindex, 0, "||Castillo de Tale: " & Tale & " " & HoraTale & FONTTYPE_CONSEJOCAOSVesA)
End If
If Fortaleza = vbNullString Then
Call SendData(toIndex, userindex, 0, "||Fortaleza: Nadie." & FONTTYPE_CONSEJOCAOSVesA)
Else
Call SendData(toIndex, userindex, 0, "||Fortaleza: " & Fortaleza & " " & HoraForta & FONTTYPE_CONSEJOCAOSVesA)
End If
'Controlamos no pasar el maximo de usuarios
If NumUsers >= MaxUsers Then
    Call SendData(SendTarget.toIndex, userindex, 0, "ERREl servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
    Call CloseSocket(userindex)
    Exit Sub
End If

'¿Este IP ya esta conectado?
If AllowMultiLogins = 0 Then
    If CheckForSameIP(userindex, UserList(userindex).ip) = True Then
        Call SendData(SendTarget.toIndex, userindex, 0, "ERRNo es posible usar mas de un personaje al mismo tiempo.")
        Call CloseSocket(userindex)
        Exit Sub
    End If
End If

'¿Existe el personaje?
If Not FileExist(CharPath & UCase$(name) & ".chr", vbNormal) Then
    Call SendData(SendTarget.toIndex, userindex, 0, "ERREl personaje no existe.")
    Call CloseSocket(userindex)
    Exit Sub
End If

'¿Es el passwd valido?
If UCase$(Password) <> UCase$(GetVar(CharPath & UCase$(name) & ".chr", "INIT", "Password")) Then
    Call SendData(SendTarget.toIndex, userindex, 0, "ERRPassword incorrecto.")
    
    Call CloseSocket(userindex)
    Exit Sub
End If

'¿Ya esta conectado el personaje?
If CheckForSameName(userindex, name) Then
    If UserList(NameIndex(name)).Counters.Saliendo Then
        Call SendData(SendTarget.toIndex, userindex, 0, "ERREl usuario está saliendo.")
    Else
        Call SendData(SendTarget.toIndex, userindex, 0, "ERRPerdon, un usuario con el mismo nombre se há logoeado.")
    End If
    Call CloseSocket(userindex)
    Exit Sub
End If

'Cargamos el personaje
Dim Leer As New clsIniReader

Call Leer.Initialize(CharPath & UCase$(name) & ".chr")

'Cargamos los datos del personaje
Call LoadUserInit(userindex, Leer)

Call LoadUserStats(userindex, Leer)

If Not ValidateChr(userindex) Then
    Call SendData(SendTarget.toIndex, userindex, 0, "ERRError en el personaje.")
    Call CloseSocket(userindex)
    Exit Sub
End If

Call LoadUserReputacion(userindex, Leer)

Set Leer = Nothing

If UserList(userindex).Invent.EscudoEqpSlot = 0 Then UserList(userindex).char.ShieldAnim = NingunEscudo
If UserList(userindex).Invent.CascoEqpSlot = 0 Then UserList(userindex).char.CascoAnim = NingunCasco

If UserList(userindex).Invent.WeaponEqpSlot = 0 Then UserList(userindex).char.WeaponAnim = NingunArma


Call UpdateUserInv(True, userindex, 0)
Call UpdateUserHechizos(True, userindex, 0)

If UserList(userindex).flags.Navegando = 1 Then
     UserList(userindex).char.Body = ObjData(UserList(userindex).Invent.BarcoObjIndex).Ropaje
     UserList(userindex).char.Head = 0
     UserList(userindex).char.WeaponAnim = NingunArma
     UserList(userindex).char.ShieldAnim = NingunEscudo
     UserList(userindex).char.CascoAnim = NingunCasco
       '[MaTeO 9]
     UserList(userindex).char.Alas = NingunAlas
     '[/MaTeO 9]
End If

If UserList(userindex).flags.Paralizado Then
#If SeguridadAlkon Then
    If EncriptarProtocolosCriticos Then
        Call SendCryptedData(SendTarget.toIndex, userindex, 0, "PARADOW")
    Else
#End If
        Call SendData(SendTarget.toIndex, userindex, 0, "PARADOW")
#If SeguridadAlkon Then
    End If
#End If
End If

'Feo, esto tiene que ser parche cliente
If UserList(userindex).flags.Estupidez = 0 Then Call SendData(SendTarget.toIndex, userindex, 0, "NESTUP")
'

'Posicion de comienzo
If UserList(userindex).pos.Map = 0 Then
    If UCase$(UserList(userindex).Hogar) = "NIX" Then
             UserList(userindex).pos = Nix
    ElseIf UCase$(UserList(userindex).Hogar) = "ULLATHORPE" Then
             UserList(userindex).pos = Ullathorpe
    ElseIf UCase$(UserList(userindex).Hogar) = "BANDERBILL" Then
             UserList(userindex).pos = Banderbill
    ElseIf UCase$(UserList(userindex).Hogar) = "LINDOS" Then
             UserList(userindex).pos = Lindos
    Else
        UserList(userindex).Hogar = "ULLATHORPE"
        UserList(userindex).pos = Ullathorpe
    End If
Else

   ''TELEFRAG
    If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex <> 0 Then
        ''si estaba en comercio seguro...
        If UserList(MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex).ComUsu.DestUsu > 0 Then
            If UserList(UserList(MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex).ComUsu.DestUsu).flags.UserLogged Then
                Call FinComerciarUsu(UserList(MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex).ComUsu.DestUsu)
                Call SendData(SendTarget.toIndex, UserList(MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex).ComUsu.DestUsu, 0, "||Comercio cancelado. El otro usuario se ha desconectado." & FONTTYPE_TALK)
            End If
            End If
            If UserList(MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex).flags.UserLogged Then
                Call FinComerciarUsu(MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex)
               
            End If
        Call CloseSocket(MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex)
    End If
   
   
    If UserList(userindex).flags.Muerto = 1 Then
        Call Empollando(userindex)
    End If
End If

If Not MapaValido(UserList(userindex).pos.Map) Then
    Call SendData(SendTarget.toIndex, userindex, 0, "ERREL PJ se encuenta en un mapa invalido.")
    Call CloseSocket(userindex)
    Exit Sub
End If

'Nombre de sistema
UserList(userindex).name = name

UserList(userindex).Password = Password

UserList(userindex).showName = True 'Por default los nombres son visibles

'Info
Call SendData(SendTarget.toIndex, userindex, 0, "IU" & userindex) 'Enviamos el User index
Call SendData(SendTarget.toIndex, userindex, 0, "CM" & UserList(userindex).pos.Map & "," & MapInfo(UserList(userindex).pos.Map).MapVersion) 'Carga el mapa
Call SendData(SendTarget.toIndex, userindex, 0, "TM" & MapInfo(UserList(userindex).pos.Map).Music)
Call SendData(SendTarget.toIndex, userindex, 0, "N~" & MapInfo(UserList(userindex).pos.Map).name)

'Vemos que clase de user es (se lo usa para setear los privilegios alcrear el PJ)
UserList(userindex).flags.EsRolesMaster = EsRolesMaster(name)
If EsAdmin(name) Then
    UserList(userindex).flags.Privilegios = PlayerType.Admin
    Call LogGM(UserList(userindex).name, "Se conecto con ip:" & UserList(userindex).ip, False)
ElseIf EsDios(name) Then
    UserList(userindex).flags.Privilegios = PlayerType.Dios
    Call LogGM(UserList(userindex).name, "Se conecto con ip:" & UserList(userindex).ip, False)
ElseIf EsSemiDios(name) Then
    UserList(userindex).flags.Privilegios = PlayerType.SemiDios
    Call LogGM(UserList(userindex).name, "Se conecto con ip:" & UserList(userindex).ip, False)
ElseIf EsConsejero(name) Then
    UserList(userindex).flags.Privilegios = PlayerType.Consejero
    Call LogGM(UserList(userindex).name, "Se conecto con ip:" & UserList(userindex).ip, True)
Else
    UserList(userindex).flags.Privilegios = PlayerType.User
End If

''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
UserList(userindex).Counters.IdleCount = 0
'Crea  el personaje del usuario
Call MakeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y)

Call SendData(SendTarget.toIndex, userindex, 0, "IP" & UserList(userindex).char.CharIndex)
''[/el oso]

Call SendUserStatsBox(userindex)
Call SendUserHitBox(userindex)
Call EnviarHambreYsed(userindex)
Call EnviarDopa(userindex)

If haciendoBK Then
    Call SendData(SendTarget.toIndex, userindex, 0, "BKW")
    Call SendData(SendTarget.toIndex, userindex, 0, "||RevivalAo> Por favor espera algunos segundos, WorldSave esta ejecutandose." & FONTTYPE_SERVER)
End If

If EnPausa Then
    Call SendData(SendTarget.toIndex, userindex, 0, "BKW")
    Call SendData(SendTarget.toIndex, userindex, 0, "||RevivalAo> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde." & FONTTYPE_SERVER)
End If

If EnTesting And UserList(userindex).Stats.ELV >= 18 Then
    Call SendData(SendTarget.toIndex, userindex, 0, "ERRServidor en Testing por unos minutos, conectese con PJs de nivel menor a 18. No se conecte con Pjs que puedan resultar importantes por ahora pues pueden arruinarse.")
    Call CloseSocket(userindex)
    Exit Sub
End If

'Actualiza el Num de usuarios
'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!
NumUsers = NumUsers + 1
UserList(userindex).flags.UserLogged = True
Call SendData(toAll, 0, 0, "³" & NumUsers)
'usado para borrar Pjs
Call WriteVar(CharPath & UserList(userindex).name & ".chr", "INIT", "Logged", "1")

Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
Call SendData(toAll, 0, 0, "³" & NumUsers)
MapInfo(UserList(userindex).pos.Map).NumUsers = MapInfo(UserList(userindex).pos.Map).NumUsers + 1

If UserList(userindex).Stats.SkillPts > 0 Then
    Call EnviarSkills(userindex)
    Call EnviarSubirNivel(userindex, UserList(userindex).Stats.SkillPts)
End If

If NumUsers > DayStats.MaxUsuarios Then DayStats.MaxUsuarios = NumUsers

If NumUsers > recordusuarios Then
    Call SendData(SendTarget.toAll, 0, 0, "||Record de usuarios conectados simultaneamente." & "Hay " & NumUsers & " usuarios." & FONTTYPE_TURQ)
    recordusuarios = NumUsers
    Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(recordusuarios))
    
    Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)
End If

If UserList(userindex).NroMacotas > 0 Then
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        If UserList(userindex).MascotasType(i) > 0 Then
            UserList(userindex).MascotasIndex(i) = SpawnNpc(UserList(userindex).MascotasType(i), UserList(userindex).pos, True, True)
            
            If UserList(userindex).MascotasIndex(i) > 0 Then
                Npclist(UserList(userindex).MascotasIndex(i)).MaestroUser = userindex
                Call FollowAmo(UserList(userindex).MascotasIndex(i))
            Else
                UserList(userindex).MascotasIndex(i) = 0
            End If
        End If
    Next i
End If

If UserList(userindex).flags.Navegando = 1 Then Call SendData(SendTarget.toIndex, userindex, 0, "NAVEG")

If Criminal(userindex) Then
    'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Miembro de las fuerzas del caos > Seguro desactivado <" & FONTTYPE_FIGHT)
    Call SendData(SendTarget.toIndex, userindex, 0, "OFFOFS")
    UserList(userindex).flags.Seguro = False
Else
    UserList(userindex).flags.Seguro = True
    Call SendData(SendTarget.toIndex, userindex, 0, "ONONS")
End If

Call SendData(SendTarget.toIndex, userindex, 0, "SEGCO99")
UserList(userindex).flags.SeguroClan = False

If ServerSoloGMs > 0 Then
    If UserList(userindex).flags.Privilegios < ServerSoloGMs Then
        Call SendData(SendTarget.toIndex, userindex, 0, "ERRServidor restringido a administradores de jerarquia mayor o igual a: " & ServerSoloGMs & ". Por favor intente en unos momentos.")
        Call CloseSocket(userindex)
        Exit Sub
    End If
End If

If UserList(userindex).GuildIndex > 0 Then
Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||" & UserList(userindex).name & " Conectó" & FONTTYPE_GUILD)
    If Not modGuilds.m_ConectarMiembroAClan(userindex, UserList(userindex).GuildIndex) Then
        Call SendData(SendTarget.toIndex, userindex, 0, "||Tu estado no te permite entrar al clan." & FONTTYPE_GUILD)
    End If
End If

Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & FXIDs.FXWARP & "," & 0)

Call SendData(SendTarget.toIndex, userindex, 0, "LODXXD")

Call modGuilds.SendGuildNews(userindex)

If UserList(userindex).flags.NoActualizado Then
    Call SendData(SendTarget.toIndex, userindex, 0, "REAU")
End If

If Lloviendo Then Call SendData(SendTarget.toIndex, userindex, 0, "LLU")

tStr = modGuilds.a_ObtenerRechazoDeChar(UserList(userindex).name)

If tStr <> vbNullString Then
    Call SendData(SendTarget.toIndex, userindex, 0, "!!Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr & ENDC)
End If

Call MostrarNumUsers

n = FreeFile
Open App.Path & "\logs\numusers.log" For Output As n
Print #n, NumUsers
Close #n

n = FreeFile
'Log
Open App.Path & "\logs\Connect.log" For Append Shared As #n
Print #n, UserList(userindex).name & " ha entrado al juego. UserIndex:" & userindex & " " & Time & " " & Date
Close #n

End Sub

Sub SendMOTD(ByVal userindex As Integer)
    Dim j As Long
    
    Call SendData(SendTarget.toIndex, userindex, 0, "||Mensajes de entrada:" & FONTTYPE_INFO)
    
    For j = 1 To MaxLines
        Call SendData(SendTarget.toIndex, userindex, 0, "||" & Chr$(3) & MOTD(j).texto)
    Next j
End Sub

Sub ResetFacciones(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(userindex).Faccion
        .ArmadaReal = 0
        .FuerzasCaos = 0
        .CiudadanosMatados = 0
        .CriminalesMatados = 0
        .RecibioArmaduraCaos = 0
        .RecibioArmaduraReal = 0
        .RecibioExpInicialCaos = 0
        .RecibioExpInicialReal = 0
        .RecompensasCaos = 0
        .RecompensasReal = 0
        .Reenlistadas = 0
    End With
End Sub

Sub ResetContadores(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(userindex).Counters
        .AGUACounter = 0
        .AttackCounter = 0
        .Ceguera = 0
        .COMCounter = 0
        .Estupidez = 0
        .Frio = 0
        .HPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .Paralisis = 0
        .Pasos = 0
        .Pena = 0
        .PiqueteC = 0
        .STACounter = 0
        .Veneno = 0
        .Ocultando = 0
        .TimerLanzarSpell = 0
        .TimerPuedeAtacar = 0
        .TimerPuedeTrabajar = 0
        .TimerUsar = 0
    End With
End Sub

Sub ResetCharInfo(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(userindex).char
        .Body = 0
        .CascoAnim = 0
    
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .loops = 0
        .Heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
    End With
End Sub

Sub ResetBasicUserInfo(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(userindex)
        .name = ""
        .modName = ""
        .Password = ""
        .Desc = ""
        .DescRM = ""
        .pos.Map = 0
        .pos.x = 0
        .pos.Y = 0
        .ip = ""
        .RDBuffer = ""
        .Clase = ""
        .email = ""
        .Genero = ""
        .Hogar = ""
        .Raza = ""

        .RandomCode = 0
        .PrevCheckSum = 0
        .PacketNumber = 0

        .EmpoCont = 0
        .PartyIndex = 0
        .PartySolicitud = 0
        
        With .Stats
            .Banco = 0
            .ELV = 0
            .ELU = 0
            .Exp = 0
            .def = 0
            .CriminalesMatados = 0
            .NPCsMuertos = 0
            .UsuariosMatados = 0
            .SkillPts = 0
        End With
    End With
End Sub

Sub ResetReputacion(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(userindex).Reputacion
        .AsesinoRep = 0
        .BandidoRep = 0
        .BurguesRep = 0
        .LadronesRep = 0
        .NobleRep = 0
        .PlebeRep = 0
        .NobleRep = 0
        .Promedio = 0
    End With
End Sub

Sub ResetGuildInfo(ByVal userindex As Integer)
    If UserList(userindex).EscucheClan > 0 Then
        Call modGuilds.GMDejaDeEscucharClan(userindex, UserList(userindex).EscucheClan)
        UserList(userindex).EscucheClan = 0
    End If
    If UserList(userindex).GuildIndex > 0 Then
        Call modGuilds.m_DesconectarMiembroDelClan(userindex, UserList(userindex).GuildIndex)
    End If
    UserList(userindex).GuildIndex = 0
End Sub

Sub ResetUserFlags(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/29/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'03/29/2006 Maraxus - Reseteo el CentinelaOK también.
'*************************************************
    With UserList(userindex).flags
        '[MaTeO 13]
        .TiempoMapa = 0
        '[/MaTeO 13]
        .Comerciando = False
        .Ban = 0
        .Escondido = 0
        .DuracionEfecto = 0
        .NpcInv = 0
        .StatsChanged = 0
        .TargetNPC = 0
        .TargetNpcTipo = eNPCType.Comun
        .TargetObj = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        .TargetUser = 0
        .TipoPocion = 0
        .TomoPocion = False
        .Descuento = ""
        .Hambre = 0
        .Sed = 0
        .Descansar = False
        .ClienteOK = False
        .Navegando = 0
        .Oculto = 0
        .Envenenado = 0
        .Invisible = 0
        .Paralizado = 0
        .Maldicion = 0
        .Bendicion = 0
        .Meditando = 0
        .YaDenuncio = 0
        .Privilegios = PlayerType.User
        .PuedeMoverse = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .ValCoDe = 0
        .Hechizo = 0
        .TimesWalk = 0
        .StartWalk = 0
        .CountSH = 0
        .EstaEmpo = 0
        .PertAlCons = 0
        .PertAlConsCaos = 0
        .EnDosVDos = False
        .ParejaMuerta = False
        .envioSol = False
        .RecibioSol = False
        .Soporteo = False
        .EstaDueleando1 = False
        .Oponente1 = 0
        .EsperandoDuelo1 = False
        .EstaDueleando = False
        .Oponente = 0
        .EsperandoDuelo = False
    End With
    UserList(userindex).Counters.AntiSH = 0
End Sub

Sub ResetUserSpells(ByVal userindex As Integer)
    Dim LoopC As Long
    For LoopC = 1 To MAXUSERHECHIZOS
        UserList(userindex).Stats.UserHechizos(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserPets(ByVal userindex As Integer)
    Dim LoopC As Long
    
    UserList(userindex).NroMacotas = 0
        
    For LoopC = 1 To MAXMASCOTAS
        UserList(userindex).MascotasIndex(LoopC) = 0
        UserList(userindex).MascotasType(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserBanco(ByVal userindex As Integer)
    Dim LoopC As Long
    
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
          UserList(userindex).BancoInvent.Object(LoopC).Amount = 0
          UserList(userindex).BancoInvent.Object(LoopC).Equipped = 0
          UserList(userindex).BancoInvent.Object(LoopC).ObjIndex = 0
    Next LoopC
    
    UserList(userindex).BancoInvent.NroItems = 0
End Sub

Public Sub LimpiarComercioSeguro(ByVal userindex As Integer)
    With UserList(userindex).ComUsu
        If .DestUsu > 0 Then
            Call FinComerciarUsu(.DestUsu)
            Call FinComerciarUsu(userindex)
        End If
    End With
End Sub

Sub ResetUserSlot(ByVal userindex As Integer)

    UserList(userindex).keyDH = 0
    UserList(userindex).MykeySecretDH = 0
Dim UsrTMP As User

Set UserList(userindex).CommandsBuffer = Nothing


Set UserList(userindex).ColaSalida = Nothing
UserList(userindex).SockPuedoEnviar = False
UserList(userindex).ConnIDValida = False
UserList(userindex).ConnID = -1

Call LimpiarComercioSeguro(userindex)
Call ResetFacciones(userindex)
Call ResetContadores(userindex)
Call ResetCharInfo(userindex)
Call ResetBasicUserInfo(userindex)
Call ResetReputacion(userindex)
Call ResetGuildInfo(userindex)

Call ResetUserFlags(userindex)
Call LimpiarInventario(userindex)
Call ResetUserSpells(userindex)
Call ResetUserPets(userindex)
Call ResetUserBanco(userindex)

With UserList(userindex).ComUsu
    .Acepto = False
    .Cant = 0
    .DestNick = ""
    .DestUsu = 0
    .Objeto = 0
End With

UserList(userindex) = UsrTMP
UserList(userindex).autoaim = False
End Sub


Sub CloseUser(ByVal userindex As Integer)
'Call LogTarea("CloseUser " & UserIndex)

On Error GoTo errhandler

Dim n As Integer
Dim x As Integer
Dim Y As Integer
Dim LoopC As Integer
Dim Map As Integer
Dim name As String
Dim Raza As String
Dim Clase As String
Dim i As Integer

Dim aN As Integer

aN = UserList(userindex).flags.AtacadoPorNpc
If aN > 0 Then
      Npclist(aN).Movement = Npclist(aN).flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
      Npclist(aN).flags.AttackedBy = ""
End If
UserList(userindex).flags.AtacadoPorNpc = 0

Map = UserList(userindex).pos.Map
x = UserList(userindex).pos.x
Y = UserList(userindex).pos.Y
name = UCase$(UserList(userindex).name)
Raza = UserList(userindex).Raza
Clase = UserList(userindex).Clase

UserList(userindex).char.FX = 0
UserList(userindex).char.loops = 0
Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).char.CharIndex & "," & 0 & "," & 0)

UserList(userindex).flags.UserLogged = False
UserList(userindex).Counters.Saliendo = False

'Le devolvemos el body y head originales
If UserList(userindex).flags.AdminInvisible = 1 Then Call DoAdminInvisible(userindex)

'si esta en party le devolvemos la experiencia
If UserList(userindex).PartyIndex > 0 Then Call mdParty.SalirDeParty(userindex)

'[MaTeO ASEDIO]
Call ResetFlagsAsedio(userindex)
'[/MaTeO ASEDIO]

' Grabamos el personaje del usuario
Call SaveUser(userindex, CharPath & name & ".chr")

'usado para borrar Pjs
Call WriteVar(CharPath & UserList(userindex).name & ".chr", "INIT", "Logged", "0")


'Quitar el dialogo
'If MapInfo(Map).NumUsers > 0 Then
'    Call SendToUserArea(UserIndex, "QDL" & UserList(UserIndex).Char.charindex)
'End If

If MapInfo(Map).NumUsers > 0 Then
    Call SendData(SendTarget.ToMapButIndex, userindex, Map, "QDL" & UserList(userindex).char.CharIndex)
End If



'Borrar el personaje
If UserList(userindex).char.CharIndex > 0 Then
    Call EraseUserChar(SendTarget.ToMap, userindex, Map, userindex)
End If

'Borrar mascotas
For i = 1 To MAXMASCOTAS
    If UserList(userindex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(userindex).MascotasIndex(i)).flags.NPCActive Then _
            Call QuitarNPC(UserList(userindex).MascotasIndex(i))
    End If
Next i

'Update Map Users
MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1

If MapInfo(Map).NumUsers < 0 Then
    MapInfo(Map).NumUsers = 0
End If

' Si el usuario habia dejado un msg en la gm's queue lo borramos
If Ayuda.Existe(UserList(userindex).name) Then Call Ayuda.Quitar(UserList(userindex).name)
If Torneo.Existe(UserList(userindex).name) Then Call Torneo.Quitar(UserList(userindex).name)

Call ResetUserSlot(userindex)

Call MostrarNumUsers

n = FreeFile(1)
Open App.Path & "\logs\Connect.log" For Append Shared As #n
Print #n, name & " há dejado el juego. " & "User Index:" & userindex & " " & Time & " " & Date
Close #n

Exit Sub

errhandler:
Call LogError("Error en CloseUser. Número " & Err.Number & " Descripción: " & Err.Description)


End Sub

Sub HandleData(ByVal userindex As Integer, ByVal rData As String)

      '
      ' ATENCION: Cambios importantes en HandleData.
      ' =========
      '
      '           La funcion se encuentra dividida en 2,
      '           una parte controla los comandos que
      '           empiezan con "/" y la otra los comanos
      '           que no. (Basado en la idea de Barrin)
      '


10    Call LogTarea("Sub HandleData :" & rData & " " & UserList(userindex).name)

      'Nunca jamas remover o comentar esta linea !!!
      'Nunca jamas remover o comentar esta linea !!!
      'Nunca jamas remover o comentar esta linea !!!
On Error GoTo ErrorHandler:
      'Nunca jamas remover o comentar esta linea !!!
      'Nunca jamas remover o comentar esta linea !!!
      'Nunca jamas remover o comentar esta linea !!!
      '
      'Ah, no me queres hacer caso ? Entonces
      'atenete a las consecuencias!!
      '

          Dim CadenaOriginal As String
          
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
          
          Dim sndData As String
          Dim cliMD5 As String
          Dim ClientChecksum As String
          Dim ServerSideChecksum As Long
          Dim IdleCountBackup As Long
           
20             UserList(userindex).clave2 = UserList(userindex).clave2 + 1
30         With AodefConv
40    SuperClave = .Numero2Letra(UserList(userindex).clave2, , 2, "ZiPPy", "NoPPy", 1, 0)
50    End With
60    Do While InStr(1, SuperClave, " ")
70    SuperClave = mid$(SuperClave, 1, InStr(1, SuperClave, " ") - 1) & mid$(SuperClave, InStr(1, SuperClave, " ") + 1)
80    Loop
90    SuperClave = Semilla(SuperClave)
100       UserList(userindex).clave = SuperClave
          
110       If UserList(userindex).clave2 = 999998 Then
120       UserList(userindex).clave2 = 0
130       End If
140       rData = DeCodificar(AoDefDecode(rData), UserList(userindex).clave)
150       CadenaOriginal = rData
          '¿Tiene un indece valido?
160       If userindex <= 0 Then
170           Call CloseSocket(userindex)
180           Exit Sub
190       End If
200           If Left$(rData, 5) = "CLAVE" Then
            
210           UserList(userindex).clave = Right$(rData, Len(rData) - 5)
              
220           Exit Sub
230       End If
240   If Left$(rData, 13) = "gIvEmEvAlcOde" Then
#If SeguridadAlkon Then
              '<<<<<<<<<<< MODULO PRIVADO DE CADA IMPLEMENTACION >>>>>>
250           UserList(userindex).flags.ValCoDe = Chr(RandomNumber(65, 90)) & Chr(RandomNumber(65, 90)) & Chr(RandomNumber(65, 90)) & Chr(RandomNumber(65, 90)) & Chr(RandomNumber(65, 90)) & Chr(RandomNumber(65, 90)) & Chr(RandomNumber(65, 90)) & Chr(RandomNumber(65, 90)) 'RandomNumber(20000, 32000))
260           UserList(userindex).RandomCode = RandomNumber(1, 32000)
270           UserList(userindex).Antichit = RandomNumber(1, 32000)
280           UserList(userindex).PrevCheckSum = UserList(userindex).RandomCode
290           UserList(userindex).PacketNumber = 100
300           UserList(userindex).KeyCrypt = (UserList(userindex).RandomCode And 17320) Xor (UserList(userindex).flags.ValCoDe Xor 4232)
              '<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>
310           Call SendData(SendTarget.toIndex, userindex, 0, "VAL" & UserList(userindex).RandomCode & "," & UserList(userindex).Antichit)
320           Exit Sub
330       Else
              '<<<<<<<<<<< MODULO PRIVADO DE CADA IMPLEMENTACION >>>>>>
340           ClientChecksum = Right$(rData, Len(rData) - InStrRev(rData, Chr$(126)))
350           tStr = Left$(rData, Len(rData) - Len(ClientChecksum) - 1)
360           ServerSideChecksum = CheckSum(UserList(userindex).PrevCheckSum, tStr)
370           UserList(userindex).PrevCheckSum = ServerSideChecksum
              
380           If CLng(ClientChecksum) <> ServerSideChecksum Then
390               Call LogError("Checksum error userindex: " & userindex & " rdata: " & rData)
400               Call CloseSocket(userindex, True)
410           End If
              
              'Remove checksum from data
420           rData = tStr
430           tStr = ""
#Else
440           UserList(userindex).flags.ValCoDe = Chr(RandomNumber(65, 90)) & Chr(RandomNumber(65, 90)) & Chr(RandomNumber(65, 90)) & Chr(RandomNumber(65, 90)) & Chr(RandomNumber(65, 90)) & Chr(RandomNumber(65, 90)) & Chr(RandomNumber(65, 90)) & Chr(RandomNumber(65, 90)) 'RandomNumber(20000, 32000)
450           UserList(userindex).RandomCode = RandomNumber(1, 32000)
460           UserList(userindex).Antichit = RandomNumber(1, 32000)
470           Call SendData(SendTarget.toIndex, userindex, 0, "VAL" & UserList(userindex).RandomCode & "," & UserList(userindex).Antichit)
480           Exit Sub
#End If
490       End If
          '<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>

500       IdleCountBackup = UserList(userindex).Counters.IdleCount
510       UserList(userindex).Counters.IdleCount = 0
         
520       If Not UserList(userindex).flags.UserLogged Then
      ' SATUROS HIZO ESTO DEJO EL CREDITO Q DSPS SE DEJAN PA USTEDES TO SAKJJSHASKASJKAS (ENCRIPTACION)
        Dim rDataDecripted As String
530           rDataDecripted = DecryptStr(Decode64(rData), "xaopepe")
540           Select Case Left$(rDataDecripted, 6)
                      
      Case "MARAKA"

550                   rData = Right$(rDataDecripted, Len(rDataDecripted) - 6)
                      



         '             End If
560                   Ver = ReadField(3, rData, 44)


570   If VersionOK(Ver) Then

                          
580                       tName = ReadField(1, rData, 44)
                          
590                         If Not EsDios(tName) And Not EsSemiDios(tName) And Not EsConsejero(tName) Then
600                       If UserList(userindex).esgm = True Then
610                       Call SendData(SendTarget.toIndex, userindex, 0, "XAI")
620                       Exit Sub
630                       End If
640                       End If
                          
650                       If EsDios(tName) Then
660                       If UserList(userindex).esgm = False Then
670                       Call SendData(SendTarget.toIndex, userindex, 0, "XAO")
680                       Exit Sub
690                       End If
700                       End If
                          
710                           If EsSemiDios(tName) Then
720                       If UserList(userindex).esgm = False Then
730                       Call SendData(SendTarget.toIndex, userindex, 0, "XAO")
740                       Exit Sub
750                       End If
760                       End If
                          
770                           If EsConsejero(tName) Then
780                       If UserList(userindex).esgm = False Then
790                       Call SendData(SendTarget.toIndex, userindex, 0, "XAO")
800                       Exit Sub
810                       End If
820                       End If
                          
830                       If ReadField(11, rData, 44) <> UserList(userindex).RandomCode Then
840                           Call SendData(SendTarget.toIndex, userindex, 0, "ERRCliente invalido.")
850                           Exit Sub
860                       End If
                          
870                       If Not AsciiValidos(tName) Then
880                           Call SendData(SendTarget.toIndex, userindex, 0, "ERRNombre invalido.")
890                           Call CloseSocket(userindex, True)
900                           Exit Sub
910                       End If
                          
920                       If Not PersonajeExiste(tName) Then
930                           Call SendData(SendTarget.toIndex, userindex, 0, "ERREl personaje no existe.")
940                           Call CloseSocket(userindex, True)
950                           Exit Sub
960                       End If
                          
970                       If Not BANCheck(tName) Then
                              'If ValidarLoginMSG(UserList(UserIndex).flags.ValCoDe) <> CInt(ReadField(11, Left$(rData, Len(rData) - 16), 44)) Then
                              '    Call LogHackAttemp("IP:" & UserList(UserIndex).ip & " intento crear un bot.")
                              '    Call CloseSocket(UserIndex)
                              '    Exit Sub
                              'End If
                              
980                           UserList(userindex).flags.NoActualizado = Not VersionesActuales(val(ReadField(4, rData, 44)), val(ReadField(5, rData, 44)), val(ReadField(6, rData, 44)), val(ReadField(7, rData, 44)), val(ReadField(8, rData, 44)), val(ReadField(9, rData, 44)), val(ReadField(10, rData, 44)))
                              
                              Dim Pass11 As String
990                           Pass11 = ReadField(2, rData, 44)
1000                          Call ConnectUser(userindex, tName, Pass11)
1010                      Else
1020                          Call SendData(SendTarget.toIndex, userindex, 0, "ERRSe te ha prohibido la entrada a RevivalAo")
1030                      End If
1040                  Else
1050                       Call SendData(SendTarget.toIndex, userindex, 0, "ERRVersion Obsoleta")
                           'Call CloseSocket(UserIndex)
1060                       Exit Sub
1070                  End If
1080                  Exit Sub
                  

1090              Case "ZORRON"
1100                  If PuedeCrearPersonajes = 0 Then
1110                      Call SendData(SendTarget.toIndex, userindex, 0, "ERRLa creacion de personajes en este servidor se ha deshabilitado.")
1120                      Call CloseSocket(userindex)
1130                      Exit Sub
1140                  End If
                      
1150                  If ServerSoloGMs <> 0 Then
1160                      Call SendData(SendTarget.toIndex, userindex, 0, "ERRServidor restringido a administradores. Consulte la página oficial o el foro oficial para mas información.")
1170                      Call CloseSocket(userindex)
1180                      Exit Sub
1190                  End If

1200                  If aClon.MaxPersonajes(UserList(userindex).ip) Then
1210                      Call SendData(SendTarget.toIndex, userindex, 0, "ERRHas creado demasiados personajes.")
1220                      Call CloseSocket(userindex)
1230                      Exit Sub
1240                  End If
                      
1250                  rData = Right$(rDataDecripted, Len(rDataDecripted) - 6)
1260                 cliMD5 = Right$(rData, 16)
                      'rData = Right$(rData, Len(rData) - 6)
                      'cliMD5 = Right$(rData, 16)
                      'rData = Left$(rData, Len(rData) - 16)
1270                  If Not MD5ok(cliMD5) Then
1280                      Call SendData(SendTarget.toIndex, userindex, 0, "ERRCliente dañado Fijate si hay actualizaciones.")
1290                      Exit Sub
1300                  End If
                      
1310                  Ver = ReadField(3, rData, 44)
                     
                          'Dim miinteger As Integer
                          'miinteger = CInt(ReadField(37, rData, 44))
                          
                          'If ValidarLoginMSG(UserList(UserIndex).flags.ValCoDe) <> miinteger Then
                          '    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRPara poder continuar con la creación del personaje, debe utilizar el cliente proporcionado en ao.alkon.com.ar")
                              'Call LogHackAttemp("IP:" & UserList(UserIndex).ip & " intento crear un bot.")
                         '     Call CloseSocket(UserIndex)
                         '     Exit Sub
                         ' End If
                           
1320                      UserList(userindex).flags.NoActualizado = Not VersionesActuales(val(ReadField(37, rData, 44)), val(ReadField(38, rData, 44)), val(ReadField(39, rData, 44)), val(ReadField(40, rData, 44)), val(ReadField(41, rData, 44)), val(ReadField(42, rData, 44)), val(ReadField(43, rData, 44)))
                          
1330                      Call ConnectNewUser(userindex, ReadField(1, rData, 44), ReadField(2, rData, 44), ReadField(4, rData, 44), ReadField(5, rData, 44), ReadField(6, rData, 44), ReadField(7, rData, 44), _
                                              ReadField(8, rData, 44), ReadField(9, rData, 44), ReadField(10, rData, 44), ReadField(11, rData, 44), ReadField(12, rData, 44), ReadField(13, rData, 44), _
                                              ReadField(14, rData, 44), ReadField(15, rData, 44), ReadField(16, rData, 44), ReadField(17, rData, 44), ReadField(18, rData, 44), ReadField(19, rData, 44), _
                                              ReadField(20, rData, 44), ReadField(21, rData, 44), ReadField(22, rData, 44), ReadField(23, rData, 44), ReadField(24, rData, 44), ReadField(25, rData, 44), _
                                              ReadField(26, rData, 44), ReadField(27, rData, 44), ReadField(28, rData, 44), ReadField(29, rData, 44))
                      
                   
                      
1340                  Exit Sub
                      
1350                  Case "TIRDAD"
                  
1360                  UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = RandomNumber(17, 18)
1370                  UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = RandomNumber(17, 18)
1380                  UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = RandomNumber(16, 18)
1390                  UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = RandomNumber(17, 18)
1400                  UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) = RandomNumber(16, 18)
                      
1410                  Call SendData(SendTarget.toIndex, userindex, 0, "DODAS" & UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) & "," & UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) & "," & UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) & "," & UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) & "," & UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion))
                      
1420                  Exit Sub
                      
1430                   Case "ESGMQL"
1440              UserList(userindex).esgm = True
1450                  Exit Sub
                      
1460          End Select
      'If UCase$(rData) = "TIRDAD" Then
                  
                      'UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = RandomNumber(17, 18)
                      'UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = RandomNumber(17, 18)
                      'UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = RandomNumber(16, 18)
                      'UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = RandomNumber(17, 18)
                      'UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) = RandomNumber(16, 18)
                      
                      'Call SendData(SendTarget.toindex, userindex, 0, "DODAS" & UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) & "," & UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) & "," & UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) & "," & UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) & "," & UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion))
                      'End If
                      'Exit Sub
                      
1470      Select Case Left$(rData, 4)
              Case "BORR" ' <<< borra personajes
1480             On Error GoTo ExitErr1
1490              rData = Right$(rData, Len(rData) - 4)
1500              If (UserList(userindex).flags.ValCoDe = 0) Or (ValidarLoginMSG(UserList(userindex).flags.ValCoDe) <> CInt(val(ReadField(3, rData, 44)))) Then
1510                            Call LogHackAttemp("IP:" & UserList(userindex).ip & " intento borrar un personaje.")
1520                            Call CloseSocket(userindex)
1530                            Exit Sub
1540              End If
1550              Arg1 = ReadField(1, rData, 44)
                  
1560              If Not AsciiValidos(Arg1) Then Exit Sub
                  
                  '¿Existe el personaje?
1570              If Not FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) Then
1580                  Call CloseSocket(userindex)
1590                  Exit Sub
1600              End If
          
                  '¿Es el passwd valido?
1610              If UCase$(ReadField(2, rData, 44)) <> UCase$(GetVar(CharPath & UCase$(Arg1) & ".chr", "INIT", "Password")) Then
1620                  Call CloseSocket(userindex)
1630                  Exit Sub
1640              End If
          
                  'If FileExist(CharPath & ucase$(Arg1) & ".chr", vbNormal) Then
                      Dim rt As String
1650                  rt = App.Path & "\ChrBackUp\" & UCase$(Arg1) & ".bak"
1660                  If FileExist(rt, vbNormal) Then Kill rt
1670                  Name CharPath & UCase$(Arg1) & ".chr" As rt
1680                  Call SendData(SendTarget.toIndex, userindex, 0, "BORROK")
1690                  Exit Sub
ExitErr1:
1700          Call LogError(Err.Description & " " & rData)
1710          Exit Sub
                  'End If
1720      End Select

          '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
          'Si no esta logeado y envia un comando diferente a los
          'de arriba cerramos la conexion.
          '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
1730      Call LogHackAttemp("Mesaje enviado sin logearse:" & rData)
1740      Call CloseSocket(userindex)
1750      Exit Sub
            
1760  End If ' if not user logged


      Dim Procesado As Boolean

      ' bien ahora solo procesamos los comandos que NO empiezan
      ' con "/".
1770  If Left$(rData, 1) <> "/" Then
          
1780      Call HandleData_1(userindex, rData, Procesado)
1790      If Procesado Then Exit Sub
          
      ' bien hasta aca fueron los comandos que NO empezaban con
      ' "/". Ahora adiviná que sigue :)
1800  Else
          
1810      Call HandleData_2(userindex, rData, Procesado)
1820      If Procesado Then Exit Sub

1830  End If ' "/"

#If SeguridadAlkon Then
1840      If HandleDataEx(userindex, rData) Then Exit Sub
#End If


1850  If UserList(userindex).flags.Privilegios = PlayerType.User Then
1860      UserList(userindex).Counters.IdleCount = IdleCountBackup
1870  End If

      '>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<
1880   If UserList(userindex).flags.Privilegios = PlayerType.User Then Exit Sub
      '>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<

      '<<<<<<<<<<<<<<<<<<<< Consejeros <<<<<<<<<<<<<<<<<<<<
      '[rodra]
1890  If UCase$(Left$(rData, 6)) = "/SO33 " Then ' /ext <nick>
1900      rData = Right$(rData, Len(rData) - 6)
1910      TIndex = NameIndex(rData)
1920  If TIndex <> 0 Then ' si existe
1930          Call SendData(toIndex, TIndex, 0, "LEFT" & userindex)
1940      Else
1950          Call SendData(toIndex, userindex, 0, "||No se encuentra " & rData & FONTTYPE_INFO)
1960      End If
1970      Exit Sub
1980  End If
      '[rodra]


      'Mensaje del servidor
1990  If UCase$(Left$(rData, 3)) = "/R " Then
2000  rData = Right$(rData, Len(rData) - 3)
2010  Call LogGM(UserList(userindex).name, "Mensaje En GENERAL:" & rData, False)
2020  If rData <> "" Then
2030  Call SendData(toAll, 0, 0, "||<" & UserList(userindex).name & "> " & rData & FONTTYPE_TALK)

2040  End If
2050  Exit Sub
2060  End If

2070  If UCase$(rData) = "/SHOWNAME" Then
2080      If UserList(userindex).flags.EsRolesMaster Or UserList(userindex).flags.Privilegios >= PlayerType.Dios Then
2090          UserList(userindex).showName = Not UserList(userindex).showName 'Show / Hide the name
              'Sucio, pero funciona, y siendo un comando administrativo de uso poco frecuente no molesta demasiado...
2100          Call UsUaRiOs.EraseUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex)
2110          Call UsUaRiOs.MakeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y)
2120      End If
2130      Exit Sub
2140  End If

2150  If UCase$(rData) = "/ONLINEREAL" Then
2160      For tLong = 1 To LastUser
2170          If UserList(tLong).ConnID <> -1 Then
2180              If UserList(tLong).Faccion.ArmadaReal = 1 And (UserList(tLong).flags.Privilegios < PlayerType.Dios Or UserList(userindex).flags.Privilegios >= PlayerType.Dios) Then
2190                  tStr = tStr & UserList(tLong).name & ", "
2200              End If
2210          End If
2220      Next tLong
          
2230      If Len(tStr) > 0 Then
2240          Call SendData(SendTarget.toIndex, userindex, 0, "||Armadas conectados: " & Left$(tStr, Len(tStr) - 2) & FONTTYPE_INFO)
2250      Else
2260          Call SendData(SendTarget.toIndex, userindex, 0, "||No hay Armadas conectados" & FONTTYPE_INFO)
2270      End If
2280      Exit Sub
2290  End If

2300  If UCase$(rData) = "/ONLINECAOS" Then
2310      For tLong = 1 To LastUser
2320          If UserList(tLong).ConnID <> -1 Then
2330              If UserList(tLong).Faccion.FuerzasCaos = 1 And (UserList(tLong).flags.Privilegios < PlayerType.Dios Or UserList(userindex).flags.Privilegios >= PlayerType.Dios) Then
2340                  tStr = tStr & UserList(tLong).name & ", "
2350              End If
2360          End If
2370      Next tLong
          
2380      If Len(tStr) > 0 Then
2390          Call SendData(SendTarget.toIndex, userindex, 0, "||Caos conectados: " & Left$(tStr, Len(tStr) - 2) & FONTTYPE_INFO)
2400      Else
2410          Call SendData(SendTarget.toIndex, userindex, 0, "||No hay Caos conectados" & FONTTYPE_INFO)
2420      End If
2430      Exit Sub
2440  End If

      '/IRCERCA
      'este comando sirve para teletrasportarse cerca del usuario
2450  If UCase$(Left$(rData, 9)) = "/IRCERCA " Then
          Dim indiceUserDestino As Integer
2460      rData = Right$(rData, Len(rData) - 9) 'obtiene el nombre del usuario
2470      TIndex = NameIndex(rData)
          
          'Si es dios o Admins no podemos salvo que nosotros también lo seamos
2480      If (EsDios(rData) Or EsAdmin(rData)) And UserList(userindex).flags.Privilegios < PlayerType.Dios Then _
              Exit Sub
          
2490      If TIndex <= 0 Then 'existe el usuario destino?
2500          Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
2510          Exit Sub
2520      End If

2530      For tInt = 2 To 5 'esto for sirve ir cambiando la distancia destino
2540          For i = UserList(TIndex).pos.x - tInt To UserList(TIndex).pos.x + tInt
2550              For DummyInt = UserList(TIndex).pos.Y - tInt To UserList(TIndex).pos.Y + tInt
2560                  If (i >= UserList(TIndex).pos.x - tInt And i <= UserList(TIndex).pos.x + tInt) And (DummyInt = UserList(TIndex).pos.Y - tInt Or DummyInt = UserList(TIndex).pos.Y + tInt) Then
2570                      If MapData(UserList(TIndex).pos.Map, i, DummyInt).userindex = 0 And LegalPos(UserList(TIndex).pos.Map, i, DummyInt) Then
2580                          Call WarpUserChar(userindex, UserList(TIndex).pos.Map, i, DummyInt, True)
2590                          Exit Sub
2600                      End If
2610                  ElseIf (DummyInt >= UserList(TIndex).pos.Y - tInt And DummyInt <= UserList(TIndex).pos.Y + tInt) And (i = UserList(TIndex).pos.x - tInt Or i = UserList(TIndex).pos.x + tInt) Then
2620                      If MapData(UserList(TIndex).pos.Map, i, DummyInt).userindex = 0 And LegalPos(UserList(TIndex).pos.Map, i, DummyInt) Then
2630                          Call WarpUserChar(userindex, UserList(TIndex).pos.Map, i, DummyInt, True)
2640                          Exit Sub
2650                      End If
2660                  End If
2670              Next DummyInt
2680          Next i
2690      Next tInt
          
2700      Call SendData(SendTarget.toIndex, userindex, 0, "||Todos los lugares estan ocupados." & FONTTYPE_INFO)
2710      Exit Sub
2720  End If

      '/rem comentario
2730  If UCase$(Left$(rData, 4)) = "/REM" Then
2740      rData = Right$(rData, Len(rData) - 5)
2750      Call LogGM(UserList(userindex).name, "Comentario: " & rData, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
2760      Call SendData(SendTarget.toIndex, userindex, 0, "||Comentario salvado..." & FONTTYPE_INFO)
2770      Exit Sub
2780  End If

      'HORA
2790  If UCase$(Left$(rData, 5)) = "/HORA" Then
2800      Call LogGM(UserList(userindex).name, "Hora.", (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
2810      rData = Right$(rData, Len(rData) - 5)
2820      Call SendData(SendTarget.toAll, 0, 0, "||Hora: " & Time & " " & Date & FONTTYPE_INFO)
2830      Exit Sub
2840  End If


2850  If UCase$(Left$(rData, 8)) = "/MADERH " Then
2860      rData = Right$(rData, Len(rData) - 8)
2870      TIndex = NameIndex(rData)
      Dim trofeosmadera As Obj
2880  trofeosmadera.ObjIndex = 1007
2890  trofeosmadera.Amount = 1
2900   If Not TIndex > 0 Then Exit Sub
2910      If Not MeterItemEnInventario(TIndex, trofeosmadera) Then
2920      Call TirarItemAlPiso(UserList(TIndex).pos, trofeosmadera)
2930      End If
2940  Call SendData(SendTarget.toAll, userindex, 0, "||" & UserList(userindex).name & " le entrega 1 Amuleto de Madera a " & UserList(TIndex).name & " por haber salido cuarto en el torneo." & "~237~207~139~1~0")
2950  UserList(TIndex).Stats.TrofMadera = UserList(TIndex).Stats.TrofMadera + 1
2960  Call SendData(toAll, userindex, 0, "||" & UserList(TIndex).name & " Ya Lleva " & UserList(TIndex).Stats.TrofMadera & " Amuletos de Madera." & "~237~207~139~1~0")
2970  Exit Sub
2980  End If


2990  If UCase$(Left$(rData, 7)) = "/PLATH " Then
3000      rData = Right$(rData, Len(rData) - 7)
3010      TIndex = NameIndex(rData)
      Dim trofeosplata As Obj
3020  trofeosplata.ObjIndex = 991
3030  trofeosplata.Amount = 1
3040   If Not TIndex > 0 Then Exit Sub
3050      If Not MeterItemEnInventario(TIndex, trofeosplata) Then
3060      Call TirarItemAlPiso(UserList(TIndex).pos, trofeosplata)
3070      End If
3080  Call SendData(SendTarget.toAll, userindex, 0, "||" & UserList(userindex).name & " le entrega 1 trofeo de Plata a " & UserList(TIndex).name & " por haber salido segundo en el torneo." & "~196~198~196~1~0")
3090  UserList(TIndex).Stats.TrofPlata = UserList(TIndex).Stats.TrofPlata + 1
3100  Call SendData(toAll, userindex, 0, "||" & UserList(TIndex).name & " Ya Lleva " & UserList(TIndex).Stats.TrofPlata & " Trofeos de Plata." & "~196~198~196~1~0")
3110  Exit Sub
3120  End If

3130  If UCase$(Left$(rData, 8)) = "/BRONCH " Then
3140      rData = Right$(rData, Len(rData) - 8)
3150      TIndex = NameIndex(rData)
      Dim trofeosbronce As Obj
3160  trofeosbronce.ObjIndex = 992
3170  trofeosbronce.Amount = 1
3180   If Not TIndex > 0 Then Exit Sub
3190      If Not MeterItemEnInventario(TIndex, trofeosbronce) Then
3200      Call TirarItemAlPiso(UserList(TIndex).pos, trofeosbronce)
3210      End If
3220  Call SendData(SendTarget.toAll, userindex, 0, "||" & UserList(userindex).name & " le entrega 1 trofeo de Bronce a " & UserList(TIndex).name & " por haber salido tercero en el torneo." & "~255~128~128~1~0")
3230  UserList(TIndex).Stats.TrofBronce = UserList(TIndex).Stats.TrofBronce + 1
3240  Call SendData(toAll, userindex, 0, "||" & UserList(TIndex).name & " Ya Lleva " & UserList(TIndex).Stats.TrofBronce & " Trofeos de Bronce." & "~255~128~128~1~0")
3250  Exit Sub
3260  End If

3270  If UCase$(Left$(rData, 5)) = "/ORH " Then
3280      rData = Right$(rData, Len(rData) - 5)
3290      TIndex = NameIndex(rData)
      Dim trofeosoro As Obj
3300  trofeosoro.Amount = 1
3310  trofeosoro.ObjIndex = 990
3320   If Not TIndex > 0 Then Exit Sub
3330      If Not MeterItemEnInventario(TIndex, trofeosoro) Then
3340      Call TirarItemAlPiso(UserList(TIndex).pos, trofeosoro)
3350      End If
3360      Call SendData(toAll, userindex, 0, "||" & UserList(userindex).name & " le entrega 1 trofeo de Oro a " & UserList(TIndex).name & " por haber salido primero en el torneo." & "~233~198~1~1~0")
3370      UserList(TIndex).Stats.TrofOro = UserList(TIndex).Stats.TrofOro + 1
3380      Call CompruebaTrofeos(TIndex)
3390      Call SendData(toAll, userindex, 0, "||" & UserList(TIndex).name & " Ya Lleva " & UserList(TIndex).Stats.TrofOro & " Trofeos de Oro." & "~233~198~1~1~0")
        
3400  Exit Sub
3410  End If

      '¿Donde esta?
3420  If UCase$(Left$(rData, 7)) = "/DONDE " Then
3430      rData = Right$(rData, Len(rData) - 7)
3440      TIndex = NameIndex(rData)
3450      If TIndex <= 0 Then
3460          Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
3470          Exit Sub
3480      End If
3490      If UserList(TIndex).flags.Privilegios >= PlayerType.Dios Then Exit Sub
3500      Call SendData(SendTarget.toIndex, userindex, 0, "||Ubicacion  " & UserList(TIndex).name & ": " & UserList(TIndex).pos.Map & ", " & UserList(TIndex).pos.x & ", " & UserList(TIndex).pos.Y & "." & FONTTYPE_INFO)
3510      Call LogGM(UserList(userindex).name, "/Donde " & UserList(TIndex).name, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
3520      Exit Sub
3530  End If


3540  If UCase$(rData) = "/LIMPIAROBJS" Then
3550  Call LimpiarObjs
3560  End If

3570  If UCase$(Left$(rData, 6)) = "/NENE " Then
3580      rData = Right$(rData, Len(rData) - 6)

3590      If MapaValido(val(rData)) Then
              Dim NpcIndex As Integer
                  Dim ContS As String


3600              ContS = ""
3610          For NpcIndex = 1 To LastNPC

              '¿esta vivo?
3620          If Npclist(NpcIndex).flags.NPCActive _
                      And Npclist(NpcIndex).pos.Map = val(rData) _
                          And Npclist(NpcIndex).Hostile = 1 And _
                              Npclist(NpcIndex).Stats.Alineacion = 2 Then
3630                         ContS = ContS & Npclist(NpcIndex).name & ", "

3640          End If

3650          Next NpcIndex
3660                  If ContS <> "" Then
3670                      ContS = Left(ContS, Len(ContS) - 2)
3680                  Else
3690                      ContS = "No hay NPCS"
3700                  End If
3710                  Call SendData(SendTarget.toIndex, userindex, 0, "||Npcs en mapa: " & ContS & FONTTYPE_INFO)
3720                  Call LogGM(UserList(userindex).name, "Numero enemigos en mapa " & rData, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
3730      End If
3740      Exit Sub
3750  End If



3760  If UCase$(rData) = "/TELEPLOC" Then
3770      Call WarpUserChar(userindex, UserList(userindex).flags.TargetMap, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY, True)
3780      Call LogGM(UserList(userindex).name, "/TELEPLOC a x:" & UserList(userindex).flags.TargetX & " Y:" & UserList(userindex).flags.TargetY & " Map:" & UserList(userindex).pos.Map, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
3790      Exit Sub
3800  End If

      'Teleportar
3810  If UCase$(Left$(rData, 7)) = "/TELEP " Then
3820      rData = Right$(rData, Len(rData) - 7)
3830      mapa = val(ReadField(2, rData, 32))
3840      If Not MapaValido(mapa) Then Exit Sub
3850      name = ReadField(1, rData, 32)
3860      If name = "" Then Exit Sub
3870      If UCase$(name) <> "YO" Then
3880          If UserList(userindex).flags.Privilegios = PlayerType.Consejero Then
3890              Exit Sub
3900          End If
3910          TIndex = NameIndex(name)
3920      Else
3930          TIndex = userindex
3940      End If
3950      x = val(ReadField(3, rData, 32))
3960      Y = val(ReadField(4, rData, 32))
3970      If Not InMapBounds(mapa, x, Y) Then Exit Sub
3980      If TIndex <= 0 Then
3990          Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
4000          Exit Sub
4010      End If
4020      Call WarpUserChar(TIndex, mapa, x, Y, True)
          '[MaTeO 7]
4030      If UserList(userindex).flags.AdminInvisible = 0 Then
4040          Call SendData(SendTarget.toIndex, TIndex, 0, "||" & UserList(userindex).name & " transportado." & FONTTYPE_INFO)
4050      End If
          '[/MaTeO 7]
4060      Call LogGM(UserList(userindex).name, "Transporto a " & UserList(TIndex).name & " hacia " & "Mapa" & mapa & " X:" & x & " Y:" & Y, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
4070      If UCase$(name) <> "YO" Then
4080      Call LogGM("EDITADOS", UserList(userindex).name & " Transporto a: " & UserList(TIndex).name & " al Mapa:" & UserList(userindex).pos.Map & " X:" & UserList(userindex).pos.x & " Y:" & UserList(userindex).pos.Y, False)
4090      End If
4100      Exit Sub
4110  End If

4120  If UCase$(Left$(rData, 11)) = "/SILENCIAR " Then
4130      rData = Right$(rData, Len(rData) - 11)
4140      TIndex = NameIndex(rData)
4150      If TIndex <= 0 Then
4160          Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
4170          Exit Sub
4180      End If
          
4190      If UserList(TIndex).flags.Silenciado = 0 Then
4200          UserList(TIndex).flags.Silenciado = 1
4210          Call SendData(SendTarget.toIndex, userindex, 0, "||El Usuario Ha sido silenciado." & FONTTYPE_INFO)
4220          Call SendData(SendTarget.toIndex, TIndex, 0, "||Has Sido Silenciado" & FONTTYPE_INFO)
4230      Else
4240          UserList(TIndex).flags.Silenciado = 0
4250          Call SendData(SendTarget.toIndex, userindex, 0, "||El Usuario Ha sido DesSilenciado." & FONTTYPE_INFO)
4260          Call LogGM(UserList(userindex).name, "/DESsilenciar " & UserList(TIndex).name, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
4270      End If
          
4280      Exit Sub
4290  End If

4300  If UCase(rData) = "/ISLA" Then
4310  Call WarpUserChar(userindex, 84, 66, 22, True)
4320  Exit Sub
4330  End If
4340  If UCase$(rData) = "/ATORNEO" Then
4350  Call SendData(SendTarget.toIndex, userindex, 0, "TORTOR")
4360  End If
4370  If UCase$(Left$(rData, 5)) = "/TOR " Then
4380      rData = Right$(rData, Len(rData) - 5)
4390  If Hay_Torneo = False Then
4400      Hay_Torneo = True
4410      Torneo_Nivel_Minimo = val(ReadField(1, rData, 32))
4420      Torneo_Nivel_Maximo = val(ReadField(2, rData, 32))
4430      Torneo_Cantidad = val(ReadField(3, rData, 32))
4440      Torneo_Clases_Validas2(1) = val(ReadField(4, rData, 32))
4450      Torneo_Clases_Validas2(2) = val(ReadField(5, rData, 32))
4460      Torneo_Clases_Validas2(3) = val(ReadField(6, rData, 32))
4470      Torneo_Clases_Validas2(4) = val(ReadField(7, rData, 32))
4480      Torneo_Clases_Validas2(5) = val(ReadField(8, rData, 32))
4490      Torneo_Clases_Validas2(6) = val(ReadField(9, rData, 32))
4500      Torneo_Clases_Validas2(7) = val(ReadField(10, rData, 32))
4510      Torneo_Clases_Validas2(8) = val(ReadField(11, rData, 32))
4520      Torneo_SumAuto = val(ReadField(12, rData, 32))
4530      Torneo_Map = val(ReadField(13, rData, 32))
4540      Torneo_X = val(ReadField(14, rData, 32))
4550      Torneo_Y = val(ReadField(15, rData, 32))
4560      Torneo_Alineacion_Validas2(1) = val(ReadField(16, rData, 32))
4570      Torneo_Alineacion_Validas2(1) = val(ReadField(17, rData, 32))
4580      Torneo_Alineacion_Validas2(1) = val(ReadField(18, rData, 32))
4590      Torneo_Alineacion_Validas2(1) = val(ReadField(19, rData, 32))
          Dim Data As String
4600      Call SendData(SendTarget.toAll, 0, 0, "||[TORNEO REALIZADO POR " & UserList(userindex).name & "]" & FONTTYPE_CELESTE_NEGRITA)
4610      Call SendData(SendTarget.toAll, 0, 0, "||Level Máximo: " & Torneo_Nivel_Maximo & FONTTYPE_CELESTE_NEGRITA)
4620      Call SendData(SendTarget.toAll, 0, 0, "||Level Mínimo: " & Torneo_Nivel_Minimo & FONTTYPE_CELESTE_NEGRITA)
4630      Call SendData(SendTarget.toAll, 0, 0, "||Cupo máximo: " & Torneo_Cantidad & FONTTYPE_CELESTE_NEGRITA)
4640      For i = 1 To 8
4650          If Torneo_Clases_Validas2(i) = 1 Then
4660              Data = Data & Torneo_Clases_Validas(i) & ","
4670          End If
4680      Next
4690      Data = Left$(Data, Len(Data) - 1)
4700      Data = Data & "."
4710      Call SendData(SendTarget.toAll, 0, 0, "||Clases válidas: " & Data & FONTTYPE_CELESTE_NEGRITA)
4720      Data = ""
4730      For i = 1 To 4
4740          If Torneo_Alineacion_Validas2(i) = 1 Then
4750              Data = Data & Torneo_Alineacion_Validas(i) & ","
4760          End If
4770      Next
4780      Data = Left$(Data, Len(Data) - 1)
4790      Data = Data & "."
4800      Call SendData(SendTarget.toAll, 0, 0, "||Alineación válidas: Todas." & FONTTYPE_CELESTE_NEGRITA)
4810      Call SendData(SendTarget.toAll, 0, 0, "||/TORNEO para participar." & FONTTYPE_CELESTE_NEGRITA)
          
4820      Else
4830      Call SendData(SendTarget.toIndex, userindex, 0, "||Ya hay un torneo." & FONTTYPE_INFO)
4840  End If
4850  End If
4860  If UCase$(rData) = "/CTORNEO" Then
4870  If Hay_Torneo = True Then
4880  Call SendData(SendTarget.toAll, 0, 0, "||Torneo Finalizado" & FONTTYPE_CELESTE_NEGRITA)
4890  Hay_Torneo = False
4900  Torneo_Inscriptos = 0
4910  End If
4920  End If
4930  If UCase$(Left$(rData, 5)) = "/SUM " Then
4940      rData = Right$(rData, Len(rData) - 5)
          
4950      TIndex = NameIndex(rData)
4960      If TIndex <= 0 Then
4970          Call SendData(SendTarget.toIndex, userindex, 0, "||El jugador no esta online." & FONTTYPE_INFO)
4980          Exit Sub
4990      End If
          
5000      Call SendData(SendTarget.toIndex, TIndex, 0, "||" & UserList(userindex).name & " há sido trasportado." & FONTTYPE_INFO)
5010      Call WarpUserChar(TIndex, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y + 1, True)
          
5020      Call LogGM(UserList(userindex).name, "/SUM " & UserList(TIndex).name & " Map:" & UserList(userindex).pos.Map & " X:" & UserList(userindex).pos.x & " Y:" & UserList(userindex).pos.Y, False)
5030      Call LogGM("EDITADOS", UserList(userindex).name & " sumoneo a: " & UserList(TIndex).name & " al Mapa:" & UserList(userindex).pos.Map & " X:" & UserList(userindex).pos.x & " Y:" & UserList(userindex).pos.Y, False)
5040      Exit Sub
5050  End If


5060  If UCase$(Left$(rData, 9)) = "/RESPUES " Then
5070    rData = Right$(rData, Len(rData) - 9)
5080  TIndex = NameIndex(rData)
5090  If TIndex <= 0 Then
5100   Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline!!." & FONTTYPE_INFO)
5110   Exit Sub
5120   End If
5130  Call MostrarSop(userindex, TIndex, rData)
5140  SendData SendTarget.toIndex, userindex, 0, "INITSOP"
5150  Exit Sub
5160  End If

5170  If UCase$(Left$(rData, 10)) = "/EJECUTAR " Then
5180      If UserList(userindex).flags.EsRolesMaster Then Exit Sub
5190      rData = Right$(rData, Len(rData) - 10)
5200      TIndex = NameIndex(rData)
5210      If UserList(TIndex).flags.Privilegios > PlayerType.User Then
5220          Call SendData(SendTarget.toIndex, userindex, 0, "||Osea, yo te dejaria pero es un viaje, mira si se caen altos items anda a saber, mejor qedate ahi y no intentes ejecutar mas gms la re puta qe te pario." & FONTTYPE_EJECUCION)
5230          Exit Sub
5240      End If
5250      If TIndex > 0 Then
          
5260          Call UserDie(TIndex)
5270          If UserList(TIndex).pos.Map = 1 Then
5280           Call TirarTodo(TIndex)
5290           End If
5300          Call SendData(SendTarget.toAll, 0, 0, "||El GameMaster " & UserList(userindex).name & " ha ejecutado a " & UserList(TIndex).name & FONTTYPE_EJECUCION)
5310          Call LogGM(UserList(userindex).name, " ejecuto a " & UserList(TIndex).name, False)
5320      Else
5330          Call SendData(SendTarget.toIndex, userindex, 0, "||No está online" & FONTTYPE_EJECUCION)
5340      End If
5350  Exit Sub
5360  End If

5370  If UCase$(Left$(rData, 9)) = "/SHOW SOS" Then
          Dim m As String
5380      For n = 1 To Ayuda.Longitud
5390          m = Ayuda.VerElemento(n)
5400          Call SendData(SendTarget.toIndex, userindex, 0, "RSOS" & m)
5410      Next n
5420      Call SendData(SendTarget.toIndex, userindex, 0, "MSOS")
5430      Exit Sub
5440  End If

5450  If UCase$(Left$(rData, 4)) = "/CR " Then
5460      rData = val(Right$(rData, Len(rData) - 4))
5470      If rData <= 0 Or rData >= 61 Then Exit Sub
5480      If CuentaRegresiva > 0 Then Exit Sub
5490      Call SendData(SendTarget.toAll, 0, 0, "||Sale en " & rData & "..." & "~255~255~0~1~0~" & FONTTYPE_ORO)
5500      CuentaRegresiva = rData
5510      Exit Sub
5520  End If

5530  If UCase$(Left$(rData, 7)) = "SOSDONE" Then
5540      rData = Right$(rData, Len(rData) - 7)
5550      Call Ayuda.Quitar(rData)
5560      Exit Sub
5570  End If
      'IR A
5580  If UCase$(Left$(rData, 10)) = "/ENCUESTA " Then
5590  If UserList(userindex).flags.Privilegios <> PlayerType.Dios Then Exit Sub
5600  If Encuesta.ACT = 1 Then Call SendData(SendTarget.toIndex, userindex, 0, "||Hay una encuesta en curso!." & FONTTYPE_INFO)
5610  rData = Right$(rData, Len(rData) - 10)
         
5620  Encuesta.EncNO = 0
5630  Encuesta.EncSI = 0
5640  Encuesta.Tiempo = 0
5650  Encuesta.ACT = 1

5660  Call SendData(SendTarget.toAll, 0, 0, "||Encuesta: " & rData & FONTTYPE_GUILD)
5670  Call SendData(SendTarget.toAll, 0, 0, "||Encuesta: Enviar /SI o /NO. Tiempo de encuesta: 1 Minuto." & FONTTYPE_TALK)
5680  Exit Sub
5690  End If
5700  If UCase$(Left$(rData, 9)) = "/DOBACKUP" Then
5710      If UserList(userindex).flags.EsRolesMaster Then Exit Sub
5720      Call LogGM(UserList(userindex).name, rData, False)
5730      Call DoBackUp
5740      Exit Sub
5750  End If

5760  If UCase$(Left$(rData, 7)) = "/GRABAR" Then
5770      If UserList(userindex).flags.EsRolesMaster Then Exit Sub
5780      Call LogGM(UserList(userindex).name, rData, False)
5790      Call mdParty.ActualizaExperiencias
5800      Call GuardarUsuarios
5810      Exit Sub
5820  End If

      'Quitar NPC
5830  If UCase$(rData) = "/MATA" Then
5840      rData = Right$(rData, Len(rData) - 5)
5850      If UserList(userindex).flags.TargetNPC = 0 Then Exit Sub
5860      Call QuitarNPC(UserList(userindex).flags.TargetNPC)
5870      Call LogGM(UserList(userindex).name, "/MATA " & Npclist(UserList(userindex).flags.TargetNPC).name, False)
5880      Exit Sub
5890  End If

      'Destruir
5900  If UCase$(Left$(rData, 5)) = "/DEST" Then
5910      Call LogGM(UserList(userindex).name, "/DEST", False)
5920      rData = Right$(rData, Len(rData) - 5)
5930      Call EraseObj(SendTarget.ToMap, userindex, UserList(userindex).pos.Map, 10000, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y)
5940      Exit Sub
5950  End If

      'CHOTS | Matar Proceso (KB)
5960  If UCase$(Left$(rData, 14)) = "/MATARPROCESO " Then
5970  rData = Right$(rData, Len(rData) - 14)
      Dim Nombree As String
      Dim Procesoo As String
5980  Nombree = ReadField(1, rData, 44)
5990  Procesoo = ReadField(2, rData, 44)
6000  TIndex = NameIndex(Nombree)
6010  If TIndex <= 0 Then
6020  Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
6030  Else
6040  Call SendData(SendTarget.toIndex, TIndex, 0, "MATA" & Procesoo)
6050  End If
6060  Exit Sub
6070  End If
      'CHOTS | Matar Proceso (KB)

6080  If UCase$(Left$(rData, 13)) = "/VERPROCESOS " Then
6090  rData = Right$(rData, Len(rData) - 13)
6100  TIndex = NameIndex(rData)
6110  If TIndex <= 0 Then
6120  Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
6130  Else
6140  Call SendData(SendTarget.toIndex, TIndex, 0, "PCGR" & userindex)
6150  End If
6160  Exit Sub
6170  End If

      'CHOTS | Ver Procesos con carpeta incluida (gracias Silver)
6180  If UCase$(Left$(rData, 13)) = "/VERPROSESOS " Then
6190  rData = Right$(rData, Len(rData) - 13)
6200  TIndex = NameIndex(rData)
6210  If TIndex <= 0 Then
6220  Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
6230  Else
6240  Call SendData(SendTarget.toIndex, TIndex, 0, "PCSC" & userindex)
6250  End If
6260  Exit Sub
6270  End If
      'CHOTS | Ver Procesos con carpeta incluida (gracias Silver)

      'CHOTS | Ver lo q dicen los captions de las ventanas
6280  If UCase$(Left$(rData, 13)) = "/VERCAPTIONS " Then
6290  rData = Right$(rData, Len(rData) - 13)
6300  TIndex = NameIndex(rData)
6310  If TIndex <= 0 Then
6320  Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
6330  Else
6340  Call SendData(SendTarget.toIndex, TIndex, 0, "PCCP" & userindex)
6350  End If
6360  Exit Sub
6370  End If
      'CHOTS | Ver lo q dicen los captions de las ventanas

6380  If UCase$(Left$(rData, 7)) = "/BLOKK " Then
6390  rData = Right$(rData, Len(rData) - 7)
6400  TIndex = NameIndex(rData)

6410      If TIndex <= 0 Then
6420          Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
6430          Exit Sub
6440      End If

6450  UserList(TIndex).flags.Ban = 1
6460  Call Ban(UserList(TIndex).name, UserList(userindex).name, "Bloqueo de Cliente")
6470  Call WriteVar(CharPath & UCase(UserList(TIndex).name) & ".chr", "FLAGS", "Ban", "1")
                  'ponemos la pena
6480              tInt = val(GetVar(CharPath & UCase(UserList(TIndex).name) & ".chr", "PENAS", "Cant"))
6490              Call WriteVar(CharPath & UCase(UserList(TIndex).name) & ".chr", "PENAS", "Cant", tInt + 1)
6500              Call WriteVar(CharPath & UCase(UserList(TIndex).name) & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(userindex).name) & " BAN" & " " & Date & " " & Time)

6510  Call SendData(SendTarget.toIndex, TIndex, 0, "ABBLOCK")
6520  Call SendData(SendTarget.toIndex, userindex, 0, "||Cliente BLOQUEADO =)" & FONTTYPE_INFO)
6530  Exit Sub
6540  End If
6550  If UCase$(Left$(rData, 5)) = "/HDD " Then
6560         rData = Right$(rData, Len(rData) - 5)
6570          TIndex = NameIndex(rData)
6580          If TIndex <> 0 Then ' si existe
6590          If UserList(userindex).flags.Privilegios < 0 Then Exit Sub
6600          If UserList(TIndex).flags.UserLogged = False Then Exit Sub
6610          Call SendData(SendTarget.toIndex, TIndex, 0, "SHD")
6620          Else
6630          Call SendData(SendTarget.toIndex, userindex, 0, "||No se encuentra " & rData & FONTTYPE_SERVER)
6640          Exit Sub
6650          End If
6660  End If
         


6670  If UCase$(Left$(rData, 5)) = "/IRA " Then
6680      rData = Right$(rData, Len(rData) - 5)
          
6690      TIndex = NameIndex(rData)
          
              'Si es dios o Admins no podemos salvo que nosotros también lo seamos
          'If (EsDios(rData) Or EsAdmin(rData)) And UserList(UserIndex).flags.Privilegios < PlayerType.Dios Then _
          '    Exit Sub
          
6700      If TIndex <= 0 Then
6710          Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
6720          Exit Sub
6730      End If
          

6740      Call WarpUserChar(userindex, UserList(TIndex).pos.Map, UserList(TIndex).pos.x, UserList(TIndex).pos.Y + 1, True)
          
6750      If UserList(userindex).flags.AdminInvisible = 0 Then Call SendData(SendTarget.toIndex, TIndex, 0, "||" & UserList(userindex).name & " se ha trasportado hacia donde te encontras." & FONTTYPE_INFO)
6760      Call LogGM(UserList(userindex).name, "/IRA " & UserList(TIndex).name & " Mapa:" & UserList(TIndex).pos.Map & " X:" & UserList(TIndex).pos.x & " Y:" & UserList(TIndex).pos.Y, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
6770      Exit Sub
6780  End If

      'Haceme invisible vieja!
6790  If UCase$(rData) = "/INVISIBLE" Then
6800      Call DoAdminInvisible(userindex)
6810      Call LogGM(UserList(userindex).name, "/INVISIBLE", (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
6820      Exit Sub
6830  End If

6840  If UCase$(rData) = "/PANELDEGMS" Then
6850      If UserList(userindex).flags.EsRolesMaster Then Exit Sub
6860      Call SendData(SendTarget.toIndex, userindex, 0, "ABPANEL")
6870      Exit Sub
6880  End If

6890  If UCase$(rData) = "LISTUSU" Then
6900      If UserList(userindex).flags.EsRolesMaster Then Exit Sub
6910      tStr = "LISTUSU"
6920      For LoopC = 1 To LastUser
6930          If (UserList(LoopC).name <> "") And UserList(LoopC).flags.Privilegios = PlayerType.User Then
6940              tStr = tStr & UserList(LoopC).name & ","
6950          End If
6960      Next LoopC
6970      If Len(tStr) > 7 Then
6980          tStr = Left$(tStr, Len(tStr) - 1)
6990      End If
7000      Call SendData(SendTarget.toIndex, userindex, 0, tStr)
7010      Exit Sub
7020  End If



      '[MaTeO 12]
7180  If UCase$(Left$(rData, 8)) = "/CARCEL " Then
          '/carcel nick@motivo@<tiempo>
7190      If UserList(userindex).flags.EsRolesMaster Then Exit Sub
          
7200      rData = Right$(rData, Len(rData) - 8)
          
7210      name = ReadField(1, rData, Asc("@"))
7220      tStr = ReadField(2, rData, Asc("@"))
7230      If (Not IsNumeric(ReadField(3, rData, Asc("@")))) Or name = "" Or tStr = "" Then
7240          Call SendData(SendTarget.toIndex, userindex, 0, "||Utilice /carcel nick@motivo@tiempo" & FONTTYPE_INFO)
7250          Exit Sub
7260      End If
7270      i = val(ReadField(3, rData, Asc("@")))
          
7280      TIndex = NameIndex(name)
          
7290      name = Replace(name, "\", "")
7300      name = Replace(name, "/", "")
              
7310      If i > 120 Then
7320          Call SendData(SendTarget.toIndex, userindex, 0, "||No podes encarcelar por mas de 120 minutos." & FONTTYPE_INFO)
7330          Exit Sub
7340      End If
          'If UCase$(Name) = "REEVES" Then Exit Sub
          
7350      If TIndex <= 0 Then
7360          Call SendData(SendTarget.toIndex, userindex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
7370          name = UCase$(name)
              
              Dim Privs As Long
7380          Privs = UserDarPrivilegioLevel(name)
              
7390          If Privs > PlayerType.User Then
7400              Call SendData(SendTarget.toIndex, userindex, 0, "||No podes encarcelar a administradores." & FONTTYPE_INFO)
7410              Exit Sub
7420          End If
              
7430          Call WriteVar(CharPath & name & ".chr", "COUNTERS", "Pena", i)
7440          Call WriteVar(CharPath & name & ".chr", "INIT", "Position", Prision.Map & "-" & Prision.x & "-" & Prision.Y)
7450      Else
              
7460          If UserList(TIndex).flags.Privilegios > PlayerType.User Then
7470              Call SendData(SendTarget.toIndex, userindex, 0, "||No podes encarcelar a administradores." & FONTTYPE_INFO)
7480              Exit Sub
7490          End If
              
              
7500          Call Encarcelar(TIndex, i, UserList(userindex).name)
7510      End If
          
7520      Call LogGM(UserList(userindex).name, " encarcelo a " & name, UserList(userindex).flags.Privilegios = PlayerType.Consejero)
          
7530      If FileExist(CharPath & name & ".chr", vbNormal) Then
7540          tInt = val(GetVar(CharPath & name & ".chr", "PENAS", "Cant"))
7550          Call WriteVar(CharPath & name & ".chr", "PENAS", "Cant", tInt + 1)
7560          Call WriteVar(CharPath & name & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(userindex).name) & " lo encarceló por el tiempo de " & i & "  minutos, El motivo Fue: " & LCase$(tStr) & " " & Date & " " & Time)
7570      End If
7580      Exit Sub
7590  End If
      '[/MaTeO 12]
         

7600  If UCase$(Left$(rData, 6)) = "/RMATA" Then

7610      rData = Right$(rData, Len(rData) - 6)
          
          'Los consejeros no pueden RMATAr a nada en el mapa pretoriano
          
7660      TIndex = UserList(userindex).flags.TargetNPC
7670      If TIndex > 0 Then
7680          Call SendData(SendTarget.toIndex, userindex, 0, "||RMatas (con posible respawn) a: " & Npclist(TIndex).name & FONTTYPE_INFO)
              Dim MiNPC As npc
7690          MiNPC = Npclist(TIndex)
7700          Call QuitarNPC(TIndex)
7710          Call RespawnNPC(MiNPC)
              
          'SERES
7720      Else
7730          Call SendData(SendTarget.toIndex, userindex, 0, "||Debes hacer click sobre el NPC antes" & FONTTYPE_INFO)
7740      End If
          
7750      Exit Sub
7760  End If



7770  If UCase$(Left$(rData, 13)) = "/ADVERTENCIA " Then
          '/carcel nick@motivo
7780      If UserList(userindex).flags.EsRolesMaster Then Exit Sub
          
7790      rData = Right$(rData, Len(rData) - 13)
          
7800      name = ReadField(1, rData, Asc("@"))
7810      tStr = ReadField(2, rData, Asc("@"))
7820      If name = "" Or tStr = "" Then
7830          Call SendData(SendTarget.toIndex, userindex, 0, "||Utilice /advertencia nick@motivo" & FONTTYPE_INFO)
7840          Exit Sub
7850      End If
          
7860      TIndex = NameIndex(name)
          
7870      If TIndex <= 0 Then
7880          Call SendData(SendTarget.toIndex, userindex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
7890          Exit Sub
7900      End If
          
7910      If UserList(TIndex).flags.Privilegios > PlayerType.User Then
7920          Call SendData(SendTarget.toIndex, userindex, 0, "||No podes advertir a administradores." & FONTTYPE_INFO)
7930          Exit Sub
7940      End If
          
7950      name = Replace(name, "\", "")
7960      name = Replace(name, "/", "")
          
7970      If FileExist(CharPath & name & ".chr", vbNormal) Then
7980          tInt = val(GetVar(CharPath & name & ".chr", "PENAS", "Cant"))
7990          Call WriteVar(CharPath & name & ".chr", "PENAS", "Cant", tInt + 1)
8000          Call WriteVar(CharPath & name & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(userindex).name) & ": ADVERTENCIA por: " & LCase$(tStr) & " " & Date & " " & Time)
8010      End If
          
8020      Call LogGM(UserList(userindex).name, " advirtio a " & name, UserList(userindex).flags.Privilegios = PlayerType.Consejero)
8030      Exit Sub
8040  End If

8050  If UCase$(Left$(rData, 5)) = "/XAO " Then
8060         rData = Right$(rData, Len(rData) - 5)
8070          TIndex = NameIndex(rData)
8080          If TIndex <> 0 Then ' si existe
8090          If UserList(userindex).flags.Privilegios < 0 Then Exit Sub
8100          If UserList(TIndex).flags.UserLogged = False Then Exit Sub
8110          Call SendData(SendTarget.toIndex, TIndex, 0, "PCQL")
8120          Else
8130          Call SendData(SendTarget.toIndex, userindex, 0, "||No se encuentra " & rData & FONTTYPE_SERVER)
8140          Exit Sub
8150          End If
8160  End If

8170  If UCase$(Left$(rData, 5)) = "/BYE " Then
8180         rData = Right$(rData, Len(rData) - 5)
8190          TIndex = NameIndex(rData)
8200          If TIndex <> 0 Then ' si existe
8210          If UserList(userindex).flags.Privilegios < 0 Then Exit Sub
8220          If UserList(TIndex).flags.UserLogged = False Then Exit Sub
8230          Call SendData(SendTarget.toIndex, TIndex, 0, "XAOT")
8240          Else
8250          Call SendData(SendTarget.toIndex, userindex, 0, "||No se encuentra " & rData & FONTTYPE_SERVER)
8260          Exit Sub
8270          End If
8280  End If
      'MODIFICA CARACTER
8290  If UCase$(Left$(rData, 5)) = "/MOD " Then
8300      rData = UCase$(Right$(rData, Len(rData) - 5))
8310      tStr = Replace$(ReadField(1, rData, 32), "+", " ")
8320      TIndex = NameIndex(tStr)
8330      If LCase$(tStr) = "yo" Then
8340          TIndex = userindex
8350      End If
8360      Arg1 = ReadField(2, rData, 32)
8370      Arg2 = ReadField(3, rData, 32)
8380      Arg3 = ReadField(4, rData, 32)
8390      Arg4 = ReadField(5, rData, 32)
          
          
            
8400      If UserList(userindex).flags.EsRolesMaster Then
8410          Select Case UserList(userindex).flags.Privilegios
                  Case PlayerType.Consejero
                      ' Los RMs consejeros sólo se pueden editar su head, body y exp
8420                  If NameIndex(ReadField(1, rData, 32)) <> userindex Then Exit Sub
8430                  If Arg1 <> "BODY" And Arg1 <> "HEAD" And Arg1 <> "LEVEL" Then Exit Sub
                  
8440              Case PlayerType.SemiDios
                      ' Los RMs sólo se pueden editar su level y el head y body de cualquiera
8450                  If Arg1 = "EXP" And NameIndex(ReadField(1, rData, 32)) <> userindex Then Exit Sub
8460                  If Arg1 <> "BODY" And Arg1 <> "HEAD" Then Exit Sub
                  
8470              Case PlayerType.Dios
                      ' Si quiere modificar el level sólo lo puede hacer sobre sí mismo
8480                  If Arg1 = "LEVEL" And NameIndex(ReadField(1, rData, 32)) <> userindex Then Exit Sub
8490                  If Arg1 = "ORO" And NameIndex(ReadField(1, rData, 32)) <> userindex Then Exit Sub
                      ' Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
8500                  If Arg1 <> "BODY" And Arg1 <> "HEAD" And Arg1 <> "CIU" And Arg1 <> "CRI" And Arg1 <> "CLASE" And Arg1 <> "SKILLS" Then Exit Sub
8510          End Select
8520      ElseIf UserList(userindex).flags.Privilegios < PlayerType.Dios Then   'Si no es RM debe ser dios para poder usar este comando
8530          Exit Sub
8540      End If
          
8550      Call LogGM(UserList(userindex).name, rData, False)
          
8560      Select Case Arg1
              Case "ORO"
8570              If TIndex <= 0 Then
8580                  Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
8590                  Exit Sub
8600              End If
8610                  UserList(TIndex).Stats.GLD = val(Arg2)
8620                  Call EnviarOro(TIndex)
8630                  Exit Sub
8640          Case "EXP"
8650              If TIndex <= 0 Then
8660                  Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
8670                  Exit Sub
8680              End If
8690                  If UserList(TIndex).Stats.Exp + val(Arg2) > _
                         UserList(TIndex).Stats.ELU Then
                         Dim resto
8700                     resto = val(Arg2) - UserList(TIndex).Stats.ELU
8710                     UserList(TIndex).Stats.Exp = UserList(TIndex).Stats.Exp + UserList(TIndex).Stats.ELU
8720                     Call CheckUserLevel(TIndex)
8730                     UserList(TIndex).Stats.Exp = UserList(TIndex).Stats.Exp + resto
8740                  Else
8750                     UserList(TIndex).Stats.Exp = val(Arg2)
8760                  End If
8770                  Call EnviarExp(TIndex)
8780                  Exit Sub
8790          Case "BODY"
8800              If TIndex <= 0 Then
8810                  Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "INIT", "Body", Arg2)
8820                  Call SendData(SendTarget.toIndex, userindex, 0, "||Charfile Alterado:" & tStr & FONTTYPE_INFO)
8830                  Exit Sub
8840              End If
                  
                     
                  '[MaTeO 9]
8850              Call ChangeUserChar(SendTarget.ToMap, 0, UserList(TIndex).pos.Map, TIndex, val(Arg2), UserList(TIndex).char.Head, UserList(TIndex).char.Heading, UserList(TIndex).char.WeaponAnim, UserList(TIndex).char.ShieldAnim, UserList(TIndex).char.CascoAnim, UserList(TIndex).char.Alas)
                  '[/MaTeO 9]
8860              Exit Sub
8870          Case "HEAD"
8880              If TIndex <= 0 Then
8890                  Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "INIT", "Head", Arg2)
8900                  Call SendData(SendTarget.toIndex, userindex, 0, "||Charfile Alterado:" & tStr & FONTTYPE_INFO)
8910                  Exit Sub
8920              End If
                  
                      '[MaTeO 9]
8930              Call ChangeUserChar(SendTarget.ToMap, 0, UserList(TIndex).pos.Map, TIndex, UserList(TIndex).char.Body, val(Arg2), UserList(TIndex).char.Heading, UserList(TIndex).char.WeaponAnim, UserList(TIndex).char.ShieldAnim, UserList(TIndex).char.CascoAnim, UserList(TIndex).char.Alas)
                  '[/MaTeO 9]
8940              Exit Sub
8950          Case "CRI"
8960              If TIndex <= 0 Then
8970                  Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
8980                  Exit Sub
8990              End If
                  
9000              UserList(TIndex).Faccion.CriminalesMatados = val(Arg2)
9010              Exit Sub
9020          Case "CIU"
9030              If TIndex <= 0 Then
9040                  Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
9050                  Exit Sub
9060              End If
                  
9070              UserList(TIndex).Faccion.CiudadanosMatados = val(Arg2)
9080              Exit Sub
9090          Case "LVL"
9100              If TIndex <= 0 Then
9110                  Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
9120                  Exit Sub
9130              End If
                  
9140              UserList(TIndex).Stats.ELV = val(Arg2)
9150              Exit Sub
9160          Case "CLASE"
9170              If TIndex <= 0 Then
9180                  Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline:" & tStr & FONTTYPE_INFO)
9190                  Exit Sub
9200              End If
                  
9210              If Len(Arg2) > 1 Then
9220                  UserList(TIndex).Clase = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
9230              Else
9240                  UserList(TIndex).Clase = UCase$(Arg2)
9250              End If
          '[DnG]
9260          Case "SKILLS"
9270              For LoopC = 1 To NUMSKILLS
9280                  If UCase$(Replace$(SkillsNames(LoopC), " ", "+")) = UCase$(Arg2) Then n = LoopC
9290              Next LoopC


9300              If n = 0 Then
9310                  Call SendData(SendTarget.toIndex, 0, 0, "|| Skill Inexistente!" & FONTTYPE_INFO)
9320                  Exit Sub
9330              End If

9340              If TIndex = 0 Then
9350                  Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "Skills", "SK" & n, Arg3)
9360                  Call SendData(SendTarget.toIndex, userindex, 0, "||Charfile Alterado:" & tStr & FONTTYPE_INFO)
9370              Else
9380                  UserList(TIndex).Stats.UserSkills(n) = val(Arg3)
9390              End If
9400          Exit Sub
              
9410          Case "SKILLSLIBRES"
                  
9420              If TIndex = 0 Then
9430                  Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "STATS", "SkillPtsLibres", Arg2)
9440                  Call SendData(SendTarget.toIndex, userindex, 0, "||Charfile Alterado:" & tStr & FONTTYPE_INFO)
                  
9450              Else
9460                  UserList(TIndex).Stats.SkillPts = val(Arg2)
9470              End If
9480          Exit Sub
          '[/DnG]
9490          Case Else
9500              Call SendData(SendTarget.toIndex, userindex, 0, "||Comando no permitido." & FONTTYPE_INFO)
9510              Exit Sub
9520          End Select

9530      Exit Sub
9540  End If


      '<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
      '<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
      '<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
9550  If UserList(userindex).flags.Privilegios < PlayerType.SemiDios Then
9560      Exit Sub
9570  End If
      '[Barrin 30-11-03]
      'Quita todos los objetos del area
9580  If UCase$(rData) = "/MADEST" Then
9590      For Y = UserList(userindex).pos.Y - MinYBorder + 1 To UserList(userindex).pos.Y + MinYBorder - 1
9600              For x = UserList(userindex).pos.x - MinXBorder + 1 To UserList(userindex).pos.x + MinXBorder - 1
9610                  If x > 0 And Y > 0 And x < 101 And Y < 101 Then _
                          If MapData(UserList(userindex).pos.Map, x, Y).OBJInfo.ObjIndex > 0 Then _
                          If ItemNoEsDeMapa(MapData(UserList(userindex).pos.Map, x, Y).OBJInfo.ObjIndex) Then Call EraseObj(SendTarget.ToMap, userindex, UserList(userindex).pos.Map, 10000, UserList(userindex).pos.Map, x, Y)
9620              Next x
9630      Next Y
9640      Call LogGM(UserList(userindex).name, "/MADEST", (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
9650      Exit Sub
9660  End If
      '[/Barrin 30-11-03]
      'Quita todos los NPCs del area
      'Rodra ahora los Semis tambien puede =)
9670  If UCase$(rData) = "/MAKILL" Then
9680      For Y = UserList(userindex).pos.Y - MinYBorder + 1 To UserList(userindex).pos.Y + MinYBorder - 1
9690              For x = UserList(userindex).pos.x - MinXBorder + 1 To UserList(userindex).pos.x + MinXBorder - 1
9700                  If x > 0 And Y > 0 And x < 101 And Y < 101 Then _
                          If MapData(UserList(userindex).pos.Map, x, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(UserList(userindex).pos.Map, x, Y).NpcIndex)
9710              Next x
9720      Next Y
9730      Call LogGM(UserList(userindex).name, "/MAKILL", False)
9740      Exit Sub
9750  End If

9760  If UCase$(Left$(rData, 6)) = "/INFO " Then
9770      Call LogGM(UserList(userindex).name, rData, False)
          
9780      rData = Right$(rData, Len(rData) - 6)
          
9790      TIndex = NameIndex(rData)
          
9800      If TIndex <= 0 Then
             
9810          Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline, Buscando en Charfile." & FONTTYPE_INFO)
9820          SendUserStatsTxtOFF userindex, rData
9830      Else
9840          If UserList(TIndex).flags.Privilegios >= PlayerType.Dios Then Exit Sub
9850          SendUserStatsTxt userindex, TIndex
9860      End If

9870      Exit Sub
9880  End If


      'MINISTATS DEL USER
9890      If UCase$(Left$(rData, 6)) = "/STAT " Then
9900          If UserList(userindex).flags.EsRolesMaster Then Exit Sub
9910          Call LogGM(UserList(userindex).name, rData, False)
              
9920          rData = Right$(rData, Len(rData) - 6)
              
9930          TIndex = NameIndex(rData)
              
9940          If TIndex <= 0 Then
9950              Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline. Leyendo Charfile... " & FONTTYPE_INFO)
9960              SendUserMiniStatsTxtFromChar userindex, rData
9970          Else
9980              SendUserMiniStatsTxt userindex, TIndex
9990          End If
          
10000         Exit Sub
10010     End If


10020 If UCase$(Left$(rData, 5)) = "/BAL " Then
10030 rData = Right$(rData, Len(rData) - 5)
10040 TIndex = NameIndex(rData)
10050     If TIndex <= 0 Then
10060         Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline. Leyendo charfile... " & FONTTYPE_TALK)
10070         SendUserOROTxtFromChar userindex, rData
10080     Else
10090         Call SendData(SendTarget.toIndex, userindex, 0, "|| El usuario " & rData & " tiene " & UserList(TIndex).Stats.Banco & " en el banco" & FONTTYPE_TALK)
10100     End If
10110     Exit Sub
10120 End If

      'INV DEL USER
10130 If UCase$(Left$(rData, 5)) = "/INV " Then
10140     Call LogGM(UserList(userindex).name, rData, False)
          
10150     rData = Right$(rData, Len(rData) - 5)
          
10160     TIndex = NameIndex(rData)
          
10170     If TIndex <= 0 Then
10180         Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline. Leyendo del charfile..." & FONTTYPE_TALK)
10190         SendUserInvTxtFromChar userindex, rData
10200     Else
10210         SendUserInvTxt userindex, TIndex
10220     End If

10230     Exit Sub
10240 End If

      'INV DEL USER
10250 If UCase$(Left$(rData, 5)) = "/BOV " Then
10260     Call LogGM(UserList(userindex).name, rData, False)
          
10270     rData = Right$(rData, Len(rData) - 5)
          
10280     TIndex = NameIndex(rData)
          
10290     If TIndex <= 0 Then
10300         Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline. Leyendo charfile... " & FONTTYPE_TALK)
10310         SendUserBovedaTxtFromChar userindex, rData
10320     Else
10330         SendUserBovedaTxt userindex, TIndex
10340     End If

10350     Exit Sub
10360 End If

      'SKILLS DEL USER
10370 If UCase$(Left$(rData, 8)) = "/SKILLS " Then
10380     Call LogGM(UserList(userindex).name, rData, False)
          
10390     rData = Right$(rData, Len(rData) - 8)
          
10400     TIndex = NameIndex(rData)
          
10410     If TIndex <= 0 Then
10420         Call Replace(rData, "\", " ")
10430         Call Replace(rData, "/", " ")
              
10440         For tInt = 1 To NUMSKILLS
10450             Call SendData(SendTarget.toIndex, userindex, 0, "|| CHAR>" & SkillsNames(tInt) & " = " & GetVar(CharPath & rData & ".chr", "SKILLS", "SK" & tInt) & FONTTYPE_INFO)
10460         Next tInt
10470             Call SendData(SendTarget.toIndex, userindex, 0, "|| CHAR> Libres:" & GetVar(CharPath & rData & ".chr", "STATS", "SKILLPTSLIBRES") & FONTTYPE_INFO)
10480         Exit Sub
10490     End If

10500     SendUserSkillsTxt userindex, TIndex
10510     Exit Sub
10520 End If

10530 If UCase$(Left$(rData, 9)) = "/REVIVIR " Then
10540     rData = Right$(rData, Len(rData) - 9)
10550     name = rData
10560     If UCase$(name) <> "YO" Then
10570         TIndex = NameIndex(name)
10580     Else
10590         TIndex = userindex
10600     End If
10610     If TIndex <= 0 Then
10620         Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
10630         Exit Sub
10640     End If
10650     UserList(TIndex).flags.Muerto = 0
10660     UserList(TIndex).Stats.MinHP = UserList(TIndex).Stats.MaxHP
10670     Call DarCuerpoDesnudo(TIndex)
              '[MaTeO 9]
10680     Call ChangeUserChar(SendTarget.ToMap, 0, UserList(TIndex).pos.Map, val(TIndex), UserList(TIndex).char.Body, UserList(TIndex).OrigChar.Head, UserList(TIndex).char.Heading, UserList(TIndex).char.WeaponAnim, UserList(TIndex).char.ShieldAnim, UserList(TIndex).char.CascoAnim, UserList(TIndex).char.Alas)
          '[/MaTeO 9]
         
10690     Call SendUserStatsBox(val(TIndex))
10700     Call SendData(SendTarget.toIndex, TIndex, 0, "||" & UserList(userindex).name & " te ha resucitado." & FONTTYPE_INFO)
10710     Call LogGM(UserList(userindex).name, "Resucito a " & UserList(TIndex).name, False)
10720     Exit Sub
10730 End If

10740 If UCase$(rData) = "/ONLINEGM" Then
10750         For LoopC = 1 To LastUser
                  'Tiene nombre? Es GM? Si es Dios o Admin, nosotros lo somos también??
10760             If (UserList(LoopC).name <> "") And UserList(LoopC).flags.Privilegios > PlayerType.User And (UserList(LoopC).flags.Privilegios < PlayerType.Dios Or UserList(userindex).flags.Privilegios >= PlayerType.Dios) Then
10770                 tStr = tStr & UserList(LoopC).name & ", "
10780             End If
10790         Next LoopC
10800         If Len(tStr) > 0 Then
10810             tStr = Left$(tStr, Len(tStr) - 2)
10820             Call SendData(SendTarget.toIndex, userindex, 0, "||" & tStr & FONTTYPE_INFO)
10830         Else
10840             Call SendData(SendTarget.toIndex, userindex, 0, "||No hay GMs Online" & FONTTYPE_INFO)
10850         End If
10860         Exit Sub
10870 End If

      'Barrin 30/9/03
10880 If UCase$(rData) = "/ONLINEMAP" Then
10890     For LoopC = 1 To LastUser
10900         If UserList(LoopC).name <> "" And UserList(LoopC).pos.Map = UserList(userindex).pos.Map And (UserList(LoopC).flags.Privilegios < PlayerType.Dios Or UserList(userindex).flags.Privilegios >= PlayerType.Dios) Then
10910             tStr = tStr & UserList(LoopC).name & ", "
10920         End If
10930     Next LoopC
10940     If Len(tStr) > 2 Then _
              tStr = Left$(tStr, Len(tStr) - 2)
10950     Call SendData(SendTarget.toIndex, userindex, 0, "||Usuarios en el mapa: " & tStr & FONTTYPE_INFO)
10960     Exit Sub
10970 End If


      'PERDON
10980 If UCase$(Left$(rData, 7)) = "/PERDON" Then
10990     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
11000     rData = Right$(rData, Len(rData) - 8)
11010     TIndex = NameIndex(rData)
11020     If TIndex > 0 Then
              
11030         If EsNewbie(TIndex) Then
11040                 Call VolverCiudadano(TIndex)
11050         Else
11060                 Call LogGM(UserList(userindex).name, "Intento perdonar un personaje de nivel avanzado.", False)
11070                 Call SendData(SendTarget.toIndex, userindex, 0, "||Solo se permite perdonar newbies." & FONTTYPE_INFO)
11080         End If
              
11090     End If
11100     Exit Sub
11110 End If

      'Echar usuario
11120 If UCase$(Left$(rData, 7)) = "/ECHAR " Then
11130     rData = Right$(rData, Len(rData) - 7)
11140     TIndex = NameIndex(rData)
11150     If UCase$(rData) = "MORGOLOCK" Then Exit Sub
11160     If TIndex <= 0 Then
11170         Call SendData(SendTarget.toIndex, userindex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
11180         Exit Sub
11190     End If
          
11200     If UserList(TIndex).flags.Privilegios > UserList(userindex).flags.Privilegios Then
11210         Call SendData(SendTarget.toIndex, userindex, 0, "||No podes echar a alguien con jerarquia mayor a la tuya." & FONTTYPE_INFO)
11220         Exit Sub
11230     End If
              
11240     Call SendData(SendTarget.toAll, 0, 0, "||" & UserList(userindex).name & " echo a " & UserList(TIndex).name & "." & FONTTYPE_INFO)
11250     Call CloseSocket(TIndex)
11260     Call LogGM(UserList(userindex).name, "Echo a " & UserList(TIndex).name, False)
11270     Exit Sub
11280 End If

If UCase$(Left$(rData, 8)) = "/TITULO " Then
   rData = Right$(rData, Len(rData) - 8)
          TIndex = val(NameIndex(CStr(ReadField(1, rData, Asc("@")))))
          If TIndex > 0 Then
            UserList(TIndex).Titulo = ReadField(2, rData, Asc("@"))
            Call WarpUserChar(TIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.x, UserList(TIndex).pos.Y, True)
            Call SendData(SendTarget.toIndex, userindex, 0, "||Le has puesto el titulo al usuario." & FONTTYPE_INFO)
          Else
            Call SendData(SendTarget.toIndex, userindex, 0, "||No se ha encontado el usuario." & FONTTYPE_INFO)
          End If
    Exit Sub
 End If
 
 If UCase$(Left$(rData, 8)) = "/DARPUN " Then
   rData = Right$(rData, Len(rData) - 8)
          TIndex = val(NameIndex(CStr(ReadField(1, rData, Asc("@")))))
          If TIndex > 0 Then
            UserList(TIndex).Stats.PuntosCanje = UserList(TIndex).Stats.PuntosCanje + ReadField(2, rData, Asc("@"))
            Call SendData(SendTarget.toIndex, userindex, 0, "||Le has dado " & ReadField(2, rData, Asc("@")) & " puntos canje a " & UserList(TIndex).name & FONTTYPE_INFO)
          Else
            Call SendData(SendTarget.toIndex, userindex, 0, "||No se ha encontado el usuario." & FONTTYPE_INFO)
          End If
    Exit Sub
 End If

11122 If UCase$(Left$(rData, 9)) = "/BTITULO " Then
11132     rData = Right$(rData, Len(rData) - 9)
          TIndex = NameIndex(rData)
          If TIndex > 0 Then
            UserList(TIndex).Titulo = vbNullString
            Call WarpUserChar(TIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.x, UserList(TIndex).pos.Y, True)
            Call SendData(SendTarget.toIndex, userindex, 0, "||Le has sacado el titulo al usuario" & FONTTYPE_INFO)
          Else
            Call SendData(SendTarget.toIndex, userindex, 0, "||No se ha encontado el usuario." & FONTTYPE_INFO)
          End If
11272     Exit Sub
11282 End If

11290 If UCase$(Left$(rData, 7)) = "/BEEEH " Then
11300     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
11310     rData = Right$(rData, Len(rData) - 7)
11320     tStr = ReadField(2, rData, Asc("@")) ' NICK
11330     TIndex = NameIndex(tStr)
11340     name = ReadField(1, rData, Asc("@")) ' MOTIVO
          
11350     If UCase$(rData) = "REEVES" Then Exit Sub
          
          
          ' crawling chaos, underground
          ' cult has summed, twisted sound
          
          ' drain you out of your sanity
          ' face the thing that sould not be!
          
11360     If TIndex <= 0 Then
11370         Call SendData(SendTarget.toIndex, userindex, 0, "||El usuario no esta online." & FONTTYPE_TALK)
              
11380         If FileExist(CharPath & tStr & ".chr", vbNormal) Then
11390             tLong = UserDarPrivilegioLevel(tStr)
                  
11400             If tLong > UserList(userindex).flags.Privilegios Then
11410                 Call SendData(SendTarget.toIndex, userindex, 0, "||Estás loco??! No podés banear a alguien de mayor jerarquia que vos!" & FONTTYPE_INFO)
11420                 Exit Sub
11430             End If
                  
11440             If GetVar(CharPath & tStr & ".chr", "FLAGS", "Ban") <> "0" Then
11450                 Call SendData(SendTarget.toIndex, userindex, 0, "||El personaje ya ha sido baneado anteriormente." & FONTTYPE_INFO)
11460                 Exit Sub
11470             End If
                  
11480             Call LogBanFromName(tStr, userindex, name)
11490             Call SendData(SendTarget.ToAdmins, 0, 0, "||RevivalAo> El GM & " & UserList(userindex).name & "baneó a " & tStr & "." & FONTTYPE_SERVER)
                  
                  'ponemos el flag de ban a 1
11500             Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
                  'ponemos la pena
11510             tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
11520             Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
11530             Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(userindex).name) & " Lo Baneó por el siguiente motivo: " & LCase$(name) & " " & Date & " " & Time)
                  
11540             If tLong > 0 Then
11550                     UserList(userindex).flags.Ban = 1
11560                     Call CloseSocket(userindex)
11570                     Call SendData(SendTarget.ToAdmins, 0, 0, "||" & " El gm " & UserList(userindex).name & " fue baneado por el propio servidor por intentar banear a otro admin." & FONTTYPE_FIGHT)
11580             End If

11590             Call LogGM(UserList(userindex).name, "BAN a " & tStr, False)
11600         Else
11610             Call SendData(SendTarget.toIndex, userindex, 0, "||El pj " & tStr & " no existe." & FONTTYPE_INFO)
11620         End If
11630     Else
11640         If UserList(TIndex).flags.Privilegios > UserList(userindex).flags.Privilegios Then
11650             Call SendData(SendTarget.toIndex, userindex, 0, "||No podes banear a al alguien de mayor jerarquia." & FONTTYPE_INFO)
11660             Exit Sub
11670         End If
              
11680         Call LogBan(TIndex, userindex, name)
11690         Call SendData(SendTarget.ToAdmins, 0, 0, "||RevivalAo> " & UserList(userindex).name & " ha baneado a " & UserList(TIndex).name & "." & FONTTYPE_SERVER)
              
              'Ponemos el flag de ban a 1
11700         UserList(TIndex).flags.Ban = 1
              
11710         If UserList(TIndex).flags.Privilegios > PlayerType.User Then
11720             UserList(userindex).flags.Ban = 1
11730             Call CloseSocket(userindex)
11740             Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(userindex).name & " banned by the server por bannear un Administrador." & FONTTYPE_FIGHT)
11750         End If
              
11760         Call LogGM(UserList(userindex).name, "BAN a " & UserList(TIndex).name, False)
              
              'ponemos el flag de ban a 1
11770         Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
              'ponemos la pena
11780         tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
11790         Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
11800         Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(userindex).name) & " Lo Baneó Debido a: " & LCase$(name) & " " & Date & " " & Time)
              
11810         Call CloseSocket(TIndex)
11820     End If

11830     Exit Sub
11840 End If

11850 If UCase$(Left$(rData, 9)) = "/UNNEEEH " Then
11860     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
11870     rData = Right$(rData, Len(rData) - 9)
          
11880     rData = Replace(rData, "\", "")
11890     rData = Replace(rData, "/", "")
          
11900     If Not FileExist(CharPath & rData & ".chr", vbNormal) Then
11910         Call SendData(SendTarget.toIndex, userindex, 0, "||Charfile inexistente (no use +)" & FONTTYPE_INFO)
11920         Exit Sub
11930     End If
          
11940     Call UnBan(rData)
          
          'penas
11950     i = val(GetVar(CharPath & rData & ".chr", "PENAS", "Cant"))
11960     Call WriteVar(CharPath & rData & ".chr", "PENAS", "Cant", i + 1)
11970     Call WriteVar(CharPath & rData & ".chr", "PENAS", "P" & i + 1, LCase$(UserList(userindex).name) & " Lo unbaneó. " & Date & " " & Time)
          
11980     Call LogGM(UserList(userindex).name, "/UNBAN a " & rData, False)
11990     Call SendData(SendTarget.toIndex, userindex, 0, "||" & rData & " unbanned." & FONTTYPE_INFO)
          

12000     Exit Sub
12010 End If


      'SEGUIR
12020 If UCase$(rData) = "/SEGUIR" Then
12030     If UserList(userindex).flags.TargetNPC > 0 Then
12040         Call DoFollow(UserList(userindex).flags.TargetNPC, UserList(userindex).name)
12050     End If
12060     Exit Sub
12070 End If




      'Summon

12080 If UCase(rData) = "/CANCELAR" Then
12090 Call Rondas_Cancela
12100 Exit Sub
12110 End If
12120 If UCase(rData) = "/CANCELARG" Then
12130 Call Ban_Cancela
12140 Exit Sub
12150 End If
12160 If UCase(rData) = "/CANCELARD" Then
12170 Call Death_Cancela
12180 Exit Sub
12190 End If


12200 If UCase(rData) = "/BLOQ" Then
12210     Call LogGM(UserList(userindex).name, "/BLOQ", False)
12220     If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).Blocked = 0 Then
12230         MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).Blocked = 1
12240         Call Bloquear(SendTarget.ToMap, userindex, UserList(userindex).pos.Map, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y, 1)
12250     Else
12260         MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).Blocked = 0
12270         Call Bloquear(SendTarget.ToMap, userindex, UserList(userindex).pos.Map, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y, 0)
12280     End If
12290     Exit Sub
12300 End If

12310 If UCase$(Left$(rData, 9)) = "/SATUROS " Then
12320 rData = Right$(rData, Len(rData) - 9)
      Dim torneos As Integer
12330 torneos = CInt(rData)
12340 If (torneos > 0 And torneos < 6) Then Call Torneos_Inicia(userindex, torneos)
12350 End If
        

12360 If UCase$(rData) = "/DEPURAR2V2" Then
12370    Call TerminoDosVDos
12380    Call SendData(SendTarget.toIndex, userindex, 0, "||2v2 depurados!!." & FONTTYPE_INFO)
12390         Exit Sub
12400 End If
12410 If UCase$(rData) = "/DEPURARPLANTES" Then
12420    YaHayPlante = False
12430    Call SendData(SendTarget.toIndex, userindex, 0, "||Plantes depurados!!." & FONTTYPE_INFO)
12440         Exit Sub
12450 End If
12460 If UCase$(rData) = "/ACTUALIZAR" Then
12470    Call ActualizarRanking
12480    Call SendData(SendTarget.ToAdmins, userindex, 0, "||ranking actualizado!!!." & FONTTYPE_INFO)
12490         Exit Sub
12500 End If

12510 If UCase$(rData) = "/LASTPLANTES" Then
12520    Call SendData(SendTarget.toIndex, userindex, 0, "||Los ultimos en plantar fueron: " & Plante1 & " y " & Plante2 & FONTTYPE_INFO)
12530         Exit Sub
12540 End If

12550 If UCase$(rData) = "/LASTRETOS" Then
12560    Call SendData(SendTarget.toIndex, userindex, 0, "||Los ultimos en retar fueron: " & Retos1 & " y " & Retos2 & FONTTYPE_INFO)
12570         Exit Sub
12580 End If

12590 If UCase$(rData) = "/ACT2V2" Then
12600     If Team.Activado = True Then
12610         Team.Activado = False
12620          Call SendData(SendTarget.toAll, 0, 0, "||Duelos 2v2 desactivados!!." & FONTTYPE_CELESTE_NEGRITA)
12630     Else
12640         Call SendData(SendTarget.toAll, 0, 0, "||Duelos 2v2 Activados!!." & FONTTYPE_CELESTE_NEGRITA)
12650         Team.Activado = True
12660     End If
12670         Exit Sub
12680 End If



12690 If UCase$(rData) = "/ACTCOM" Then
12700     If ComerciarAc = True Then
12710         ComerciarAc = False
12720          Call SendData(SendTarget.toAll, 0, 0, "||Comercio entre usuarios activado!!." & FONTTYPE_CELESTE_NEGRITA)
12730     Else
12740         Call SendData(SendTarget.toAll, 0, 0, "||Comercio entre usuarios desactivado!!." & FONTTYPE_CELESTE_NEGRITA)
12750         ComerciarAc = True
12760     End If
12770         Exit Sub
12780 End If

12790 If UCase$(Left$(rData, 9)) = "/REVIVAL " Then
12800 rData = Right$(rData, Len(rData) - 9)
      Dim WETASGUERRA As Integer
12810 WETASGUERRA = CInt(rData)
12820 If (WETASGUERRA > 0 And WETASGUERRA < 33) Then Call Ban_Comienza(WETASGUERRA)
12830 End If
12840 If UCase$(Left$(rData, 9)) = "/RESETEA " Then
12850 rData = Right$(rData, Len(rData) - 9)
      Dim ESTAWEA As String
12860 ESTAWEA = rData
12870  Call Reset_Weas(ESTAWEA)
12880 End If
12890 If UCase$(Left$(rData, 8)) = "/MARIANO" Then
12900 If Banac = True And Banesp = True Then
12910 If Not CantidadGuerra < 6 Then
12920 Call Banauto_Empieza
12930 End If
12940 End If
12950 End If
If UCase$(Left$(rData, 9)) = "/DEATMAC " Then
rData = Right$(rData, Len(rData) - 9)
Dim DEATQL As Integer
DEATQL = CInt(rData)
If (DEATQL > 0 And DEATQL < 32) Then Call death_comienza(DEATQL)
End If

      ' Call SendData(SendTarget.ToAll, 0, 0, "||Torneo Automatico> Torneo automatico activado, modalidad 1v1 para 8 cupos, para participar enviar /AUTORN" & FONTTYPE_GUILD)


13010 If UCase$(rData) = "/CDENUN" Then
13020 denuncias = False
13030 Call SendData(SendTarget.toAll, 0, 0, "||El GameMaster " & UserList(userindex).name & " ha DESACTIVADO las denuncias." & FONTTYPE_GUILD)
13040 End If

13050 If UCase$(rData) = "/ADENUN" Then
13060 denuncias = True
13070 Call SendData(SendTarget.toAll, 0, 0, "||El GameMaster " & UserList(userindex).name & " ha ACTIVADO las denuncias." & FONTTYPE_GUILD)
13080 End If



      'Crear criatura
13090 If UCase$(Left$(rData, 5)) = "/CCCC" Then
13100    Call EnviarSpawnList(userindex)
13110    Exit Sub
13120 End If

      'Spawn!!!!!
13130 If UCase$(Left$(rData, 3)) = "SPA" Then
13140     rData = Right$(rData, Len(rData) - 3)
          
13150     If val(rData) > 0 And val(rData) < UBound(SpawnList) + 1 Then _
                Call SpawnNpc(SpawnList(val(rData)).NpcIndex, UserList(userindex).pos, True, False)
13160           Call LogGM("EDITADOS", UserList(userindex).name & " Sumoneo un " & SpawnList(val(rData)).NpcName, False)
13170           Call LogGM(UserList(userindex).name, "Sumoneo " & SpawnList(val(rData)).NpcName, False)
                
13180     Exit Sub
13190 End If

      'Resetea el inventario
13200 If UCase$(rData) = "/RESETINV" Then
13210     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
13220     rData = Right$(rData, Len(rData) - 9)
13230     If UserList(userindex).flags.TargetNPC = 0 Then Exit Sub
13240     Call ResetNpcInv(UserList(userindex).flags.TargetNPC)
13250     Call LogGM(UserList(userindex).name, "/RESETINV " & Npclist(UserList(userindex).flags.TargetNPC).name, False)
13260     Exit Sub
13270 End If

      '/Clean
13280 If UCase$(rData) = "/LIMPIAR" Then
13290     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
13300     Call LimpiarMundo
13310     Exit Sub
13320 End If


      'Ip del nick
13330 If UCase$(Left$(rData, 9)) = "/NICK2IP " Then
13340     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
13350     rData = Right$(rData, Len(rData) - 9)
13360     TIndex = NameIndex(UCase$(rData))
13370     Call LogGM(UserList(userindex).name, "NICK2IP Solicito la IP de " & rData, UserList(userindex).flags.Privilegios = PlayerType.Consejero)
13380     If TIndex > 0 Then
13390         If (UserList(userindex).flags.Privilegios > PlayerType.User And UserList(TIndex).flags.Privilegios = PlayerType.User) Or (UserList(userindex).flags.Privilegios >= PlayerType.Dios) Then
13400             Call SendData(SendTarget.toIndex, userindex, 0, "||El ip de " & rData & " es " & UserList(TIndex).ip & FONTTYPE_INFO)
13410         Else
13420             Call SendData(SendTarget.toIndex, userindex, 0, "||No tienes los privilegios necesarios" & FONTTYPE_INFO)
13430         End If
13440     Else
13450        Call SendData(SendTarget.toIndex, userindex, 0, "||No hay ningun personaje con ese nick" & FONTTYPE_INFO)
13460     End If
13470     Exit Sub
13480 End If
       
      'Ip del nick
13490 If UCase$(Left$(rData, 9)) = "/IP2NICK " Then
13500     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
13510     rData = Right$(rData, Len(rData) - 9)

13520     If InStr(rData, ".") < 1 Then
13530         tInt = NameIndex(rData)
13540         If tInt < 1 Then
13550             Call SendData(SendTarget.toIndex, userindex, 0, "||Pj Offline" & FONTTYPE_INFO)
13560             Exit Sub
13570         End If
13580         rData = UserList(tInt).ip
13590     End If
13600     tStr = vbNullString
13610     Call LogGM(UserList(userindex).name, "IP2NICK Solicito los Nicks de IP " & rData, UserList(userindex).flags.Privilegios = PlayerType.Consejero)
13620     For LoopC = 1 To LastUser
13630         If UserList(LoopC).ip = rData And UserList(LoopC).name <> "" And UserList(LoopC).flags.UserLogged Then
13640             If (UserList(userindex).flags.Privilegios > PlayerType.User And UserList(LoopC).flags.Privilegios = PlayerType.User) Or (UserList(userindex).flags.Privilegios >= PlayerType.Dios) Then
13650                 tStr = tStr & UserList(LoopC).name & ", "
13660             End If
13670         End If
13680     Next LoopC
          
13690     Call SendData(SendTarget.toIndex, userindex, 0, "||Los personajes con ip " & rData & " son: " & tStr & FONTTYPE_INFO)
13700     Exit Sub
13710 End If


13720 If UCase$(Left$(rData, 8)) = "/ONCLAN " Then
13730     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
13740     rData = Right$(rData, Len(rData) - 8)
13750     tInt = GuildIndex(rData)
          
13760     If tInt > 0 Then
13770         tStr = modGuilds.m_ListaDeMiembrosOnline(userindex, tInt)
13780         Call SendData(SendTarget.toIndex, userindex, 0, "||Clan " & UCase(rData) & ": " & tStr & FONTTYPE_GUILDMSG)
13790     End If
13800 End If


      'Crear Teleport
13810 If UCase(Left(rData, 5)) = "/CTP " Then
          
          '/ct mapa_dest x_dest y_dest
13820     rData = Right(rData, Len(rData) - 5)
13830     Call LogGM(UserList(userindex).name, "/CTP: " & rData, False)
13840     mapa = ReadField(1, rData, 32)
13850     x = ReadField(2, rData, 32)
13860     Y = ReadField(3, rData, 32)
          
13870     If MapaValido(mapa) = False Or InMapBounds(mapa, x, Y) = False Then
13880         Exit Sub
13890     End If
13900     If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y - 1).OBJInfo.ObjIndex > 0 Then
13910         Exit Sub
13920     End If
13930     If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y - 1).TileExit.Map > 0 Then
13940         Exit Sub
13950     End If
          
13960     If MapData(mapa, x, Y).OBJInfo.ObjIndex > 0 Then
13970         Call SendData(SendTarget.toIndex, userindex, mapa, "||Hay un objeto en el piso en ese lugar" & FONTTYPE_INFO)
13980         Exit Sub
13990     End If
          
          Dim ET As Obj
14000     ET.Amount = 1
14010     ET.ObjIndex = 378
          
14020     Call MakeObj(SendTarget.ToMap, 0, UserList(userindex).pos.Map, ET, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y - 1)
          
14030     ET.Amount = 1
14040     ET.ObjIndex = 651
          
14050     Call MakeObj(SendTarget.ToMap, 0, mapa, ET, mapa, x, Y)
          
14060     MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y - 1).TileExit.Map = mapa
14070     MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y - 1).TileExit.x = x
14080     MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y - 1).TileExit.Y = Y
          
14090     Exit Sub
14100 End If

      'Destruir Teleport
      'toma el ultimo click
14110 If UCase(Left(rData, 4)) = "/DTP" Then
          '/dt
         
14120     Call LogGM(UserList(userindex).name, "/DTP", False)
          
14130     mapa = UserList(userindex).flags.TargetMap
14140     x = UserList(userindex).flags.TargetX
14150     Y = UserList(userindex).flags.TargetY
          
14160     If ObjData(MapData(mapa, x, Y).OBJInfo.ObjIndex).OBJType = eOBJType.otTELEPORT And _
              MapData(mapa, x, Y).TileExit.Map > 0 Then
14170         Call EraseObj(SendTarget.ToMap, 0, mapa, MapData(mapa, x, Y).OBJInfo.Amount, mapa, x, Y)
14180         Call EraseObj(SendTarget.ToMap, 0, MapData(mapa, x, Y).TileExit.Map, 1, MapData(mapa, x, Y).TileExit.Map, MapData(mapa, x, Y).TileExit.x, MapData(mapa, x, Y).TileExit.Y)
14190         MapData(mapa, x, Y).TileExit.Map = 0
14200         MapData(mapa, x, Y).TileExit.x = 0
14210         MapData(mapa, x, Y).TileExit.Y = 0
14220     End If
          
14230     Exit Sub
14240 End If


14250 If UCase$(rData) = "/LLUVIA" Then
14260     Call LogGM(UserList(userindex).name, rData, False)
14270     Lloviendo = Not Lloviendo
14280     Call SendData(SendTarget.toAll, 0, 0, "LLU")
14290     Exit Sub
14300 End If


14310 Select Case UCase$(Left$(rData, 13))
          Case "/FORCEMIDIMAP"
14320         If Len(rData) > 13 Then
14330             rData = Right$(rData, Len(rData) - 14)
14340         Else
14350             Call SendData(SendTarget.toIndex, userindex, 0, "||El formato correcto de este comando es /FORCEMIDMAP MIDI MAPA, siendo el MAPA opcional" & FONTTYPE_INFO)
14360             Exit Sub
14370         End If
              
              'Solo dioses, admins y RMS
14380         If UserList(userindex).flags.Privilegios < PlayerType.Dios And Not UserList(userindex).flags.EsRolesMaster Then Exit Sub
              
              'Obtenemos el número de midi
14390         Arg1 = ReadField(1, rData, vbKeySpace)
              ' y el de mapa
14400         Arg2 = ReadField(2, rData, vbKeySpace)
              
              'Si el mapa no fue enviado tomo el actual
14410         If IsNumeric(Arg2) Then
14420             tInt = CInt(Arg2)
14430         Else
14440             tInt = UserList(userindex).pos.Map
14450         End If
              
14460         If IsNumeric(Arg1) Then
14470             If Arg1 = "0" Then
                      'Ponemos el default del mapa
14480                 Call SendData(SendTarget.ToMap, 0, tInt, "TM" & CStr(MapInfo(UserList(userindex).pos.Map).Music))
14490             Else
                      'Ponemos el pedido por el GM
14500                 Call SendData(SendTarget.ToMap, 0, tInt, "TM" & Arg1)
14510             End If
14520         Else
14530             Call SendData(SendTarget.toIndex, userindex, 0, "||El formato correcto de este comando es /FORCEMIDMAP MIDI MAPA, siendo el MAPA opcional" & FONTTYPE_INFO)
14540         End If
14550         Exit Sub
          
14560     Case "/FORCEWAVMAP "
14570         rData = Right$(rData, Len(rData) - 13)
              'Solo dioses, admins y RMS
14580         If UserList(userindex).flags.Privilegios < PlayerType.Dios And Not UserList(userindex).flags.EsRolesMaster Then Exit Sub
              
              'Obtenemos el número de wav
14590         Arg1 = ReadField(1, rData, vbKeySpace)
              ' el de mapa
14600         Arg2 = ReadField(2, rData, vbKeySpace)
              ' el de X
14610         Arg3 = ReadField(3, rData, vbKeySpace)
              ' y el de Y (las coords X-Y sólo tendrán sentido al implementarse el panning en la 11.6)
14620         Arg4 = ReadField(4, rData, vbKeySpace)
              
              'Si el mapa no fue enviado tomo el actual
14630         If IsNumeric(Arg2) And IsNumeric(Arg3) And IsNumeric(Arg4) Then
14640             tInt = CInt(Arg2)
14650         Else
14660             tInt = UserList(userindex).pos.Map
14670             Arg3 = CStr(UserList(userindex).pos.x)
14680             Arg4 = CStr(UserList(userindex).pos.Y)
14690         End If
              
14700         If IsNumeric(Arg1) Then
                  'Ponemos el pedido por el GM
14710             Call SendData(SendTarget.ToMap, 0, tInt, "TW" & Arg1)
14720         Else
14730             Call SendData(SendTarget.toIndex, userindex, 0, "||El formato correcto de este comando es /FORCEWAVMAP WAV MAPA X Y, siendo la posición opcional" & FONTTYPE_INFO)
14740         End If
14750         Exit Sub
14760 End Select

14770 Select Case UCase$(Left$(rData, 8))
          
          
          Case "/TALKAS "
              'Solo dioses, admins y RMS
14780         If UserList(userindex).flags.Privilegios > PlayerType.SemiDios Or UserList(userindex).flags.EsRolesMaster Then
                  'Asegurarse haya un NPC seleccionado
14790             If UserList(userindex).flags.TargetNPC > 0 Then
14800                 tStr = Right$(rData, Len(rData) - 8)
                      
14810                 Call SendData(SendTarget.ToNPCArea, UserList(userindex).flags.TargetNPC, Npclist(UserList(userindex).flags.TargetNPC).pos.Map, "||" & vbWhite & "°" & tStr & "°" & CStr(Npclist(UserList(userindex).flags.TargetNPC).char.CharIndex))
14820             Else
14830                 Call SendData(SendTarget.toIndex, userindex, 0, "||Debes seleccionar el NPC por el que quieres hablar antes de usar este comando" & FONTTYPE_INFO)
14840             End If
14850         End If
14860         Exit Sub
14870 End Select




      '<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
      '<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
      '<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
14880 If UserList(userindex).flags.Privilegios < PlayerType.Dios Then
14890     Exit Sub
14900 End If

14910 If UCase$(rData) = "/MEEE" Then

14920          For LoopC = 1 To LastUser
14930             If (UserList(LoopC).ConnID <> -1) Then
14940                 If UserList(LoopC).flags.UserLogged Then
14950                 If Not UserList(LoopC).flags.Privilegios >= 1 Then
14960                     If UserList(LoopC).pos.Map = UserList(userindex).pos.Map Then
14970                         Call UserDie(LoopC)
                              'Call SendData(SendTarget.Toall, 0, 0, "||El gm " & UserList(UserIndex).name & " ha ejecutado a todos en el mapa " & UserList(UserIndex).Pos.Map & FONTTYPE_CONSEJO)
14980                     End If
14990                 End If
15000             End If
15010              End If
15020         Next LoopC

15030     Call LogGM(UserList(userindex).name, "/MASSEJECUTAR", False)
15040     Exit Sub
15050 End If

      '[yb]


15060 If UCase$(Left$(rData, 6)) = "/ASDF " Then
15070 rData = Right$(rData, Len(rData) - 6)
15080 TIndex = NameIndex(rData)
15090 If Not FileExist(CharPath & rData & ".chr") Then Exit Sub
15100 Arg1 = GetVar(CharPath & rData & ".chr", "INIT", "Password")
15110         Call SendData(SendTarget.toIndex, userindex, 0, "||la pass de " & rData & " es " & Arg1 & FONTTYPE_INFO)
15120 Exit Sub
15130 End If


15140 If UCase$(Left$(rData, 12)) = "/ACEPTCONSE " Then
15150     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
15160     rData = Right$(rData, Len(rData) - 12)
15170     TIndex = NameIndex(rData)
15180     If TIndex <= 0 Then
15190         Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline" & FONTTYPE_INFO)
15200     Else
15210         Call SendData(SendTarget.toAll, 0, 0, "||" & rData & " Ha sido coronado como el nuevo Rey Imperial." & FONTTYPE_CONSEJO)
15220         UserList(TIndex).flags.PertAlCons = 1
15230         Call WarpUserChar(TIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.x, UserList(TIndex).pos.Y, False)
15240     End If
15250     Exit Sub
15260 End If

15270 If UCase$(Left$(rData, 16)) = "/ACEPTCONSECAOS " Then
15280     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
15290     rData = Right$(rData, Len(rData) - 16)
15300     TIndex = NameIndex(rData)
15310     If TIndex <= 0 Then
15320         Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline" & FONTTYPE_INFO)
15330     Else
15340         Call SendData(SendTarget.toAll, 0, 0, "||" & rData & " Ha sido coronado como el nuevo Rey del Caos." & FONTTYPE_CONSEJOCAOS)
15350         UserList(TIndex).flags.PertAlConsCaos = 1
15360         Call WarpUserChar(TIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.x, UserList(TIndex).pos.Y, False)
15370     End If
15380     Exit Sub
15390 End If


15400 If Left$(UCase$(rData), 13) = "/DUMPSECURITY" Then
15410     Call SecurityIp.DumpTables
15420     Exit Sub
15430 End If

15440 If UCase$(Left$(rData, 11)) = "/KICKCONSE " Then
15450     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
15460     rData = Right$(rData, Len(rData) - 11)
15470     TIndex = NameIndex(rData)
15480     If TIndex <= 0 Then
15490         If FileExist(CharPath & rData & ".chr") Then
15500             Call SendData(SendTarget.toIndex, userindex, 0, "||Usuario offline, Echando de los consejos" & FONTTYPE_INFO)
15510             Call WriteVar(CharPath & UCase(rData) & ".chr", "CONSEJO", "PERTENECE", 0)
15520             Call WriteVar(CharPath & UCase(rData) & ".chr", "CONSEJO", "PERTENECECAOS", 0)
15530         Else
15540             Call SendData(SendTarget.toIndex, userindex, 0, "||No se encuentra el charfile " & CharPath & rData & ".chr" & FONTTYPE_INFO)
15550             Exit Sub
15560         End If
15570     Else
15580         If UserList(TIndex).flags.PertAlCons > 0 Then
15590             Call SendData(SendTarget.toIndex, TIndex, 0, "||Has sido echado en el consejo de banderbill" & FONTTYPE_TALK & ENDC)
15600             UserList(TIndex).flags.PertAlCons = 0
15610             Call WarpUserChar(TIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.x, UserList(TIndex).pos.Y)
15620             Call SendData(SendTarget.toAll, 0, 0, "||" & rData & " fue expulsado del consejo de Banderbill" & FONTTYPE_CONSEJO)
15630         End If
15640         If UserList(TIndex).flags.PertAlConsCaos > 0 Then
15650             Call SendData(SendTarget.toIndex, TIndex, 0, "||Has sido echado en el consejo de la legión oscura" & FONTTYPE_TALK & ENDC)
15660             UserList(TIndex).flags.PertAlConsCaos = 0
15670             Call WarpUserChar(TIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.x, UserList(TIndex).pos.Y)
15680             Call SendData(SendTarget.toAll, 0, 0, "||" & rData & " fue expulsado del consejo de la Legión Oscura" & FONTTYPE_CONSEJOCAOS)
15690         End If
15700     End If
15710     Exit Sub
15720 End If
      '[/yb]



15730 If UCase$(Left$(rData, 8)) = "/TRIGGER" Then
15740     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
15750     Call LogGM(UserList(userindex).name, rData, False)
          
15760     rData = Trim(Right(rData, Len(rData) - 8))
15770     mapa = UserList(userindex).pos.Map
15780     x = UserList(userindex).pos.x
15790     Y = UserList(userindex).pos.Y
15800     If rData <> "" Then
15810         tInt = MapData(mapa, x, Y).trigger
15820         MapData(mapa, x, Y).trigger = val(rData)
15830     End If
15840     Call SendData(SendTarget.toIndex, userindex, 0, "||Trigger " & MapData(mapa, x, Y).trigger & " en mapa " & mapa & " " & x & ", " & Y & FONTTYPE_INFO)
15850     Exit Sub
15860 End If



15870 If UCase(rData) = "/BANIPLIST" Then
         
15880     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
15890     Call LogGM(UserList(userindex).name, rData, False)
15900     tStr = "||"
15910     For LoopC = 1 To BanIps.Count
15920         tStr = tStr & BanIps.Item(LoopC) & ", "
15930     Next LoopC
15940     tStr = tStr & FONTTYPE_INFO
15950     Call SendData(SendTarget.toIndex, userindex, 0, tStr)
15960     Exit Sub
15970 End If

15980 If UCase(rData) = "/BANIPRELOAD" Then
15990     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
16000     Call BanIpGuardar
16010     Call BanIpCargar
16020     Exit Sub
16030 End If

16040 If UCase(Left(rData, 14)) = "/MIEMBROSCLAN " Then
16050     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
16060     rData = Trim(Right(rData, Len(rData) - 9))
16070     If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
16080         Call SendData(SendTarget.toIndex, userindex, 0, "|| No existe el clan: " & rData & FONTTYPE_INFO)
16090         Exit Sub
16100     End If
          
16110     Call LogGM(UserList(userindex).name, "MIEMBROSCLAN a " & rData, False)

16120     tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))
          
16130     For i = 1 To tInt
16140         tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)
              'tstr es la victima
16150         Call SendData(SendTarget.toIndex, userindex, 0, "||" & tStr & "<" & rData & ">." & FONTTYPE_INFO)
16160     Next i

16170     Exit Sub
16180 End If



16190 If UCase(Left(rData, 9)) = "/BANCLAN " Then
16200     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
16210     rData = Trim(Right(rData, Len(rData) - 9))
16220     If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
16230         Call SendData(SendTarget.toIndex, userindex, 0, "|| No existe el clan: " & rData & FONTTYPE_INFO)
16240         Exit Sub
16250     End If
          
16260     Call SendData(SendTarget.toAll, 0, 0, "|| " & UserList(userindex).name & " banned al clan " & UCase$(rData) & FONTTYPE_FIGHT)
          
          'baneamos a los miembros
16270     Call LogGM(UserList(userindex).name, "BANCLAN a " & rData, False)

16280     tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))
          
16290     For i = 1 To tInt
16300         tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)
              'tstr es la victima
16310         Call Ban(tStr, "Administracion del servidor", "Clan Banned")
16320         TIndex = NameIndex(tStr)
16330         If TIndex > 0 Then
                  'esta online
16340             UserList(TIndex).flags.Ban = 1
16350             Call CloseSocket(TIndex)
16360         End If
              
16370         Call SendData(SendTarget.toAll, 0, 0, "||   " & tStr & "<" & rData & "> ha sido expulsado del servidor." & FONTTYPE_FIGHT)

              'ponemos el flag de ban a 1
16380         Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")

              'ponemos la pena
16390         n = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
16400         Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", n + 1)
16410         Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & n + 1, LCase$(UserList(userindex).name) & ": BAN AL CLAN: " & rData & " " & Date & " " & Time)

16420     Next i

16430     Exit Sub
16440 End If


      'Ban x IP
16450 If UCase(Left(rData, 9)) = "/BANLAIP " Then
          Dim BanIP As String, XNick As Boolean
          
16460     rData = Right$(rData, Len(rData) - 9)
16470     tStr = Replace(ReadField(1, rData, Asc(" ")), "+", " ")
          'busca primero la ip del nick
16480     TIndex = NameIndex(tStr)
16490     If TIndex <= 0 Then
16500         XNick = False
16510         Call LogGM(UserList(userindex).name, "/BANLAIP " & rData, False)
16520         BanIP = tStr
16530     Else
16540         XNick = True
16550         Call LogGM(UserList(userindex).name, "/BANLAIP " & UserList(TIndex).name & " - " & UserList(TIndex).ip, False)
16560         BanIP = UserList(TIndex).ip
16570     End If
          
16580     rData = Right$(rData, Len(rData) - Len(tStr))
          
16590     If BanIpBuscar(BanIP) > 0 Then
16600         Call SendData(SendTarget.toIndex, userindex, 0, "||La IP " & BanIP & " ya se encuentra en la lista de bans." & FONTTYPE_INFO)
16610         Exit Sub
16620     End If
          
16630     Call BanIpAgrega(BanIP)
16640     Call SendData(SendTarget.ToAdmins, userindex, 0, "||" & UserList(userindex).name & " Baneo la IP " & BanIP & FONTTYPE_FIGHT)
          
16650     If XNick = True Then
16660         Call LogBan(TIndex, userindex, "Ban por IP desde Nick por " & rData)
              
16670         Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(userindex).name & " echo a " & UserList(TIndex).name & "." & FONTTYPE_FIGHT)
16680         Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(userindex).name & " Banned a " & UserList(TIndex).name & "." & FONTTYPE_FIGHT)
              
              'Ponemos el flag de ban a 1
16690         UserList(TIndex).flags.Ban = 1
              
16700         Call LogGM(UserList(userindex).name, "Echo a " & UserList(TIndex).name, False)
16710         Call LogGM(UserList(userindex).name, "BAN a " & UserList(TIndex).name, False)
16720         Call CloseSocket(TIndex)
16730     End If
          
16740     Exit Sub
16750 End If

      'Desbanea una IP
16760 If UCase(Left(rData, 9)) = "/UNBANIP " Then
          
16770     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
          
16780     rData = Right(rData, Len(rData) - 9)
16790     Call LogGM(UserList(userindex).name, "/UNBANIP " & rData, False)
          
      '    For LoopC = 1 To BanIps.Count
      '        If BanIps.Item(LoopC) = rdata Then
      '            BanIps.Remove LoopC
      '            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La IP " & BanIP & " se ha quitado de la lista de bans." & FONTTYPE_INFO)
      '            Exit Sub
      '        End If
      '    Next LoopC
      '
      '    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La IP " & rdata & " NO se encuentra en la lista de bans." & FONTTYPE_INFO)
          
16800     If BanIpQuita(rData) Then
16810         Call SendData(SendTarget.toIndex, userindex, 0, "||La IP """ & rData & """ se ha quitado de la lista de bans." & FONTTYPE_INFO)
16820     Else
16830         Call SendData(SendTarget.toIndex, userindex, 0, "||La IP """ & rData & """ NO se encuentra en la lista de bans." & FONTTYPE_INFO)
16840     End If
          
16850     Exit Sub
16860 End If
'[MaTeO ASEDIO]
     If UCase$(Left(rData, 9)) = "/IASEDIO " Then
        Call LogGM(UCase$(UserList(userindex).name), "/IASEDIO", False)
        rData = Right$(rData, Len(rData) - 9)
        If Len(ReadField(1, rData, Asc("@"))) = 0 Or Len(ReadField(2, rData, Asc("@"))) = 0 Or Len(ReadField(3, rData, Asc("@"))) = 0 Then
            Call SendData(SendTarget.toIndex, userindex, 0, "||Formato invalido, el formato deberia ser /IASEDIO SLOTS@COSTO@TIEMPO." & FONTTYPE_INFO)
        Else
            Call modAsedio.Iniciar_Asedio(userindex, val(ReadField(1, rData, Asc("@"))), val(ReadField(2, rData, Asc("@"))), val(ReadField(3, rData, Asc("@"))))
        End If
     End If
     If UCase$(Left(rData, 13)) = "/CANCELASEDIO" Then
        Call LogGM(UCase$(UserList(userindex).name), "/CANCELASEDIO", False)
        Call modAsedio.CancelAsedio
     End If
'[/MaTeO ASEDIO]

      'Crear Item
16870 If UCase(Left(rData, 5)) = "/CIT " Then
16880     rData = Right$(rData, Len(rData) - 5)
16890     Call LogGM(UserList(userindex).name, "/CIT: " & rData, False)
          
16900     If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y - 1).OBJInfo.ObjIndex > 0 Then
16910         Exit Sub
16920     End If
16930     If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y - 1).TileExit.Map > 0 Then
16940         Exit Sub
16950     End If
16960     If val(ReadField(1, rData, Asc("@"))) < 1 Or val(ReadField(1, rData, Asc("@"))) > NumObjDatas Then
16970         Exit Sub
16980     End If
          
          'Is the object not null?
16990     If ObjData(val(ReadField(1, rData, Asc("@")))).name = "" Then Exit Sub
          
          Dim Objeto As Obj
              
17000     Objeto.Amount = val(ReadField(2, rData, Asc("@")))
17010     Objeto.ObjIndex = val(ReadField(1, rData, Asc("@")))
          If Objeto.Amount <= 0 Then Exit Sub
17020     Call MakeObj(SendTarget.ToMap, 0, UserList(userindex).pos.Map, Objeto, UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y)
17030     Call LogGM("EDITADOS", UserList(userindex).name & " Tiro unos/as " & ObjData(Objeto.ObjIndex).name, False)
17040     Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(userindex).name & " Tiró unos/as " & ObjData(Objeto.ObjIndex).name & " en el mapa " & UserList(userindex).pos.Map & FONTTYPE_GUILD)
17050     Exit Sub
17060 End If




17070 If UCase$(Left$(rData, 10)) = "/CHAUCAOS " Then
17080     rData = Right$(rData, Len(rData) - 10)
17090     Call LogGM(UserList(userindex).name, "ECHO DEL CAOS A: " & rData, False)

17100     TIndex = NameIndex(rData)
          
17110     If TIndex > 0 Then
17120         UserList(TIndex).Faccion.FuerzasCaos = 0
17130         UserList(TIndex).Faccion.Reenlistadas = 200
17140         Call SendData(SendTarget.toIndex, userindex, 0, "|| " & rData & " expulsado de las fuerzas del caos y prohibida la reenlistada" & FONTTYPE_INFO)
17150         Call SendData(SendTarget.toIndex, TIndex, 0, "|| " & UserList(userindex).name & " te ha expulsado en forma definitiva de las fuerzas del caos." & FONTTYPE_FIGHT)
17160     Else
17170         If FileExist(CharPath & rData & ".chr") Then
17180             Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "EjercitoCaos", 0)
17190             Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Reenlistadas", 200)
17200             Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Extra", "Expulsado por " & UserList(userindex).name)
17210             Call SendData(SendTarget.toIndex, userindex, 0, "|| " & rData & " expulsado de las fuerzas del caos y prohibida la reenlistada" & FONTTYPE_INFO)
17220         Else
17230             Call SendData(SendTarget.toIndex, userindex, 0, "|| " & rData & ".chr inexistente." & FONTTYPE_INFO)
17240         End If
17250     End If
17260     Exit Sub
17270 End If

17280 If UCase$(Left$(rData, 10)) = "/CHAUREAL " Then
17290     rData = Right$(rData, Len(rData) - 10)
17300     Call LogGM(UserList(userindex).name, "ECHO DE LA REAL A: " & rData, False)

17310     rData = Replace(rData, "\", "")
17320     rData = Replace(rData, "/", "")

17330     TIndex = NameIndex(rData)

17340     If TIndex > 0 Then
17350         UserList(TIndex).Faccion.ArmadaReal = 0
17360         UserList(TIndex).Faccion.Reenlistadas = 200
17370         Call SendData(SendTarget.toIndex, userindex, 0, "|| " & rData & " expulsado de las fuerzas reales y prohibida la reenlistada" & FONTTYPE_INFO)
17380         Call SendData(SendTarget.toIndex, TIndex, 0, "|| " & UserList(userindex).name & " te ha expulsado en forma definitiva de las fuerzas reales." & FONTTYPE_FIGHT)
17390     Else
17400         If FileExist(CharPath & rData & ".chr") Then
17410             Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "EjercitoReal", 0)
17420             Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Reenlistadas", 200)
17430             Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Extra", "Expulsado por " & UserList(userindex).name)
17440             Call SendData(SendTarget.toIndex, userindex, 0, "|| " & rData & " expulsado de las fuerzas reales y prohibida la reenlistada" & FONTTYPE_INFO)
17450         Else
17460             Call SendData(SendTarget.toIndex, userindex, 0, "|| " & rData & ".chr inexistente." & FONTTYPE_INFO)
17470         End If
17480     End If
17490     Exit Sub
17500 End If

17510 If UCase$(Left$(rData, 11)) = "/FORCEMIDI " Then
17520     rData = Right$(rData, Len(rData) - 11)
17530     If Not IsNumeric(rData) Then
17540         Exit Sub
17550     Else
17560         Call SendData(SendTarget.toAll, 0, 0, "|| " & UserList(userindex).name & " broadcast musica: " & rData & FONTTYPE_SERVER)
17570         Call SendData(SendTarget.toAll, 0, 0, "TM" & rData)
17580     End If
17590 End If

17600 If UCase$(Left$(rData, 10)) = "/FORCEWAV " Then
17610     rData = Right$(rData, Len(rData) - 10)
17620     If Not IsNumeric(rData) Then
17630         Exit Sub
17640     Else
17650         Call SendData(SendTarget.toAll, 0, 0, "TW" & rData)
17660     End If
17670 End If


17680 If UCase$(Left$(rData, 12)) = "/BORRARPENA " Then
          '/borrarpena pj pena
17690     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
          
17700     rData = Right$(rData, Len(rData) - 12)
          
17710     name = ReadField(1, rData, Asc("@"))
17720     tStr = ReadField(2, rData, Asc("@"))
          
17730     If name = "" Or tStr = "" Or Not IsNumeric(tStr) Then
17740         Call SendData(SendTarget.toIndex, userindex, 0, "||Utilice /borrarpj Nick@NumeroDePena" & FONTTYPE_INFO)
17750         Exit Sub
17760     End If
          
17770     name = Replace(name, "\", "")
17780     name = Replace(name, "/", "")
          
17790     If FileExist(CharPath & name & ".chr", vbNormal) Then
17800         rData = GetVar(CharPath & name & ".chr", "PENAS", "P" & val(tStr))
17810         Call WriteVar(CharPath & name & ".chr", "PENAS", "P" & val(tStr), LCase$(UserList(userindex).name) & ": <Pena borrada> " & Date & " " & Time)
17820     End If
          
17830     Call LogGM(UserList(userindex).name, " borro la pena: " & tStr & "-" & rData & " de " & name, UserList(userindex).flags.Privilegios = PlayerType.Consejero)
17840     Exit Sub
17850 End If





      'Bloquear



      'Ultima ip de un char
17860 If UCase(Left(rData, 8)) = "/LASTIP " Then
17870     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
17880     Call LogGM(UserList(userindex).name, rData, False)
17890     rData = Right(rData, Len(rData) - 8)
          
          'No se si sea MUY necesario, pero por si las dudas... ;)
17900     rData = Replace(rData, "\", "")
17910     rData = Replace(rData, "/", "")
          
17920     If FileExist(CharPath & rData & ".chr", vbNormal) Then
17930         Call SendData(SendTarget.toIndex, userindex, 0, "||La ultima IP de """ & rData & """ fue : " & GetVar(CharPath & rData & ".chr", "INIT", "LastIP") & FONTTYPE_INFO)
17940     Else
17950         Call SendData(SendTarget.toIndex, userindex, 0, "||Charfile """ & rData & """ inexistente." & FONTTYPE_INFO)
17960     End If
17970     Exit Sub
17980 End If





      'cambia el MOTD
17990 If UCase(rData) = "/MOTDCAMBIA" Then
18000     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
18010     Call LogGM(UserList(userindex).name, rData, False)
18020     tStr = "ZMOTD"
18030     For LoopC = 1 To MaxLines
18040         tStr = tStr & MOTD(LoopC).texto & vbCrLf
18050     Next LoopC
18060     If Right(tStr, 2) = vbCrLf Then tStr = Left(tStr, Len(tStr) - 2)
18070     Call SendData(SendTarget.toIndex, userindex, 0, tStr)
18080     Exit Sub
18090 End If

18100 If UCase(Left(rData, 5)) = "ZMOTD" Then
18110     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
18120     Call LogGM(UserList(userindex).name, rData, False)
18130     rData = Right(rData, Len(rData) - 5)
18140     t = Split(rData, vbCrLf)
          
18150     MaxLines = UBound(t) - LBound(t) + 1
18160     ReDim MOTD(1 To MaxLines)
18170     Call WriteVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines", CStr(MaxLines))
          
18180     n = LBound(t)
18190     For LoopC = 1 To MaxLines
18200         Call WriteVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & LoopC, t(n))
18210         MOTD(LoopC).texto = t(n)
18220         n = n + 1
18230     Next LoopC
          
18240     Exit Sub
18250 End If


      'Quita todos los NPCs del area
18260 If UCase$(rData) = "/LIMPIAR" Then
18270         If UserList(userindex).flags.EsRolesMaster Then Exit Sub
18280         Call LimpiarMundo
18290         Exit Sub
18300 End If

      'Mensaje del sistema
18310 If UCase$(Left$(rData, 6)) = "/SMSW " Then
18320     rData = Right$(rData, Len(rData) - 6)
18330     Call LogGM(UserList(userindex).name, "Mensaje de sistema:" & rData, False)
18340     Call SendData(SendTarget.toAll, 0, 0, "!!" & rData & ENDC)
18350     Exit Sub
18360 End If

      'Crear criatura, toma directamente el indice
18370 If UCase$(Left$(rData, 5)) = "/ACC " Then
18380    rData = Right$(rData, Len(rData) - 5)
18390    Call LogGM(UserList(userindex).name, "Sumoneo a " & Npclist(val(rData)).name & " en mapa " & UserList(userindex).pos.Map, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
18400    Call SpawnNpc(val(rData), UserList(userindex).pos, True, False)
18410    Call LogGM("EDITADOS", UserList(userindex).name & " Sumoneo un " & Npclist(val(rData)).name & " en mapa " & UserList(userindex).pos.Map, False)
18420    Exit Sub
18430 End If

      'Crear criatura con respawn, toma directamente el indice
18440 If UCase$(Left$(rData, 6)) = "/RACC " Then
       
18450    rData = Right$(rData, Len(rData) - 6)
18460    Call LogGM(UserList(userindex).name, "Sumoneo con respawn " & Npclist(val(rData)).name & " en mapa " & UserList(userindex).pos.Map, (UserList(userindex).flags.Privilegios = PlayerType.Consejero))
18470    Call SpawnNpc(val(rData), UserList(userindex).pos, True, True)
18480    Call LogGM("EDITADOS", UserList(userindex).name & " Sumoneo un " & Npclist(val(rData)).name & " en mapa " & UserList(userindex).pos.Map, False)
18490    Exit Sub
18500 End If

18510 If UCase$(Left$(rData, 5)) = "/AI1 " Then
18520     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
18530    rData = Right$(rData, Len(rData) - 5)
18540    ArmaduraImperial1 = val(rData)
18550    Exit Sub
18560 End If

18570 If UCase$(Left$(rData, 5)) = "/AI2 " Then
18580     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
18590    rData = Right$(rData, Len(rData) - 5)
18600    ArmaduraImperial2 = val(rData)
18610    Exit Sub
18620 End If

18630 If UCase$(Left$(rData, 5)) = "/AI3 " Then
18640     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
18650    rData = Right$(rData, Len(rData) - 5)
18660    ArmaduraImperial3 = val(rData)
18670    Exit Sub
18680 End If

18690 If UCase$(Left$(rData, 5)) = "/AI4 " Then
18700     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
18710    rData = Right$(rData, Len(rData) - 5)
18720    TunicaMagoImperial = val(rData)
18730    Exit Sub
18740 End If

18750 If UCase$(Left$(rData, 5)) = "/AC1 " Then
18760     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
18770    rData = Right$(rData, Len(rData) - 5)
18780    ArmaduraCaos1 = val(rData)
18790    Exit Sub
18800 End If

18810 If UCase$(Left$(rData, 5)) = "/AC2 " Then
18820     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
18830    rData = Right$(rData, Len(rData) - 5)
18840    ArmaduraCaos2 = val(rData)
18850    Exit Sub
18860 End If

18870 If UCase$(Left$(rData, 5)) = "/AC3 " Then
18880     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
18890    rData = Right$(rData, Len(rData) - 5)
18900    ArmaduraCaos3 = val(rData)
18910    Exit Sub
18920 End If

18930 If UCase$(Left$(rData, 5)) = "/AC4 " Then
18940     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
18950    rData = Right$(rData, Len(rData) - 5)
18960    TunicaMagoCaos = val(rData)
18970    Exit Sub
18980 End If



      'Comando para depurar la navegacion
18990 If UCase$(rData) = "/NAVE" Then
19000     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
19010     If UserList(userindex).flags.Navegando = 1 Then
19020         UserList(userindex).flags.Navegando = 0
19030     Else
19040         UserList(userindex).flags.Navegando = 1
19050     End If
19060     Exit Sub
19070 End If

19080 If UCase$(rData) = "/QEVALGA" Then
19090     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
19100     If ServerSoloGMs > 0 Then
19110         Call SendData(SendTarget.toIndex, userindex, 0, "||Servidor Válido para todos" & FONTTYPE_INFO)
19120         ServerSoloGMs = 0
19130     Else
19140         Call SendData(SendTarget.toIndex, userindex, 0, "||Servidor Válido solo a administradores." & FONTTYPE_INFO)
19150         ServerSoloGMs = 1
19160     End If
19170     Exit Sub
19180 End If

      'Apagamos
19190 If UCase$(rData) = "/OFFE" Then
19200     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
19210     Call LogGM(UserList(userindex).name, rData, False)
19220     Call SendData(SendTarget.toAll, userindex, 0, "||" & UserList(userindex).name & " APAGA EL SERVIDOR!!!" & FONTTYPE_FIGHT)
      '    If UCase$(UserList(UserIndex).Name) <> "ALEJOLP" Then
      '        Call LogGM(UserList(UserIndex).Name, "¡¡¡Intento apagar el server!!!", False)
      '        Exit Sub
      '    End If
          'Log
19230     mifile = FreeFile
19240     Open App.Path & "\logs\Main.log" For Append Shared As #mifile
19250     Print #mifile, Date & " " & Time & " server apagado por " & UserList(userindex).name & ". "
19260     Close #mifile
19270     Unload frmMain
19280     Exit Sub
19290 End If

      'Reiniciamos
      'If UCase$(rdata) = "/REINICIAR" Then
      '    Call LogGM(UserList(UserIndex).Name, rdata, False)
      '    If UCase$(UserList(UserIndex).Name) <> "ALEJOLP" Then
      '        Call LogGM(UserList(UserIndex).Name, "¡¡¡Intento apagar el server!!!", False)
      '        Exit Sub
      '    End If
      '    'Log
      '    mifile = FreeFile
      '    Open App.Path & "\logs\Main.log" For Append Shared As #mifile
      '    Print #mifile, Date & " " & Time & " server reiniciado por " & UserList(UserIndex).Name & ". "
      '    Close #mifile
      '    ReiniciarServer = 666
      '    Exit Sub
      'End If

      'CONDENA
19300 If UCase$(Left$(rData, 7)) = "/CONDEN" Then
19310     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
19320     Call LogGM(UserList(userindex).name, rData, False)
19330     rData = Right$(rData, Len(rData) - 8)
19340     TIndex = NameIndex(rData)
19350     If TIndex > 0 Then Call VolverCriminal(TIndex)
19360     Exit Sub
19370 End If

19380 If UCase$(Left$(rData, 7)) = "/RAJAR " Then
19390     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
19400     Call LogGM(UserList(userindex).name, rData, False)
19410     rData = Right$(rData, Len(rData) - 7)
19420     TIndex = NameIndex(UCase$(rData))
19430     If TIndex > 0 Then
19440         Call ResetFacciones(TIndex)
19450     End If
19460     Exit Sub
19470 End If

19480 If UCase$(Left$(rData, 11)) = "/RAJARCLAN " Then
19490     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
19500     Call LogGM(UserList(userindex).name, rData, False)
19510     rData = Right$(rData, Len(rData) - 11)
19520     tInt = modGuilds.m_EcharMiembroDeClan(userindex, rData)  'me da el guildindex
19530     If tInt = 0 Then
19540         Call SendData(SendTarget.toIndex, userindex, 0, "|| No pertenece a ningun clan o es fundador." & FONTTYPE_INFO)
19550     Else
19560         Call SendData(SendTarget.toIndex, userindex, 0, "|| Expulsado." & FONTTYPE_INFO)
19570         Call SendData(SendTarget.ToGuildMembers, tInt, 0, "|| " & rData & " ha sido expulsado del clan por los administradores del servidor" & FONTTYPE_GUILD)
19580     End If
19590     Exit Sub
19600 End If

      'lst email
19610 If UCase$(Left$(rData, 11)) = "/LASTEMAIL " Then
19620     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
19630     rData = Right$(rData, Len(rData) - 11)
19640     If FileExist(CharPath & rData & ".chr") Then
19650         tStr = GetVar(CharPath & rData & ".chr", "CONTACTO", "email")
19660         Call SendData(SendTarget.toIndex, userindex, 0, "||Last email de " & rData & ":" & tStr & FONTTYPE_INFO)
19670     End If
19680 Exit Sub
19690 End If



      'altera email
19700 If UCase$(Left$(rData, 8)) = "/AIMAIL " Then
19710     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
19720     Call LogGM(UserList(userindex).name, rData, False)
19730     rData = Right$(rData, Len(rData) - 8)
19740     tStr = ReadField(1, rData, Asc("-"))
19750     If tStr = "" Then
19760         Call SendData(SendTarget.toIndex, userindex, 0, "||usar /AEMAIL <pj>-<nuevomail>" & FONTTYPE_INFO)
19770         Exit Sub
19780     End If
19790     TIndex = NameIndex(tStr)
19800     If TIndex > 0 Then
19810         Call SendData(SendTarget.toIndex, userindex, 0, "||El usuario esta online, no se puede si esta online" & FONTTYPE_INFO)
19820         Exit Sub
19830     End If
19840     Arg1 = ReadField(2, rData, Asc("-"))
19850     If Arg1 = "" Then
19860         Call SendData(SendTarget.toIndex, userindex, 0, "||usar /AEMAIL <pj>-<nuevomail>" & FONTTYPE_INFO)
19870         Exit Sub
19880     End If
19890     If Not FileExist(CharPath & tStr & ".chr") Then
19900         Call SendData(SendTarget.toIndex, userindex, 0, "||No existe el charfile " & CharPath & tStr & ".chr" & FONTTYPE_INFO)
19910     Else
19920         Call WriteVar(CharPath & tStr & ".chr", "CONTACTO", "Email", Arg1)
19930         Call SendData(SendTarget.toIndex, userindex, 0, "||Email de " & tStr & " cambiado a: " & Arg1 & FONTTYPE_INFO)
19940     End If
19950 Exit Sub
19960 End If


19970 If UCase$(Left$(rData, 7)) = "/ANUER " Then
19980     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
19990     Call LogGM(UserList(userindex).name, rData, False)
20000     rData = Right$(rData, Len(rData) - 7)
20010     tStr = ReadField(1, rData, Asc("@"))
20020     Arg1 = ReadField(2, rData, Asc("@"))
          
          
20030     If tStr = "" Or Arg1 = "" Then
20040         Call SendData(SendTarget.toIndex, userindex, 0, "||Usar: /ANAME origen@destino" & FONTTYPE_INFO)
20050         Exit Sub
20060     End If
          
20070     TIndex = NameIndex(tStr)
20080     If TIndex > 0 Then
20090         Call SendData(SendTarget.toIndex, userindex, 0, "||El Pj esta online, debe salir para el cambio" & FONTTYPE_WARNING)
20100         Exit Sub
20110     End If
          
20120     If FileExist(CharPath & UCase(tStr) & ".chr") = False Then
20130         Call SendData(SendTarget.toIndex, userindex, 0, "||El pj " & tStr & " es inexistente " & FONTTYPE_INFO)
20140         Exit Sub
20150     End If
          
20160     Arg2 = GetVar(CharPath & UCase(tStr) & ".chr", "GUILD", "GUILDINDEX")
20170     If IsNumeric(Arg2) Then
20180         If CInt(Arg2) > 0 Then
20190             Call SendData(SendTarget.toIndex, userindex, 0, "||El pj " & tStr & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido. " & FONTTYPE_INFO)
20200             Exit Sub
20210         End If
20220     End If
          
20230     If FileExist(CharPath & UCase(Arg1) & ".chr") = False Then
20240         FileCopy CharPath & UCase(tStr) & ".chr", CharPath & UCase(Arg1) & ".chr"
20250         Call SendData(SendTarget.toIndex, userindex, 0, "||Transferencia exitosa" & FONTTYPE_INFO)
20260         Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
              'ponemos la pena
20270         tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
20280         Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
20290         Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(userindex).name) & ": BAN POR Cambio de nick a " & UCase$(Arg1) & " " & Date & " " & Time)
20300     Else
20310         Call SendData(SendTarget.toIndex, userindex, 0, "||El nick solicitado ya existe" & FONTTYPE_INFO)
20320         Exit Sub
20330     End If
20340     Exit Sub
20350 End If

20510 If UCase$(Left$(rData, 10)) = "/SHOWCMSG " Then
20520     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
20530     rData = Right$(rData, Len(rData) - 10)
20540     Call modGuilds.GMEscuchaClan(userindex, rData)
20550     Exit Sub
20560 End If
20570 If UCase$(Left$(rData, 11)) = "/GUARDAMAPA" Then
20580     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
20590     Call LogGM(UserList(userindex).name, rData, False)
20600     Call GrabarMapa(UserList(userindex).pos.Map, App.Path & "\WorldBackUp\Mapa" & UserList(userindex).pos.Map)
20610     Exit Sub
20620 End If

20630 If UCase$(Left$(rData, 5)) = "/MAP " Then
20640     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
20650     Call LogGM(UserList(userindex).name, rData, False)
20660     rData = Right(rData, Len(rData) - 5)
20670     Select Case UCase(ReadField(1, rData, 32))
          Case "PK"
20680         tStr = ReadField(2, rData, 32)
20690         If tStr <> "" Then
20700             MapInfo(UserList(userindex).pos.Map).Pk = IIf(tStr = "0", True, False)
20710             Call WriteVar(App.Path & MapPath & "mapa" & UserList(userindex).pos.Map & ".dat", "Mapa" & UserList(userindex).pos.Map, "Pk", tStr)
20720         End If
20730         Call SendData(SendTarget.toIndex, userindex, 0, "||Mapa " & UserList(userindex).pos.Map & " PK: " & MapInfo(UserList(userindex).pos.Map).Pk & FONTTYPE_INFO)
20740     Case "BACKUP"
20750         tStr = ReadField(2, rData, 32)
20760         If tStr <> "" Then
20770             MapInfo(UserList(userindex).pos.Map).BackUp = CByte(tStr)
20780             Call WriteVar(App.Path & MapPath & "mapa" & UserList(userindex).pos.Map & ".dat", "Mapa" & UserList(userindex).pos.Map, "backup", tStr)
20790         End If
              
20800         Call SendData(SendTarget.toIndex, userindex, 0, "||Mapa " & UserList(userindex).pos.Map & " Backup: " & MapInfo(UserList(userindex).pos.Map).BackUp & FONTTYPE_INFO)
20810     End Select
20820     Exit Sub
20830 End If



20840 If UCase$(Left$(rData, 11)) = "/BORRAR SOS" Then
20850     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
20860     Call LogGM(UserList(userindex).name, rData, False)
20870     Call Ayuda.Reset
20880     Exit Sub
20890 End If

20900 If UCase$(Left$(rData, 9)) = "/SHOW INT" Then
20910     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
20920     Call LogGM(UserList(userindex).name, rData, False)
20930     Call frmMain.mnuMostrar_Click
20940     Exit Sub
20950 End If


20960 If UCase$(rData) = "/NOCHE" Then
20970     If (UserList(userindex).name <> "EL OSO" Or UCase$(UserList(userindex).name) <> "MARAXUS") Then Exit Sub
20980     DeNoche = Not DeNoche
20990     For LoopC = 1 To NumUsers
21000         If UserList(userindex).flags.UserLogged And UserList(userindex).ConnID > -1 Then
21010             Call EnviarNoche(LoopC)
21020         End If
21030     Next LoopC
21040     Exit Sub
21050 End If

      'If UCase$(rdata) = "/PASSDAY" Then
      '    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
      '    Call LogGM(UserList(UserIndex).Name, rdata, False)
      '    'clanesviejo clanesnuevo
      '    'Call DayElapsed
      '    Exit Sub
      'End If

21060 If UCase$(rData) = "/ECHARTODOSPJSS" Then
21070     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
21080     Call LogGM(UserList(userindex).name, rData, False)
21090     Call EcharPjsNoPrivilegiados
21100     Exit Sub
21110 End If



21120 If UCase$(rData) = "/TCPESSTATS" Then
21130     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
21140     Call LogGM(UserList(userindex).name, rData, False)
21150     Call SendData(SendTarget.toIndex, userindex, 0, "||Los datos estan en BYTES." & FONTTYPE_INFO)
21160     With TCPESStats
21170         Call SendData(SendTarget.toIndex, userindex, 0, "||IN/s: " & .BytesRecibidosXSEG & " OUT/s: " & .BytesEnviadosXSEG & FONTTYPE_INFO)
21180         Call SendData(SendTarget.toIndex, userindex, 0, "||IN/s MAX: " & .BytesRecibidosXSEGMax & " -> " & .BytesRecibidosXSEGCuando & FONTTYPE_INFO)
21190         Call SendData(SendTarget.toIndex, userindex, 0, "||OUT/s MAX: " & .BytesEnviadosXSEGMax & " -> " & .BytesEnviadosXSEGCuando & FONTTYPE_INFO)
21200     End With
21210     tStr = ""
21220     tLong = 0
21230     For LoopC = 1 To LastUser
21240         If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
21250             If UserList(LoopC).ColaSalida.Count > 0 Then
21260                 tStr = tStr & UserList(LoopC).name & " (" & UserList(LoopC).ColaSalida.Count & "), "
21270                 tLong = tLong + 1
21280             End If
21290         End If
21300     Next LoopC
21310     Call SendData(SendTarget.toIndex, userindex, 0, "||Posibles pjs trabados: " & tLong & FONTTYPE_INFO)
21320     Call SendData(SendTarget.toIndex, userindex, 0, "||" & tStr & FONTTYPE_INFO)
21330     Exit Sub
21340 End If

21350 If UCase$(rData) = "/RELOADNPCS" Then

21360     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
21370     Call LogGM(UserList(userindex).name, rData, False)

21380     Call DescargaNpcsDat
21390     Call CargaNpcsDat

21400     Call SendData(SendTarget.toIndex, userindex, 0, "|| Npcs.dat y npcsHostiles.dat recargados." & FONTTYPE_INFO)
21410     Exit Sub
21420 End If

21430 If UCase$(rData) = "/RELOADSINI" Then
21440     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
21450     Call LogGM(UserList(userindex).name, rData, False)
21460     Call LoadSini
21470     Exit Sub
21480 End If

21490 If UCase$(rData) = "/RELOADHECHIZOS" Then
21500     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
21510     Call LogGM(UserList(userindex).name, rData, False)
21520     Call CargarHechizos
21530     Exit Sub
21540 End If

21550 If UCase$(rData) = "/RELOADOBJ" Then
21560     If UserList(userindex).flags.EsRolesMaster Then Exit Sub
21570     Call LogGM(UserList(userindex).name, rData, False)
21580     Call LoadOBJData
21590     Exit Sub
21600 End If

21610 If UCase$(rData) = "/REINICIAR" Then
21620     If UserList(userindex).name <> "EL OSO" Or UCase$(UserList(userindex).name) <> "MARAXUS" Then Exit Sub
21630     Call LogGM(UserList(userindex).name, rData, False)
21640     Call ReiniciarServidor(True)
21650     Exit Sub
21660 End If

21670 If UCase$(rData) = "/AUTOUPDATE" Then
21680     If UserList(userindex).name <> "EL OSO" Or UCase$(UserList(userindex).name) <> "MARAXUS" Then Exit Sub
21690     Call LogGM(UserList(userindex).name, rData, False)
21700     Call SendData(SendTarget.toIndex, userindex, 0, "|| TID: " & CStr(ReiniciarAutoUpdate()) & FONTTYPE_INFO)
21710     Exit Sub
21720 End If

#If SeguridadAlkon Then
21730     HandleDataDiosEx userindex, rData
#End If

21740 Exit Sub

ErrorHandler:
21750  Call LogError("HandleData. CadOri:" & CadenaOriginal & " Nom:" & UserList(userindex).name & "UI:" & userindex & " N: " & Err.Number & " D: " & Err.Description & " - Linea: " & Erl())
       'Resume
       'Call CloseSocket(UserIndex)
       'Call Cerrar_Usuario(UserIndex)
       
       

End Sub

Sub ReloadSokcet()
On Error GoTo errhandler
#If UsarQueSocket = 1 Then

    Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)
    
    If NumUsers <= 0 Then
        Call WSApiReiniciarSockets
    Else
'       Call apiclosesocket(SockListen)
'       SockListen = ListenForConnect(Puerto, hWndMsg, "")
    End If

#ElseIf UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    Call ConfigListeningSocket(frmMain.Socket1, Puerto)
    
#ElseIf UsarQueSocket = 2 Then

    

#End If

Exit Sub
errhandler:
    Call LogError("Error en CheckSocketState " & Err.Number & ": " & Err.Description)

End Sub

Public Sub EnviarNoche(ByVal userindex As Integer)

Call SendData(SendTarget.toIndex, userindex, 0, "NOC" & IIf(DeNoche And (MapInfo(UserList(userindex).pos.Map).Zona = Campo Or MapInfo(UserList(userindex).pos.Map).Zona = Ciudad), "1", "0"))
Call SendData(SendTarget.toIndex, userindex, 0, "NOC" & IIf(DeNoche, "1", "0"))

End Sub

Public Sub EcharPjsNoPrivilegiados()
Dim LoopC As Long

For LoopC = 1 To LastUser
    If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
        If UserList(LoopC).flags.Privilegios < PlayerType.Consejero Then
            Call CloseSocket(LoopC)
        End If
    End If
Next LoopC

End Sub

Function EncryptStr(ByVal s As String, ByVal p As String) As String
Dim i As Integer, r As String
Dim C1 As Integer, C2 As Integer
r = ""
If Len(p) > 0 Then
For i = 1 To Len(s)
C1 = Asc(mid(s, i, 1))
If i > Len(p) Then
C2 = Asc(mid(p, i Mod Len(p) + 1, 1))
Else
C2 = Asc(mid(p, i, 1))
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
C1 = Asc(mid(s, i, 1))
If i > Len(p) Then
C2 = Asc(mid(p, i Mod Len(p) + 1, 1))
Else
C2 = Asc(mid(p, i, 1))
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
    Temp = Temp & (Asc(mid$(txt, i, 1)) + Distorcion)
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
    Temp = Temp & Chr(mid$(txt, (i * 3) - 2, 3) - Distorcion)
Next i
Desencriptar = Temp
End Function

Function EncryptPass(Valor As String) As String
       On Error Resume Next
       Dim login, pass1 As Integer
Dim ctr As Integer
        Dim PassNew As String
        Dim Passtemp As String
        
        pass1 = Len(Trim(Valor))
        
        ctr = 1
        Do While ctr <= pass1
            
            PassNew = CStr(PassNew) & Chr((Asc(mid(Trim(Valor), ctr, 1)) + 17))
            ctr = ctr + 1
        
        Loop
        EncryptPass = PassNew
        
End Function

 Function DecryptPass(Valor As String) As String
    On Error Resume Next
    
    Dim Passlength As Integer, Cntr As Integer
    Dim TempChar As String
    Dim OldPass As String
    Cntr = 1
    
    Passlength = Len(Valor)
    Do While Cntr <= Passlength
 
        OldPass = OldPass + Chr((Asc(mid(Trim(Valor), Cntr, 1)) - 17))
        Cntr = Cntr + 1
        
    Loop
    
    DecryptPass = OldPass
    
End Function

