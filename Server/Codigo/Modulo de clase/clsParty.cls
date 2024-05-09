VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_members(1 To PARTY_MAXMEMBERS) As tPartyMember
'miembros

Private p_expTotal As Long
'Estadistica :D

Private p_Fundador As Integer
'el creador

Private p_CantMiembros As Integer
'cantidad de miembros

'datos en los pjs: | indexParty(indice en p_members), partyLeader(userindex del lider) |

'Constructor de clase
Public Sub Class_Initialize()
    p_expTotal = 0
    p_CantMiembros = 0
End Sub

'Destructor de clase
Public Sub Class_Terminate()
End Sub


Public Function MiExperiencia(ByVal userindex As Integer) As Long
'Me dice cuanta experiencia tengo colectada ya en la party
Dim i As Integer
i = 0

While i <= PARTY_MAXMEMBERS And p_members(i).userindex <> userindex
    i = i + 1
Wend

If i <= PARTY_MAXMEMBERS Then
    MiExperiencia = p_members(i).Experiencia
Else        'esto no deberia pasar :p
    MiExperiencia = -1
End If

End Function

Public Sub ObtenerExito(ByVal ExpGanada As Long, ByVal mapa As Integer, x As Integer, Y As Integer)
'Se produjo un evento que da experiencia en la wp referenciada
Dim i As Integer
Dim UI As Integer
Dim expPerCapita As Double

p_expTotal = p_expTotal + ExpGanada

expPerCapita = CLng(ExpGanada / p_CantMiembros)

For i = 1 To PARTY_MAXMEMBERS
    UI = p_members(i).userindex
    If UI > 0 Then
        If mapa = UserList(UI).pos.Map And UserList(UI).flags.Muerto = 0 Then
            If Distance(UserList(UI).pos.x, UserList(UI).pos.Y, x, Y) <= PARTY_MAXDISTANCIA Then
                p_members(i).Experiencia = p_members(i).Experiencia + expPerCapita
                If p_members(i).Experiencia < 0 Then
                    p_members(i).Experiencia = 0
                End If
                If PARTY_EXPERIENCIAPORGOLPE Then
                    UserList(UI).Stats.Exp = UserList(UI).Stats.Exp + expPerCapita
                    If UserList(UI).Stats.Exp > MAXEXP Then _
                        UserList(UI).Stats.Exp = MAXEXP
                    Call CheckUserLevel(UI)
                    Call EnviarExp(UI)
                End If
            End If
        End If
    End If
Next i

End Sub

Public Sub MandarMensajeAConsola(ByVal texto As String, ByVal Sender As String)
'feo feo, muy feo acceder a senddata desde aca, pero BUEEEEEEEEEEE...
Dim i As Integer

For i = 1 To PARTY_MAXMEMBERS
    If p_members(i).userindex > 0 Then
        Call SendData(SendTarget.toIndex, p_members(i).userindex, 0, "|| [" & Sender & "] " & texto & FONTTYPE_PARTY)
    End If
Next i

End Sub

Public Function EsPartyLeader(ByVal userindex As Integer) As Boolean
    EsPartyLeader = (userindex = p_Fundador)
End Function

Public Function NuevoMiembro(ByVal userindex As Integer) As Boolean
Dim i As Integer
i = 1
While i <= PARTY_MAXMEMBERS And p_members(i).userindex > 0
    i = i + 1
Wend

If i <= PARTY_MAXMEMBERS Then
    p_members(i).Experiencia = 0
    p_members(i).userindex = userindex
    NuevoMiembro = True
    p_CantMiembros = p_CantMiembros + 1
Else
    NuevoMiembro = False
End If

End Function

Public Function SaleMiembro(ByVal userindex As Integer) As Boolean
'el valor de retorno representa si se disuelve la party
Dim i As Integer
Dim j As Integer
i = 1
SaleMiembro = False
While i <= PARTY_MAXMEMBERS And p_members(i).userindex <> userindex
    i = i + 1
Wend

If i = 1 Then
    'sale el founder, la party se disuelve
    SaleMiembro = True
    Call MandarMensajeAConsola("El lider disuelve la party.", "Servidor")
    For j = PARTY_MAXMEMBERS To 1 Step -1
        If p_members(j).userindex > 0 Then
            Call SendData(SendTarget.toIndex, p_members(j).userindex, 0, "|| Abandonas la party liderada por " & UserList(p_members(1).userindex).name & FONTTYPE_PARTY)
    
            Call SendData(SendTarget.toIndex, p_members(j).userindex, 0, "|| Durante la misma has conseguido " & p_members(j).Experiencia & " puntos de experiencia!" & FONTTYPE_PARTY)
    
            If Not PARTY_EXPERIENCIAPORGOLPE Then
                UserList(p_members(j).userindex).Stats.Exp = UserList(p_members(j).userindex).Stats.Exp + p_members(j).Experiencia
                If UserList(p_members(j).userindex).Stats.Exp > MAXEXP Then _
                    UserList(p_members(j).userindex).Stats.Exp = MAXEXP
                Call CheckUserLevel(p_members(j).userindex)
                Call EnviarExp(p_members(j).userindex)
            End If
            Call MandarMensajeAConsola(UserList(p_members(j).userindex).name & " abandona la party." & FONTTYPE_PARTY, "Servidor")
            UserList(p_members(j).userindex).PartyIndex = 0
            p_CantMiembros = p_CantMiembros - 1
            p_members(j).userindex = 0
            p_members(j).Experiencia = 0
        End If
    Next j
Else
    If i <= PARTY_MAXMEMBERS Then
        If Not PARTY_EXPERIENCIAPORGOLPE Then
            UserList(p_members(i).userindex).Stats.Exp = UserList(p_members(i).userindex).Stats.Exp + p_members(i).Experiencia
            If UserList(p_members(i).userindex).Stats.Exp > MAXEXP Then _
                UserList(p_members(i).userindex).Stats.Exp = MAXEXP
            Call CheckUserLevel(p_members(i).userindex)
            Call EnviarExp(p_members(i).userindex)
        End If
        Call MandarMensajeAConsola(UserList(p_members(i).userindex).name & " abandona la party." & FONTTYPE_PARTY, "Servidor")
        Call SendData(SendTarget.toIndex, p_members(i).userindex, 0, "|| Durante la misma has conseguido " & p_members(i).Experiencia & " puntos de experiencia!" & FONTTYPE_PARTY)
        p_CantMiembros = p_CantMiembros - 1
        p_members(i).userindex = 0
        p_members(i).Experiencia = 0
        CompactMemberList
    End If
End If

End Function

Public Function HacerLeader(ByVal userindex As Integer) As Boolean
Dim i As Integer
Dim OldLeader As Integer
Dim oldExp As Long
Dim UserIndexIndex As Integer

UserIndexIndex = 0
HacerLeader = True

For i = 1 To PARTY_MAXMEMBERS
    If p_members(i).userindex > 0 Then
        If p_members(i).userindex = userindex Then
            UserIndexIndex = i
            Exit For
        End If
        'chequeo la dif entre levels
        If Abs(UserList(userindex).Stats.ELV - UserList(p_members(i).userindex).Stats.ELV) > MAXPARTYDELTALEVEL Then
            Call SendData(SendTarget.toIndex, p_members(i).userindex, 0, "|| No se puede hacer líder a " & UserList(userindex).name & " pues la diferencia entre niveles de los miembros de la party sería mayor a " & MAXPARTYDELTALEVEL & FONTTYPE_PARTY)
            HacerLeader = False
        End If
    End If
Next i

If Not HacerLeader Then Exit Function

If UserIndexIndex = 0 Then
    'catastrofe! esto no deberia pasar nunca! pero como es AO.... :p
    Call LogError("INCONSISTENCIA DE PARTIES")
    Call SendData(SendTarget.ToAdmins, 0, 0, "|| Inconsistencia de parties en HACERLEADER (UII = 0), AVISE A UN PROGRAMADOR ESTO ES UNA CATASTROFE!!!!")
    HacerLeader = False
    Exit Function
End If


'aca esta todo bien y doy vuelta las collections
OldLeader = p_members(1).userindex
oldExp = p_members(1).Experiencia

p_members(1).userindex = p_members(UserIndexIndex).userindex     'que en realdiad es el userindex, pero no quiero inconsistencias moviendo experiencias
p_members(1).Experiencia = p_members(UserIndexIndex).Experiencia

p_members(i).userindex = OldLeader
p_members(i).Experiencia = oldExp

p_Fundador = p_members(1).userindex

'no need to compact
End Function


Public Sub ObtenerMiembrosOnline(ByRef MemberList As String)
'esta funcion viola un poco la correcta forma de hacer la clase
'pero bue...
Dim i As Integer
MemberList = "Nombre(Exp): "
For i = 1 To PARTY_MAXMEMBERS
    If p_members(i).userindex > 0 Then
        MemberList = MemberList & " - " & UserList(p_members(i).userindex).name & " (" & p_members(i).Experiencia & ")"
    End If
Next i

MemberList = MemberList & ". Experiencia total: " & p_expTotal

End Sub


Public Function PuedeEntrar(ByVal userindex As Integer, ByRef razon As String) As Boolean
'DEFINE LAS REGLAS DEL JUEGO PARA DEJAR ENTRAR A MIEMBROS
Dim esArmada As Boolean
Dim esCaos As Boolean
Dim MyLevel As Integer
Dim i As Integer
Dim rv As Boolean
Dim UI As Integer

rv = True
esArmada = (UserList(userindex).Faccion.ArmadaReal = 1)
esCaos = (UserList(userindex).Faccion.FuerzasCaos = 1)
MyLevel = UserList(userindex).Stats.ELV

rv = Distancia(UserList(p_members(1).userindex).pos, UserList(userindex).pos) <= MAXDISTANCIAINGRESOPARTY
If rv Then
    rv = (p_members(PARTY_MAXMEMBERS).userindex = 0)
    If rv Then
        For i = 1 To PARTY_MAXMEMBERS
            UI = p_members(i).userindex
            'pongo los casos que evitarian que pueda entrar
            'aspirante armada en party crimi
            If UI > 0 Then
                If esArmada And Criminal(UI) Then
                    razon = "Los miembros de la armada no entran a una party con criminales."
                    rv = False
                End If
                'aspirante caos en party ciuda
                If esCaos And Not Criminal(UI) Then
                    razon = "Los miembros de la legión oscura no entran a una party con ciudadanos."
                    rv = False
                End If
                'aspirante crimi en party armada
                If UserList(UI).Faccion.ArmadaReal = 1 And Criminal(userindex) Then
                    razon = "Los criminales no entran a partys con miembros de la armada real."
                    rv = False
                End If
                'aspirante ciuda en party caos
                If UserList(UI).Faccion.FuerzasCaos = 1 And Not Criminal(userindex) Then
                    razon = "Los miembros de la legión oscura no entran a una party con ciudadanos."
                    rv = False
                End If
                'diferencia entre niveles
                If Abs(UserList(UI).Stats.ELV - UserList(userindex).Stats.ELV) > MAXPARTYDELTALEVEL Then
                    razon = "La diferencia con " & UserList(UI).name & " es mayor a " & MAXPARTYDELTALEVEL & " niveles."
                    rv = False
                End If
            
                If Not rv Then Exit For 'violate una programacion estructurada
            End If
        Next i
    Else
        razon = "La mayor cantidad de miembros es " & PARTY_MAXMEMBERS
    End If
Else
    razon = "Te encuentras muy lejos del fundador."
End If

PuedeEntrar = rv

End Function


Public Sub FlushExperiencia()
'esta funcion se invoca frente a cerradas del servidor. Flushea la experiencia
'acumulada a los usuarios.

Dim i As Integer
If Not PARTY_EXPERIENCIAPORGOLPE Then   'esto sirve SOLO cuando acumulamos la experiencia!
    For i = 1 To PARTY_MAXMEMBERS
        If p_members(i).userindex > 0 Then
            If p_members(i).Experiencia > 0 Then
                UserList(p_members(i).userindex).Stats.Exp = UserList(p_members(i).userindex).Stats.Exp + p_members(i).Experiencia
                If UserList(p_members(i).userindex).Stats.Exp > MAXEXP Then _
                    UserList(p_members(i).userindex).Stats.Exp = MAXEXP
                Call CheckUserLevel(p_members(i).userindex)
            Else
                If Abs(UserList(p_members(i).userindex).Stats.Exp) > Abs(p_members(i).Experiencia) Then
                    UserList(p_members(i).userindex).Stats.Exp = UserList(p_members(i).userindex).Stats.Exp + p_members(i).Experiencia
                Else
                    UserList(p_members(i).userindex).Stats.Exp = 0
                End If
            End If
            p_members(i).Experiencia = 0
            Call EnviarExp(p_members(i).userindex)
        End If
    Next i
End If

End Sub




Private Sub CompactMemberList()
Dim i As Integer
Dim freeIndex As Integer
Dim tPlayer As tPartyMember
i = 1
While i <= PARTY_MAXMEMBERS
    If p_members(i).userindex = 0 And freeIndex = 0 Then
        freeIndex = i
    ElseIf p_members(i).userindex > 0 And freeIndex > 0 Then
        p_members(freeIndex).Experiencia = p_members(i).Experiencia
        p_members(freeIndex).userindex = p_members(i).userindex
        p_members(i).userindex = 0
        p_members(i).Experiencia = 0
        'muevo el de la pos i a freeindex
        i = freeIndex
        freeIndex = 0
    End If
    i = i + 1
Wend

End Sub

Public Function CantMiembros() As Integer
CantMiembros = p_CantMiembros
End Function
