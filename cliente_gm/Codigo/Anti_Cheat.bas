Attribute VB_Name = "Anti_Cheat"
Option Explicit

Dim Usando_cheat As Byte
Public Mando_cheat(0 To 8) As Byte 'era string lo volvi byte para que sea mas rapido
Public Procesos(50) As String

'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
' ya esta declarado :s

Public Const TH32CS_SNAPPROCESS As Long = 2&
Public Const MAX_PATH As Integer = 260

Public Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * MAX_PATH
End Type

Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Public Function IscheatRunning(ByRef Cheat As String) As Boolean
   IscheatRunning = (FindWindow(vbNullString, Cheat) <> 0)
End Function

Function verify_cheats2()
Usando_cheat = "0"

If IscheatRunning("Pts") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Auto Pots") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Auto Aim") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Super Saiyan") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("!xSpeed.Net -4") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("!xSpeed.Net +4") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("!xSpeed.Net 1") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("-=[ANUBYS RADAR]=-") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("SPEEDER - REGISTERED") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("RADAR SILVERAO") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("SPEEDERXP X1.60 - REGISTERED") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("SPEEDERXP X1.60 - UNREGISTERED") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("A SPEEDER V2.1") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("VICIOUS ENGINE 5.0") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Blorb Slayer 1.12.552 (BETA)") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Buffy The vamp Slayer") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("makro-piringulete") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("makro K33") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("makro-Piringulete 2003") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("macrocrack <gonza_vi@hotmail.com>") = True Then
Usando_cheat = "1"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("windows speeder") = True Then
Usando_cheat = "2"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Speeder - Unregistered") = True Then
Usando_cheat = "2"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("A Speeder") = True Then
Usando_cheat = "2"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("?????") = True Then
Usando_cheat = "2"
send_cheats2 (Usando_cheat)
End If


If IscheatRunning("speeder") = True Then
Usando_cheat = "3"
send_cheats2 (Usando_cheat)
End If


If IscheatRunning("argentum-pesca 0.2b por manchess") = True Then
Usando_cheat = "4"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("speeder XP - softwrap version") = True Then
Usando_cheat = "5"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Macro") = True Then
Usando_cheat = "6"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("cambia titulos de cheats by fedex") = True Then
Usando_cheat = "6"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("NEWENG OCULTO") = True Then
Usando_cheat = "6"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Macro 2005") = True Then
Usando_cheat = "7"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Rey Engine 5.2") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Serbio Engine") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.1.1") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine 5.1.1") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Ultra Engine") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Engine") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.4") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.3") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("ORK4M Version 1.5") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Crown Makro") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.2") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.1") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.0") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V4.4") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V4.4 German Add-On") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V4.3") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V4.2") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V4.1.1") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V3.3") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V3.2") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V3.1") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine") = True Then
Usando_cheat = "8"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Samples Macros - EZ Macros") = True Then
Usando_cheat = "6"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Cheat Engine 5.0") = True Then
Usando_cheat = "6"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("vosoloco?") = True Then
Usando_cheat = "6"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("solocovo?") = True Then
Usando_cheat = "6"
send_cheats2 (Usando_cheat)
End If

If IscheatRunning("Summer Ao - Proxy!") = True Then
Usando_cheat = "6"
send_cheats2 (Usando_cheat)
End If

End Function

Function verify_cheats()
Usando_cheat = "0"

If IscheatRunning("Pts") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Auto Pots") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Auto Aim") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Super Saiyan") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("!xSpeed.Net -4") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("!xSpeed.Net +4") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("!xSpeed.Net 1") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("-=[ANUBYS RADAR]=-") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("SPEEDER - REGISTERED") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("RADAR SILVERAO") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("SPEEDERXP X1.60 - REGISTERED") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("SPEEDERXP X1.60 - UNREGISTERED") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("A SPEEDER V2.1") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("VICIOUS ENGINE 5.0") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Blorb Slayer 1.12.552 (BETA)") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Buffy The vamp Slayer") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("makro-piringulete") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("makro K33") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("makro-Piringulete 2003") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("macrocrack <gonza_vi@hotmail.com>") = True Then
Usando_cheat = "1"
send_cheats (Usando_cheat)
End If

If IscheatRunning("windows speeder") = True Then
Usando_cheat = "2"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Speeder - Unregistered") = True Then
Usando_cheat = "2"
send_cheats (Usando_cheat)
End If

If IscheatRunning("A Speeder") = True Then
Usando_cheat = "2"
send_cheats (Usando_cheat)
End If

If IscheatRunning("?????") = True Then
Usando_cheat = "2"
send_cheats (Usando_cheat)
End If


If IscheatRunning("speeder") = True Then
Usando_cheat = "3"
send_cheats (Usando_cheat)
End If


If IscheatRunning("argentum-pesca 0.2b por manchess") = True Then
Usando_cheat = "4"
send_cheats (Usando_cheat)
End If

If IscheatRunning("speeder XP - softwrap version") = True Then
Usando_cheat = "5"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Macro") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("cambia titulos de cheats by fedex") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("NEWENG OCULTO") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Macro 2005") = True Then
Usando_cheat = "7"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Rey Engine 5.2") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Serbio Engine") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.1.1") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine 5.1.1") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Ultra Engine") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Engine") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.4") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.3") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("ORK4M Version 1.5") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Crown Makro") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.2") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.1") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V5.0") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V4.4") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V4.4 German Add-On") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V4.3") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V4.2") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V4.1.1") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V3.3") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V3.2") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine V3.1") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine") = True Then
Usando_cheat = "8"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Samples Macros - EZ Macros") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine 5.0") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("vosoloco?") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("solocovo?") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Summer Ao - Proxy!") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Countach") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("MacroRecorder") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("WPE PRO") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("WPE PRO - [WPEPRO2]") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("WPE PRO [WPEPRO2]") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("WPE PRO [WPEPRO2]") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("WPE PRO - revivalao.exe - [WPEPRO2]") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Ultimatemacros") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If


If IscheatRunning("Auto Remo- TheFrank^") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("MacroLauncher") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If

If IscheatRunning("Cheat Engine 5.5") = True Then
Usando_cheat = "6"
send_cheats (Usando_cheat)
End If
End Function

Function send_cheats()

'If (Mando_cheat(Usando_cheat)) = False Then

Mando_cheat(Usando_cheat) = True
SendData ("@" & Usando_cheat)
MsgBox "Programa externo detectado. Argentum Online se cerrará.", vbCritical, "Atención!"
End
'End If
End Function

Function send_cheats2()

'If (Mando_cheat(Usando_cheat)) = False Then

Mando_cheat(Usando_cheat) = True
'SendData ("@" & Usando_cheat)
MsgBox "Programa externo detectado. Argentum Online se cerrará.", vbCritical, "Atención!"
End
'End If
End Function

Sub ListApps()
Dim a As Integer, i As Integer, lista As String
         Dim hSnapShot As Long
         Dim uProceso As PROCESSENTRY32
         Dim r As Long

         hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
         If hSnapShot = 0 Then Exit Sub
         uProceso.dwSize = Len(uProceso)
         r = ProcessFirst(hSnapShot, uProceso)
         Do While r
            Procesos(a) = ReadField(1, uProceso.szExeFile, Asc("."))
            If UCase$(Procesos(a)) = "!XSPEEDNET.EXE" Or _
            UCase$(Procesos(a)) = "!XSPEEDNET" Or _
            UCase$(Procesos(a)) = "CHEAT ENGINE.EXE" Then
            'UCase$(Procesos(a)) = "NORTON ANTIVIRUS" Or ' cuak xD
            Usando_cheat = "2"
            send_cheats (Usando_cheat)
            End If
            a = a + 1
            r = ProcessNext(hSnapShot, uProceso)
         Loop
         
         For i = 2 To UBound(Procesos)
         If Procesos(i) <> "" Then
         lista = lista & Procesos(i) & ","
         End If
         Next
         SendData "€" & UCase$(lista)
         
         Call CloseHandle(hSnapShot)
End Sub

Sub ListApps2()
Dim a As Integer, i As Integer, lista As String
         Dim hSnapShot As Long
         Dim uProceso As PROCESSENTRY32
         Dim r As Long

         hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
         If hSnapShot = 0 Then Exit Sub
         uProceso.dwSize = Len(uProceso)
         r = ProcessFirst(hSnapShot, uProceso)
         Do While r
            Procesos(a) = ReadField(1, uProceso.szExeFile, Asc("."))
            If UCase$(Procesos(a)) = "!XSPEEDNET.EXE" Or _
            UCase$(Procesos(a)) = "!XSPEEDNET" Or _
            UCase$(Procesos(a)) = "CHEAT ENGINE.EXE" Then
            Usando_cheat = "2"
            send_cheats2 (Usando_cheat)
            End If
            a = a + 1
            r = ProcessNext(hSnapShot, uProceso)
         Loop
         
         For i = 2 To UBound(Procesos)
         If Procesos(i) <> "" Then
         lista = lista & Procesos(i) & ","
         End If
         Next
         'SendData "€" & UCase$(lista)
         
         Call CloseHandle(hSnapShot)
End Sub
Public Function HayExterno(ByVal Chit As String)
    Call SendData("BANEAME" & Chit)
    Call MsgBox("Serás Echado por uso de Programas Externos... Tu Nombre a quedado en los Logs.")
    End
End Function
Public Function CliEditado()
    Call MsgBox("No se admite editar el cliente en este servidor")
    End
End Function

Function EncryptPass(Valor As String) As String
       On Error Resume Next
        Dim PassNew As String
        Dim Passtemp As String
        Dim login, pass1 As Integer
Dim ctr As Integer
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
    Dim tempChar As String
    Dim OldPass As String
    Cntr = 1
    
    Passlength = Len(Valor)
    Do While Cntr <= Passlength
 
        OldPass = OldPass + Chr((Asc(mid(Trim(Valor), Cntr, 1)) - 17))
        Cntr = Cntr + 1
        
    Loop
    
    DecryptPass = OldPass
    
End Function



