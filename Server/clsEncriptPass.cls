VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEncriptPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim login, pass1 As Integer
Dim ctr As Integer


 Function EncryptPass(Valor As String) As String
       On Error Resume Next
        Dim PassNew As String
        Dim Passtemp As String
        
        pass1 = Len(Trim(Valor))
        
        ctr = 1
        Do While ctr <= pass1
            
            PassNew = CStr(PassNew) & Chr((Asc(Mid(Trim(Valor), ctr, 1)) + 17))
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
 
        OldPass = OldPass + Chr((Asc(Mid(Trim(Valor), Cntr, 1)) - 17))
        Cntr = Cntr + 1
        
    Loop
    
    DecryptPass = OldPass
    
End Function

