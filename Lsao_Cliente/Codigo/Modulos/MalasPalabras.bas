Attribute VB_Name = "MalasPalabras"
Option Compare Text
'[Gabriel Mellace]
'Utilizamos opcion de comparar para que no halla distincion entre minusculas y mayusculas ya que si no hay que poner Pt PT pT
'[Gabriel Mellace]
Public Function EsMalapalabra(ByVal Rdata As String)
Select Case UCase$(Rdata)
Case "NW"
Case "PT"
Case "PUTO"
Case "CONCHA"
Case "TROLA"
Case "PUTAZO"
Case "PIJA"
Case "VERGA"
Case "MIERDA"
Case "C o n c h a"
Case "MANKO"
Case "MANCO"
Case "NEWBIE"
Case "CAJETA"
Case "SENEB"
Case "ORGASMO"
Case "FANNY"
Case "PORONGA"
Case "CHUPAME"
Case "SOBALA"
Case "CABE"
Case "KB"
Case "TKB"
Case "GIL"
Case "BOBO"
Case "TONTO"
Case "SALAME"
Case "VIRGO"
Case "VIRGEN"
Case "NEGRO"
Case "PUTASO"
Case "PUTAZO"
Case "MANKISIMO"
Case "DROGADO"
Case "ORTO"
Case "OJETE"
Case "CULO"
Case Else
EsMalapalabra = False
Exit Function
End Select

EsMalapalabra = True
Call SendData(";" & "$_@%$!")
frmMain.SendTxt.Text = ""
stxtbuffer = ""
KeyCode = 0
frmMain.SendTxt.Visible = False
'[Gabriel Mellace]
End Function
