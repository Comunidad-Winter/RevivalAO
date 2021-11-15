Attribute VB_Name = "Soporte"
Public Sub EnviarWeas()
SendData "CTMR" & _
frmSoporteGm.Label1.Caption _
& Chr$(2) & frmSoporteGm.Text2.Text
End Sub

