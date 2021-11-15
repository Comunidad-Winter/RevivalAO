Attribute VB_Name = "ModConsola"
Public Sub addConsole(Texto As String, Rojo As Byte, Verde As Byte, Azul As Byte, Bold As Boolean, Italic As Boolean, Optional ByVal Enter As Boolean = False)
    With frmMain.RichTextBox1
        If (Len(.Text)) > 700 Then .Text = ""
        
        .SelStart = Len(.Text)
        .SelLength = 0
        
        .SelBold = Bold
        .SelItalic = Italic
        
        .SelColor = RGB(Rojo, Verde, Azul)
        
        .SelText = IIf(Enter, Texto, Texto & vbCrLf)
        
        .Refresh
    End With

End Sub

Public Sub addConsolee(Texto As String, Rojo As Byte, Verde As Byte, Azul As Byte, Bold As Boolean, Italic As Boolean, Optional ByVal Enter As Boolean = False)
    With frmMain.RichTextBox2
        If (Len(.Text)) > 700 Then .Text = ""
        
        .SelStart = Len(.Text)
        .SelLength = 0
        
        .SelBold = Bold
        .SelItalic = Italic
        
        .SelColor = RGB(Rojo, Verde, Azul)
        
        .SelText = IIf(Enter, Texto, Texto & vbCrLf)
        
        .Refresh
    End With

End Sub
