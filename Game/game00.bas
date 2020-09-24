Attribute VB_Name = "Module1"
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Function DoKeys()
If GetAsyncKeyState(vbKeyLeft) Then
Form1.player1.Left = Form1.player1.Left - 20
Form1.player1.Picture = Form1.dir(2).Picture
If Form1.bullet.Visible = False Then
Form1.direction.Caption = "left"
End If
End If
If GetAsyncKeyState(vbKeyRight) Then
Form1.player1.Left = Form1.player1.Left + 20
Form1.player1.Picture = Form1.dir(3).Picture
If Form1.bullet.Visible = False Then
Form1.direction.Caption = "right"
End If
End If
If GetAsyncKeyState(vbKeyUp) Then
Form1.player1.Top = Form1.player1.Top - 20
Form1.player1.Picture = Form1.dir(0).Picture
If Form1.bullet.Visible = False Then
Form1.direction.Caption = "up"
End If
End If
If GetAsyncKeyState(vbKeyDown) Then
Form1.player1.Top = Form1.player1.Top + 20
Form1.player1.Picture = Form1.dir(1).Picture
If Form1.bullet.Visible = False Then
Form1.direction.Caption = "down"
End If
End If
If Form1.player1.Left < 0 Then
Form1.player1.Left = 0
End If
If Form1.player1.Top < 0 Then
Form1.player1.Top = 0
End If

If Form1.player1.Left > Form1.Back.ScaleWidth - Form1.player1.Width Then
Form1.player1.Left = Form1.Back.ScaleWidth - Form1.player1.Width
End If
If Form1.player1.Top > Form1.Back.ScaleHeight - Form1.player1.Height Then
Form1.player1.Top = Form1.Back.ScaleHeight - Form1.player1.Height
End If
End Function


