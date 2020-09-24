VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   2280
   End
   Begin VB.TextBox time 
      Height          =   285
      Left            =   4560
      TabIndex        =   18
      Text            =   "0"
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton NewGame 
      Caption         =   "New Game"
      Height          =   255
      Left            =   5880
      TabIndex        =   8
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2880
      Top             =   2280
   End
   Begin VB.PictureBox Back 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DrawWidth       =   8
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   0
      Picture         =   "game00.frx":0000
      ScaleHeight     =   4905
      ScaleWidth      =   6945
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6975
      Begin VB.PictureBox dir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   3
         Left            =   6720
         Picture         =   "game00.frx":717E2
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.PictureBox dir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   2
         Left            =   6480
         Picture         =   "game00.frx":71920
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.PictureBox dir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   1
         Left            =   6600
         Picture         =   "game00.frx":71A5E
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.PictureBox dir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   0
         Left            =   6600
         Picture         =   "game00.frx":71B9C
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   2640
         Picture         =   "game00.frx":71CDA
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   7
         Top             =   3840
         Width           =   135
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   2280
         OLEDropMode     =   1  'Manual
         Picture         =   "game00.frx":71E18
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   6
         Top             =   3480
         Width           =   135
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   2520
         Picture         =   "game00.frx":71F56
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   5
         Top             =   3720
         Width           =   135
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   2160
         Picture         =   "game00.frx":72094
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   4
         Top             =   3360
         Width           =   135
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   2400
         Picture         =   "game00.frx":721D2
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   3
         Top             =   3600
         Width           =   135
      End
      Begin VB.PictureBox bullet 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   1440
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.PictureBox player1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   1320
         Picture         =   "game00.frx":72310
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   1
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label direction 
         Caption         =   "down"
         Height          =   255
         Left            =   6480
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Label hidden 
      Caption         =   "0"
      Height          =   495
      Left            =   2880
      TabIndex        =   20
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Seconds"
      Height          =   255
      Left            =   4920
      TabIndex        =   19
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Time Limit:"
      Height          =   255
      Left            =   3720
      TabIndex        =   17
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label sshoot 
      BackColor       =   &H000000FF&
      Caption         =   "Ready To Shoot"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Score:"
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label Score 
      Caption         =   "0"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   4920
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function Check_Collision_BF(Sprite1 As PictureBox, Sprite2 As PictureBox) As Boolean
 
 'Simply check the bounderies of the 2 Picture Box's
 If Sprite1.Left > Sprite2.Left + Sprite2.Width Or _
    Sprite1.Left + Sprite1.Width < Sprite2.Left Or _
    Sprite1.Top > Sprite2.Top + Sprite2.Height Or _
    Sprite1.Top + Sprite1.Height < Sprite2.Top Then
  Check_Collision_BF = False
 Else
  Check_Collision_BF = True
 End If
 
End Function

Private Sub NewGame_Click()
If time.Text < 1 Then
Exit Sub
End If
Back.Visible = True
Score.Caption = "0"
hidden.Caption = time.Text
Timer2.Enabled = True
time.Enabled = False
NewGame.Enabled = False
Picture4.BackColor = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
Picture5.BackColor = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
Picture6.BackColor = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
Picture7.BackColor = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
Picture8.BackColor = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
Picture4.Move Int(Rnd * Back.ScaleWidth), Int(Rnd * Back.ScaleHeight)
If Picture4.Top < 100 Then
Picture4.Top = 200
End If
Picture5.Move Int(Rnd * Back.ScaleWidth), Int(Rnd * Back.ScaleHeight)
If Picture5.Top < 100 Then
Picture5.Top = 200
End If
Picture6.Move Int(Rnd * Back.ScaleWidth), Int(Rnd * Back.ScaleHeight)
If Picture6.Top < 100 Then
Picture6.Top = 200
End If
Picture7.Move Int(Rnd * Back.ScaleWidth), Int(Rnd * Back.ScaleHeight)
If Picture7.Top < 100 Then
Picture7.Top = 200
End If
Picture8.Move Int(Rnd * Back.ScaleWidth), Int(Rnd * Back.ScaleHeight)
If Picture8.Top < 100 Then
Picture8.Top = 200
End If
End Sub

Private Sub Timer1_Timer()
If bullet.Visible = False Then
If player1.Picture = dir(0).Picture Then
direction.Caption = "up"
End If
If player1.Picture = dir(1).Picture Then
direction.Caption = "down"
End If
If player1.Picture = dir(2).Picture Then
direction.Caption = "left"
End If
If player1.Picture = dir(3).Picture Then
direction.Caption = "right"
End If
End If
If bullet.Visible = True Then
Back.Line (bullet.Left, bullet.Top - 5)-(bullet.Left, bullet.Top), RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
Back.Line (bullet.Left + bullet.Width, bullet.Top + bullet.Height - 5)-(bullet.Left + bullet.Width, bullet.Top + bullet.Height), RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
End If
If Picture4.BackColor = vbBlack Then
Picture4.Top = Picture4.Top + 60
End If
If Picture5.BackColor = vbBlack Then
Picture5.Top = Picture5.Top + 60
End If
If Picture6.BackColor = vbBlack Then
Picture6.Top = Picture6.Top + 60
End If
If Picture7.BackColor = vbBlack Then
Picture7.Top = Picture7.Top + 60
End If
If Picture8.BackColor = vbBlack Then
Picture8.Top = Picture8.Top + 60
End If

DoKeys
If GetAsyncKeyState(vbKeyReturn) Then
bullet.Visible = True
End If
If bullet.Visible = False Then
bullet.Move player1.Left, player1.Top
sshoot.Visible = True
Exit Sub
End If
sshoot.Visible = False
Form1.Caption = "Wierdo"
'collision'''''''''''''''''''''''''''''
 If Check_Collision_BF(bullet, Picture4) = True Then
Picture4.Move Int(Rnd * Back.ScaleWidth), Int(Rnd * Back.ScaleHeight)
If Picture4.Top < 100 Then
Picture4.Top = 200
End If
Score.Caption = Score.Caption + 1
End If
 If Check_Collision_BF(bullet, Picture5) = True Then
Picture5.Move Int(Rnd * Back.ScaleWidth), Int(Rnd * Back.ScaleHeight)
If Picture5.Top < 100 Then
Picture5.Top = 200
End If
Score.Caption = Score.Caption + 1
End If
 If Check_Collision_BF(bullet, Picture6) = True Then
Picture6.Move Int(Rnd * Back.ScaleWidth), Int(Rnd * Back.ScaleHeight)
If Picture6.Top < 100 Then
Picture6.Top = 200
End If
Score.Caption = Score.Caption + 1
End If
 If Check_Collision_BF(bullet, Picture7) = True Then
Picture7.Move Int(Rnd * Back.ScaleWidth), Int(Rnd * Back.ScaleHeight)
If Picture7.Top < 100 Then
Picture7.Top = 200
End If
Score.Caption = Score.Caption + 1
End If
 If Check_Collision_BF(bullet, Picture8) = True Then
Picture8.Move Int(Rnd * Back.ScaleWidth), Int(Rnd * Back.ScaleHeight)
If Picture8.Top < 100 Then
Picture8.Top = 200
End If
Score.Caption = Score.Caption + 1
End If

''''''''''''''''''
';;;;Moving Bomb;;;;;'
If bullet.Visible = False Then
bullet.Move player1.Left, player1.Top
End If
If bullet.Visible = True Then
If direction.Caption = "up" Then
bullet.Top = bullet.Top - 70
End If
If direction.Caption = "down" Then
bullet.Top = bullet.Top + 70
End If
If direction.Caption = "right" Then
bullet.Left = bullet.Left + 70
End If
If direction.Caption = "left" Then
bullet.Left = bullet.Left - 70
End If
 If Check_Collision_BF(bullet, Back) = False Then
bullet.Visible = False
End If
End If
If bullet.Top > Back.ScaleHeight Then
bullet.Visible = False
End If
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
time.Text = time.Text - 1
If time.Text = 0 Then
Timer2.Enabled = False
MsgBox "Game over. You made " & Score.Caption & " hits in " & hidden.Caption & " seconds!"
Score.Caption = "0"
time.Enabled = True
NewGame.Enabled = True
Back.Visible = False
End If
End Sub
