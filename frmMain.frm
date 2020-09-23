VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SFCave"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3420
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3420
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2880
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   2880
      Top             =   240
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "End Game"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "About"
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Star Game"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim UpDown As Boolean
Dim EnemyTrue1 As Boolean
Dim EnemyTrue2 As Boolean
Dim EnemyTrue3 As Boolean
Dim Score1 As Integer
Dim Score2 As Integer

Private Sub Form_Load()
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Me.Show
End Sub

Public Sub MainLoop()
Dim TEMPa As Integer
Dim TEMPb As Integer
Do
DoEvents
Me.Cls
Sleep 5
Game.SnakeSpeed = Game.SnakeSpeed + Game.WorldGravity

If Game.SnakePositionY + Game.SnakeSpeed + 120 > 2500 Then
    Game.SnakeSpeed = -Int(Game.SnakeSpeed * 0.8)
    Game.SnakePositionY = 2500 - 120
Else
    Game.SnakePositionY = Game.SnakePositionY + Game.SnakeSpeed
End If

If Game.SnakeDirection = True Then
    If Game.SnakePositionX + 120 + Game.SnakeAcceleration > 3200 Then
        Game.SnakeDirection = False
        Game.SnakePositionX = 3200 - 120
        Game.SnakeAcceleration = Game.SnakeAcceleration + 4
    Else
    Game.SnakePositionX = Game.SnakePositionX + Game.SnakeAcceleration
    End If
End If

If Game.SnakeDirection = False Then
    If Game.SnakePositionX - Game.SnakeAcceleration < 350 Then
        Game.SnakeDirection = True
        Game.SnakePositionX = 350
        Game.SnakeAcceleration = Game.SnakeAcceleration + 4
    Else
        Game.SnakePositionX = Game.SnakePositionX - Game.SnakeAcceleration
    End If
End If

If Game.SnakePositionY <= 350 Then
Game.SnakePositionY = 350
Game.SnakeSpeed = -Int(Game.SnakeSpeed * 0.8)
End If

If GetAsyncKeyState(vbKeySpace) Then
Sleep 10
Game.SnakeSpeed = Game.SnakeSpeed - 4
End If

Game.DrawSnake
Game.DrawWorld
Game.DrawEnemys

If Not frmMain.Point(Game.SnakePositionX - 75, Game.SnakePositionY - 75) = vbBlack Then
Game.SnakeOver = True
End If
If Not frmMain.Point(Game.SnakePositionX + 75, Game.SnakePositionY - 75) = vbBlack Then
Game.SnakeOver = True
End If
If Not frmMain.Point(Game.SnakePositionX - 75, Game.SnakePositionY + 75) = vbBlack Then
Game.SnakeOver = True
End If
If Not frmMain.Point(Game.SnakePositionX + 75, Game.SnakePositionY + 75) = vbBlack Then
Game.SnakeOver = True
End If
Score1 = Score1 + 1
frmMain.CurrentX = 1100
frmMain.CurrentY = 2200
frmMain.Print "Score: " & Score1
frmMain.CurrentX = 900
frmMain.CurrentY = 0
frmMain.Print "Highscore: " & Score2
Loop Until Game.SnakeOver = True
frmMain.Cls
Game.DrawWorld
For i = 1 To 100
Sleep 5
Game.GradientCircle frmMain, Game.SnakePositionX, Game.SnakePositionY, i * 5, 200, 50, 50, 5, False, False
frmMain.Refresh
Next i

Sleep 1000
frmMain.Cls
frmMain.CurrentX = 1100
frmMain.CurrentY = 500
frmMain.Print "Game Over!!"
frmMain.CurrentX = 1100
frmMain.CurrentY = 2200
frmMain.Print "Score: " & Score1
If Score2 < Score1 Then
Open App.Path & "\Score.Dat" For Binary As #1
Put #1, , Score1
Close #1
MsgBox "NEW HIGHSCORE!!!"
End If
Me.Refresh
Sleep 3000
Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label1_Click()
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Score1 = 0
Score2 = 0
If FileExists(App.Path & "\Score.Dat") = True Then
Open App.Path & "\Score.Dat" For Binary As #1
Get #1, , Score2
Close #1
End If
Game.SnakePositionX = 400
Game.SnakePositionY = 1250
Game.SnakeSpeed = 0
Game.SnakeDirection = True
Game.WorldGravity = 1.6
Game.DummiWorld
Game.SnakeAcceleration = 10
Game.SideMove = 25
Game.WorldWidth = 150
UpDown = True
Call MainLoop
End Sub

Private Sub Label2_Click()
MsgBox "2002 by SmokingFish!" & vbCrLf & "mail@SmokingFish.de", vbInformation
End Sub

Private Sub Label3_Click()
End
End Sub

Private Sub Timer1_Timer()
If UpDown = True Then
Game.WorldWidth = Game.WorldWidth - 1
Else
Game.WorldWidth = Game.WorldWidth + 1
End If
If Game.WorldWidth = 80 Then UpDown = False
If Game.WorldWidth = 110 Then UpDown = True
End Sub
