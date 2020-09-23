VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public SnakeSpeed As Integer
Public SnakePositionX As Integer
Public SnakePositionY As Integer
Public SnakeOver As Boolean
Public WorldGravity As Integer
Public SnakeDirection As Boolean
Public SnakeAcceleration As Integer
Private EnemyX(0 To 2) As Integer, EnemyY(0 To 2) As Integer
Private LX(0 To 250) As Long, RX(0 To 250) As Long
Public SideMove As Integer
Public WorldWidth As Long

Private Sub Cycle()
    Dim a As Integer
    For a = 250 To 1 Step -1
        LX(a) = LX(a - 1)
        LX(0) = ((150 - WorldWidth) / 2) + SideMove
        RX(a) = RX(a - 1)
        RX(0) = LX(0) + WorldWidth
    Next a
End Sub

Public Sub DrawEnemys()
EnemyY(0) = EnemyY(0) + 3
EnemyY(1) = EnemyY(1) + 5
EnemyY(2) = EnemyY(2) + 2
For i = 0 To 2
If frmMain.Point(EnemyX(i), EnemyY(i)) = vbBlue Then
EnemyY(i) = 400
End If
Next i
GradientCircle frmMain, EnemyX(0), EnemyY(0), 50, 200, 50, 30, 4, True, False
GradientCircle frmMain, EnemyX(1), EnemyY(1), 50, 200, 50, 30, 4, True, False
GradientCircle frmMain, EnemyX(2), EnemyY(2), 50, 200, 50, 30, 4, True, False
End Sub

Private Sub SidesChange()
    SideMove = SideMove + Round((Rnd * 2), 1) - 1
    If SideMove > 100 Then SideMove = 100
    If SideMove < 5 Then SideMove = 5
End Sub

Public Sub DrawSnake()
GradientCircle frmMain, SnakePositionX, SnakePositionY, 60, 0, 200, 200, 5, True, False
End Sub

Public Sub DummiWorld()
EnemyX(0) = 500
EnemyX(1) = 1800
EnemyX(2) = 3000
EnemyY(0) = 400
EnemyY(1) = 400
EnemyY(2) = 400
    For a = 1 To 250
        LX(a) = (24 + (Rnd * 1))
        RX(a) = LX(a) + WorldWidth
    Next a
End Sub

Public Sub DrawWorld()
Dim a As Integer, b As Integer
    For a = 1 To 250 Step 1
        RX(a) = LX(a) + WorldWidth
        frmMain.Line (a * 12 + 220, 200)-(a * 12 + 220, 10 + LX(a) * 12), vbRed
        frmMain.Line (a * 12 + 220, RX(a) * 12)-(a * 12 + 220, 2500), vbBlue
    Next a
    Cycle
SidesChange
GradientLine frmMain, 200, 200, 200, 2500, 100, 200, 100, 5, False
GradientLine frmMain, 200, 200, 3200, 200, 100, 200, 100, 5, False
GradientLine frmMain, 200, 2500, 3200, 2500, 100, 200, 100, 5, False
GradientLine frmMain, 3200, 200, 3200, 2500, 100, 200, 100, 5, False
GradientCircle frmMain, 200, 200, 50, 100, 200, 200, 5, True, False
GradientCircle frmMain, 3200, 200, 50, 100, 200, 200, 5, True, False
GradientCircle frmMain, 200, 2500, 50, 100, 200, 200, 5, True, False
GradientCircle frmMain, 3200, 2500, 50, 100, 200, 200, 5, True, False
End Sub

Public Function GradientLine(Obj, Xpos1, Ypos1, xpos2, Ypos2, ColorRed, ColorGreen, ColorBlue, BorderWidth, invert As Boolean)
    TBorder = Obj.DrawWidth
    If BorderWidth < 1 Then BorderWidth = 1
    For q = 1 To BorderWidth
        twidth = BorderWidth + 1 - q
        rval = Val(ColorRed / BorderWidth * q)
        If rval > 255 Then rval = 255
        gval = Val(ColorGreen / BorderWidth * q)
        If gval > 255 Then gval = 255
        bval = Val(ColorBlue / BorderWidth * q)
        If bval > 255 Then bval = 255
        If invert Then
            rval = 255 - rval
            gval = 255 - gval
            bval = 255 - bval
        End If
        Obj.DrawWidth = twidth
        Obj.Line (Xpos1, Ypos1)-(xpos2, Ypos2), RGB(rval, gval, bval)
    Next q
    Obj.DrawWidth = TBorder
End Function

Public Function GradientCircle(Obj, Xpos, Ypos, Radius, ColorRed, ColorGreen, ColorBlue, BorderWidth, Fill As Boolean, invert As Boolean)
    TBorder = Obj.DrawWidth
    tfill = Obj.FillStyle
    tfc = Obj.FillColor
    If Fill Then
        Obj.FillStyle = 0
        If Not invert Then Obj.FillColor = RGB(ColorRed, ColorGreen, ColorBlue) Else Obj.FillColor = RGB(255 - ColorRed, 255 - ColorGreen, 255 - ColorBlue)
    End If
    If BorderWidth < 1 Then BorderWidth = 1
    For q = 1 To BorderWidth
        twidth = BorderWidth + 1 - q
        rval = Val(ColorRed / BorderWidth * q)
        If rval > 255 Then rval = 255
        gval = Val(ColorGreen / BorderWidth * q)
        If gval > 255 Then gval = 255
        bval = Val(ColorBlue / BorderWidth * q)
        If bval > 255 Then bval = 255
        If invert Then
            rval = 255 - rval
            gval = 255 - gval
            bval = 255 - bval
        End If
        Obj.DrawWidth = twidth
        Obj.Circle (Xpos, Ypos), Radius, RGB(rval, gval, bval)
    Next q
    Obj.FillStyle = tfill
    Obj.FillColor = tfc
    Obj.DrawWidth = TBorder
End Function
