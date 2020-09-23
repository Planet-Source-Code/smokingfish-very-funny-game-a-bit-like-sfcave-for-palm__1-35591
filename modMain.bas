Attribute VB_Name = "modMain"
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Game As New clsMain
Public Function FileExists(FullFileName As String) As Boolean
    On Error GoTo MakeF
        Open FullFileName For Input As #1
        Close #1
        FileExists = True
    Exit Function
MakeF:
        FileExists = False
    Exit Function
End Function
