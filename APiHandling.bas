Attribute VB_Name = "APiHandling"
Public Declare Function SetDoubleClickTime Lib "user32" (ByVal wCount As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


Public Sub SetDCLick(tTime As Long)

End Sub
