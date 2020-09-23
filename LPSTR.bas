Attribute VB_Name = "LPSTR"


Public Debuging As Boolean
Public Watching As Boolean
Public Const Run_File = 0
Public Const Debug_File = 1
Public Const Watch_File = 2
Public CL As Long
Public RunThis As String
Public ScriptCommand As String
Public Somthing As String 'Can't think of any varibles names lol
Public File_Type As Byte

Public Sub GetScript()
Debug.Print "Getting it"
Open App.Path & "/start.gff" For Input As #1
      Do While Not EOF(1)
        Line Input #1, Somthing
        CL = CL + 1
        If Somthing = "" Then Exit Sub
        If CL = 1 Then File_Type = Somthing
        If CL = 2 Then RunThis = Somthing
    Loop
Close #1
Debug.Print "Got FIle Type " & File_Type
Select Case File_Type
    Case Run_File
        RunFile (RunThis)
        Debug.Print "Running it"
End Select

End Sub

