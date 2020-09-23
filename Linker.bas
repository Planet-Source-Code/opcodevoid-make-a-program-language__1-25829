Attribute VB_Name = "Linker"
Public ParentName() As String
Public SelParent As Byte
Public SelSub As Byte
Public SelMenu As Integer
Public NumChild() As Long


'File opening Constance
Public Const Run_File = 0
Public Const Debug_File = 1
Public Const Watch_File = 2

'For watch mode saves Height
Public Const WatchTop = 6960
Public Const Non_Watch = 5310

Public WatchFile As Boolean
Public Type Project 'Project Info
    pName As String
    pPath As String
End Type

Public FlagTrue As Boolean 'is a project open

Public p As Project 'Project type

