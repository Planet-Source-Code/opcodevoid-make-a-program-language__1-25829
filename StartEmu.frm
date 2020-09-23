VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form StartEmu 
   Caption         =   "Flamming Ice Station"
   ClientHeight    =   5205
   ClientLeft      =   1290
   ClientTop       =   930
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   10335
   Begin VB.CommandButton CmdLPSTR 
      Caption         =   "Command1"
      Height          =   255
      Left            =   9120
      TabIndex        =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   9120
      Top             =   -240
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   9120
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File "
      Begin VB.Menu mnuLoad 
         Caption         =   "Load File"
      End
   End
   Begin VB.Menu mnuRom 
      Caption         =   "Rom"
      Begin VB.Menu mnuDebug 
         Caption         =   "Debug"
      End
      Begin VB.Menu mnuWatch 
         Caption         =   "Watch"
      End
   End
End
Attribute VB_Name = "StartEmu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdLPSTR_Click()
Open App.Path & "/start.gff" For Input As #5
      Do While Not EOF(5)
        Line Input #5, Somthing
        CL = CL + 1
        If Somthing = "" Then Exit Sub
        If CL = 1 Then File_Type = Somthing
        If CL = 2 Then RunThis = Somthing
    Loop
Close #5
Select Case File_Type
    Case Run_File
        RunFile (RunThis)
     Case Debug_File
        Debuging = True
        RunFile (RunThis)
    Case Watch_File
        Watching = True
        RunFile (RunThis)
End Select
Open App.Path & "/start.gff" For Output As #5
    Print #5, ""
Close #5
End Sub



Private Sub Form_Load()
MenuClick = 3000

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuLoad_Click()
'A little Test Function
'CD.ShowOpen
'Dim c As String
'c = 1
'Dim d
'Dim onebit As String * 1
'
'Open CD.FileName For Binary Access Read As #2
 '   For I = 0 To LOF(2)
  '      Get #2, c, OneBit
   '     d = d & OneBit
    '    If OneBit = ";" Then
     '       Debug.Print d
      '  End If
       ' c = c + 3
    'Next I
Close #2
CD.ShowOpen
RunFile (CD.FileName)

End Sub

Private Sub Timer1_Timer()
Call cmdLPSTR_Click
Close #5


Timer1.Enabled = False
End Sub
