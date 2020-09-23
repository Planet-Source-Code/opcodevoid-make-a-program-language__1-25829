VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form ObjectWindow 
   ClientHeight    =   4980
   ClientLeft      =   -255
   ClientTop       =   645
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   7080
   Begin VB.CommandButton CmdRemove 
      Caption         =   "Remove Slected list"
      Height          =   255
      Left            =   5280
      TabIndex        =   10
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear List"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "Stop Watch"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Timer timWatch 
      Interval        =   500
      Left            =   6360
      Top             =   4440
   End
   Begin VB.ListBox LstWatch 
      Height          =   1035
      ItemData        =   "ObjectWindow.frx":0000
      Left            =   0
      List            =   "ObjectWindow.frx":0002
      TabIndex        =   7
      Top             =   5520
      Width           =   7095
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   810
      Left            =   0
      TabIndex        =   1
      Top             =   -120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1429
      _CBWidth        =   6975
      _CBHeight       =   810
      _Version        =   "6.0.8450"
      MinHeight1      =   360
      Width1          =   2880
      NewRow1         =   0   'False
      MinHeight2      =   360
      Width2          =   1440
      NewRow2         =   -1  'True
      MinHeight3      =   360
      Width3          =   1440
      NewRow3         =   0   'False
      Begin VB.CommandButton CmdWatch 
         Caption         =   "Watch"
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton CmdDebug 
         Caption         =   "Debug"
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton CmdRun 
         Caption         =   "Run"
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton CmdLoad 
         Caption         =   "Load"
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7435
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"ObjectWindow.frx":0004
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   6480
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   0
      Y1              =   5040
      Y2              =   5520
   End
   Begin VB.Line Line5 
      X1              =   6960
      X2              =   6960
      Y1              =   5040
      Y2              =   5520
   End
   Begin VB.Line Line4 
      X1              =   3600
      X2              =   3600
      Y1              =   5040
      Y2              =   4920
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   7080
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   7080
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line1 
      X1              =   7080
      X2              =   5400
      Y1              =   5040
      Y2              =   5040
   End
End
Attribute VB_Name = "ObjectWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewString As String
Dim PrevString As String
Dim GiveString As String
Dim L As Long
Dim PL As Long

Private Sub CmdClear_Click()
LstWatch.Clear
End Sub

Private Sub CmdDebug_Click()
On Error GoTo Pscerror
Open App.Path & "/start.gff" For Output As #1
    Print #1, Debug_File & vbCrLf & p.pPath
Close #1
Shell App.Path & "/flamming ice.exe", vbNormalFocus
'Erase but not kill
'Do
  '  DoEvents
 '   For i = 0 To 30000
'Next i
'Loop

Open App.Path & "/start.gff" For Output As #2
 Print #2, ""
Close #2

Exit Sub
Pscerror:
    MsgBox "PSC want let me summit ""EXE"" so to get this feature wokring you must create a file called flammingice.exe in the same path your running this program from"
End Sub

Private Sub CmdLoad_Click()
CD.ShowOpen
RTB.LoadFile CD.FileName
p.pPath = CD.FileName
End Sub

Private Sub CmdRemove_Click()
LstWatch.RemoveItem LstWatch.ListIndex

End Sub

Private Sub CmdRun_Click()
On Error GoTo Pscerror
Open App.Path & "/start.gff" For Output As #1
    Print #1, Run_File & vbCrLf & p.pPath
Close #1
Shell App.Path & "/flamming ice.exe", vbNormalFocus
'Erase but not kill
'Do
  '  DoEvents
 '   For i = 0 To 30000
'Next i
'Loop

'Open App.Path & "/start.gff" For Output As #2
 'Print #2, ""
'Close #2

Exit Sub
Pscerror:
    MsgBox "PSC want let me summit ""EXE"" so to get this feature wokring you must create a file called flammingice.exe in the same path your running this program from"
End Sub

Private Sub CmdSave_Click()
CD.ShowSave
Open CD.FileName For Output As #12
    Print #12, RTB.Text
Close #12
End Sub

Private Sub CmdStop_Click()
WatchFile = False
Me.Height = Non_Watch

End Sub

Private Sub CmdWatch_Click()
On Error GoTo Pscerror
WatchFile = True
Open App.Path & "/start.gff" For Output As #1
    Print #1, Watch_File & vbCrLf & p.pPath
Close #1
Me.Height = WatchTop
Shell App.Path & "/flamming ice.exe", vbNormalFocus
'Erase but not kill
'Do
  '  DoEvents
 '   For i = 0 To 30000
'Next i
'Loop

'Open App.Path & "/start.gff" For Output As #2
 'Print #2, ""
'Close #2

Exit Sub
Pscerror:
    MsgBox "PSC want let me summit ""EXE"" so to get this feature wokring you must create a file called flammingice.exe in the same path your running this program from"
    WatchFile = False
End Sub

Private Sub timWatch_Timer()
L = 0

If WatchFile = True Then
    Close #45
    Open App.Path & "/watch.pas" For Input As #45
    Do While Not EOF(45)
        Line Input #45, NewString
            L = L + 1
    Loop
End If

If L = PL Then
    Else
    LstWatch.AddItem NewString
End If
Close #45
PL = L
End Sub
