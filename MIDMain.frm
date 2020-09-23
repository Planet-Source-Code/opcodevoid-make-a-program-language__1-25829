VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm MIDMain 
   BackColor       =   &H8000000C&
   ClientHeight    =   8595
   ClientLeft      =   1500
   ClientTop       =   825
   ClientWidth     =   11400
   LinkTopic       =   "MDIForm1"
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   11400
      Top             =   960
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1429
      _CBWidth        =   11400
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
      Begin VB.CommandButton CmdMenu 
         Caption         =   "Menu Edit"
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton NewCode 
         Caption         =   "Code Gen"
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton CmdTextBox 
         Caption         =   "TextBox"
         Height          =   255
         Left            =   3120
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton CmdLabel 
         Caption         =   "Label"
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton CmdButton 
         Caption         =   "Button"
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "New"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton CmdLoad 
         Caption         =   "Load"
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   11400
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "MIDMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdButton_Click()
PopUpBox vbNullString, App.Path & "/idedata/commandbutton.txt"


End Sub

Private Sub CmdExit_Click()
End
End Sub

Private Sub CmdLabel_Click()
PopUpBox vbNullString, App.Path & "/idedata/label.txt"
End Sub

Private Sub CmdLoad_Click()
Call CmdNew_Click

End Sub

Private Sub CmdMenu_Click()
frmMenuedit.Show
End Sub

Private Sub CmdNew_Click()
CD.ShowSave
p.pPath = CD.FileName
FlagTrue = True
ObjectWindow.Show
ObjectWindow.RTB.LoadFile CD.FileName
End Sub

Private Sub CmdSave_Click()

End Sub

Private Sub CmdTextBox_Click()
PopUpBox vbNullString, App.Path & "/idedata/textbox.txt"
End Sub

Private Sub MDIForm_Load()
ReDim Preserve ParentName(0)
ParentName(0) = "Menu 0"


End Sub

Private Sub Timer1_Timer()
If FlagTrue = False Then
    CmdButton.Enabled = False
    CmdLabel.Enabled = False
    CmdTextBox.Enabled = False
    CmdMenu.Enabled = False
Else
    CmdButton.Enabled = True
    CmdLabel.Enabled = True
    CmdTextBox.Enabled = True
    CmdMenu.Enabled = True
End If
End Sub
