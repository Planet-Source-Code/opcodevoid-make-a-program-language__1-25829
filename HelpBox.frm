VERSION 5.00
Begin VB.Form HelpBox 
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   10680
   ClientTop       =   7875
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox lblHelp 
      BackColor       =   &H80000000&
      Height          =   2775
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   4695
   End
   Begin VB.CommandButton CmdHide 
      Caption         =   "Hide"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2760
      Width           =   4695
   End
End
Attribute VB_Name = "HelpBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdHide_Click()
Me.Hide
Me.Height = 0
End Sub
