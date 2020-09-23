VERSION 5.00
Begin VB.Form ProgramWindow 
   ClientHeight    =   5700
   ClientLeft      =   2520
   ClientTop       =   1245
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   6180
   Begin VB.TextBox CTEXT 
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton CBUTTON 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   5280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label mnusubs 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sub Menu1"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label mnuNames 
      Caption         =   "Menu 1"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label CLABEL 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   2055
   End
End
Attribute VB_Name = "ProgramWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBUTTON_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    MessageGot = BS_RightClicked
    MessageId(Index) = Index
Else
    MessageGot = BS_CLICKED
    MessageId(Index) = Index
End If
End Sub

Private Sub CLABEL_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    MessageGot = BS_RightClicked
    MessageId(Index) = Index
Else
    MessageGot = BS_CLICKED
    MessageId(Index) = Index
End If
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
Source.Move X, Y
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
MessageGot = KeyCode
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HideMenus
End Sub

Private Sub mnuNames_Click(Index As Integer)
Dim k
For k = 0 To NumSubMenus
    If mnusubs(k).Left = mnuNames(Index).Left Then
        mnusubs(k).Visible = True
    End If
Next k
End Sub

Sub HideMenus()
Dim L
For L = 0 To NumSubMenus
mnusubs(L).Visible = False
Next L
End Sub

Private Sub mnusubs_Click(Index As Integer)
MenuClick = Index
End Sub
