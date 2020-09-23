VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form menuPSpy 
   BorderStyle     =   0  'None
   ClientHeight    =   1200
   ClientLeft      =   7680
   ClientTop       =   4320
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   ScaleHeight     =   1200
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CD 
      Left            =   3840
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtCaption 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Test"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Back Color"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Menu Captions"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "          OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "menuPSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
With frmMenuedit
    .mnunames(SelMenu).Caption = txtCaption.Text
    .mnunames(SelMenu).BackColor = Label4.BackColor
End With
Me.Hide
End Sub

Private Sub Label4_Click()
CD.ShowColor: Label4.BackColor = CD.Color 'Get color and show it
End Sub
