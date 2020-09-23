VERSION 5.00
Begin VB.Form InputSubMenu 
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   7200
   ClientTop       =   1920
   ClientWidth     =   2685
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   2685
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox LstMenus 
      Height          =   3180
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "menu's Parent "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "InputSubMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LstMenus_Click()
SelParent = LstMenus.ListIndex
Me.Hide
LstMenus.Clear
End Sub
