VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MenuSpy 
   BorderStyle     =   0  'None
   ClientHeight    =   1230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   ScaleHeight     =   1230
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   3600
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtCaption 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FF00&
      Caption         =   "         Ok"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Test"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Menu Back Color"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Menu Caption"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "MenuSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label3_Click()
CD.ShowColor
Label3.BackColor = CD.Color
End Sub

Private Sub Label4_Click()
With frmMenuedit
   .mnusubs(SelSub).Caption = txtCaption.Text
   .mnusubs(SelSub).BackColor = Label3.BackColor
End With


Me.Hide
End Sub
