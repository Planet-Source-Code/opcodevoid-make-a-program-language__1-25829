VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMenuedit 
   BackColor       =   &H00808080&
   Caption         =   "Menu Editor"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   10005
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Cd 
      Left            =   9840
      Top             =   -240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   $"frmMenuedit.frx":0000
      Height          =   1695
      Left            =   0
      TabIndex        =   2
      Top             =   3600
      Width           =   9975
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   8400
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   7200
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Flamming Ice"
         Height          =   255
         Left            =   8880
         TabIndex        =   21
         Top             =   720
         Width           =   975
      End
      Begin VB.Line Line3 
         X1              =   8760
         X2              =   8760
         Y1              =   0
         Y2              =   1680
      End
      Begin VB.Label Label14 
         BackColor       =   &H0000FF00&
         Caption         =   "                   Generate Code"
         Height          =   375
         Left            =   5760
         TabIndex        =   20
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "                New Sub Menu"
         Height          =   375
         Left            =   3000
         TabIndex        =   19
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         Caption         =   "                    New Menu"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label lblSubId 
         Caption         =   "0"
         Height          =   255
         Left            =   4320
         TabIndex        =   17
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblSubCaption 
         Caption         =   "Sub menu 1"
         Height          =   255
         Left            =   4200
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Menu Id"
         Height          =   255
         Left            =   3000
         TabIndex        =   15
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Menu Caption"
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblsubColor 
         Caption         =   "Test"
         Height          =   255
         Left            =   4200
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Menu Color"
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.Line Line2 
         X1              =   2880
         X2              =   2880
         Y1              =   120
         Y2              =   1680
      End
      Begin VB.Label Label5 
         Caption         =   "Generate were"
         Height          =   255
         Left            =   5760
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   5640
         X2              =   5640
         Y1              =   120
         Y2              =   1680
      End
      Begin VB.Label lblmnuId 
         Caption         =   "0"
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Menu Id"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblmnuCaption 
         Caption         =   "Menu 1"
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Menu Caption"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblColor 
         Caption         =   "TEST"
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Menu Color"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label mnusubs 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sub menu 0"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label mnunames 
      Caption         =   "Menu 0"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmMenuedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Menu Editor uses the same way to make menus as flamming ice does
'Also for testing
Option Explicit
Dim SelSubMenu As Integer
Dim NumMenu As Integer
Dim NumSubMenus As Integer
Dim I




Sub MakeSubMenu(mnuParent As Byte)
            Dim Add1
                    NumChild(mnuParent) = NumChild(mnuParent) + 1
                       NumSubMenus = NumSubMenus + 1
                    Load mnusubs(NumSubMenus)
                    
                    
                    mnusubs(NumSubMenus).Top = mnusubs(NumSubMenus).Top * NumChild(mnuParent) + mnusubs(NumSubMenus).Height
    
                 If CByte(mnuParent) <> 0 Then mnusubs(NumSubMenus).Left = CByte(mnuParent) * mnusubs(0).Width
            mnusubs(NumSubMenus).Visible = True
            mnusubs(NumSubMenus).Caption = "Sub Menu " & NumSubMenus
End Sub

Private Sub Command1_Click()
CD.ShowOpen: Text1.Text = CD.FileName
End Sub

Private Sub Label12_Click()
MakeMenu
End Sub

Private Sub Label13_Click()
SelParent = 255
PopupSubMenu
'The cpu is faster than the user so we must pause it until the user chooses
Do While SelParent = 255
    DoEvents
Loop

MakeSubMenu (SelParent)
End Sub

Sub MakeMenu()

                    NumMenu = NumMenu + 1
                     ReDim Preserve NumChild(NumMenu)
                     NumChild(NumMenu) = -1
                     
                    Load mnunames(NumMenu)
                    Dim HH 'Hold value 1
                    Dim HH2 'Hold Value 2 not used though hmm wonder why
                    HH = NumMenu * mnunames(0).Width
                    mnunames(NumMenu).Left = mnunames(NumMenu).Left + HH
                    'Letting the program know you created a menu
                    RegisterParent (NumMenu) 'Create a special little spot in the array for your menu
                    ParentName(NumMenu) = "Menu " & NumMenu
                    'Letting the user know you created a menu
                    mnunames(NumMenu).Visible = True
                    mnunames(NumMenu).Caption = ParentName(NumMenu)
End Sub
Sub RegisterParent(PNumber As Byte)
ReDim Preserve ParentName(PNumber)
End Sub

Sub GetMenuInfo(WhichOne As Byte)
'Get Menus info for the user
lblmnuCaption = mnunames(WhichOne).Caption
lblColor.BackColor = mnunames(WhichOne).BackColor
lblmnuId.Caption = WhichOne
End Sub
Sub GetSubMenuInfo(Which As Byte)
lblSubCaption.Caption = mnunames(Which).Caption
lblsubColor.BackColor -mnunames(Which).BackColor
lblSubId.Caption = Which
End Sub

Private Sub Label14_Click()
On Error GoTo Crap
'Note this dosen't work yet
Dim MnuPrint As String
Dim MnuInit As String
Dim SubPrint As String
Open Text1.Text For Output As #1
    For I = 0 To NumMenu
        MnuPrint = MnuPrint & "new>menu>" & I & ";" & vbCrLf
        MnuInit = MnuInit & "initmenu," & I & ";" & vbCrLf
    Next I
    For I = 0 To NumSubMenus
        SubPrint = SubPrint & "new>submenu>" & I & ";" & vbCrLf
    Next I
    Print #1, " ?Created by Flamming Ide Menu Editor _ " & vbCrLf & vbCrLf & vbCrLf & " ?Parent Menus _ " & vbCrLf & MnuPrint & vbCrLf & vbCrLf & vbCrLf & "? Child Menus _ " & vbCrLf & SubPrint & vbCrLf & vbCrLf & vbCrLf & "?Initalize Menus _" & vbCrLf & MnuInit
Close #1
MsgBox "Generated Successfuly"
Exit Sub
Crap:
MsgBox "Error Occured while generating"
End Sub

Private Sub mnunames_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SelMenu = Index
If Button = 1 Then GetMenuInfo (Index)
If Button = 2 Then ShowMenuPSpy (SelMenu)

End Sub

Private Sub mnusubs_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SelSub = Index
If Button = 1 Then GetSubMenuInfo (Index)
If Button = 2 Then ShowMenuSpy (Index)
End Sub


