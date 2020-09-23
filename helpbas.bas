Attribute VB_Name = "helpbas"
'Contains functions use for HelpBox
'your free to use the help box in your apps you don't even have to give me credit
'if you want
'aslo contains function for other windows
Public Const HBoxH = 3195 'Help Box Height

Public Sub PopUpBox(Text As String, ReadFrom)
Dim Somthing 'Can't think of anything
Dim CurrentLine
Dim FullLine
'tight way to pop up a window
HelpBox.Height = 0
HelpBox.Visible = True

Do While HelpBox.Height < HBoxH
    DoEvents
        For Somthing = 0 To 1500000 ' i have a 700mhz cpu so its really hard to slow down you might have to adjust this
        Next
        HelpBox.Height = HelpBox.Height + 250
Loop

If ReadFrom <> "" Then   'if paremeter is gotten
    
    Open ReadFrom For Input As #1
        Do While Not EOF(1)
            Line Input #1, CurrentLine
            'just for fun i put some script in
            If Trim(CurrentLine) = "/skipline" Then FullLine = FullLine & vbCrLf Else FullLine = FullLine & CurrentLine
        Loop
    Close #1
    HelpBox.lblHelp.Text = FullLine
Else
    HelpBox.lblHelp.Text = Text
End If
End Sub

Sub PopupSubMenu()
InputSubMenu.Show
For i = 0 To UBound(ParentName)
    InputSubMenu.LstMenus.AddItem ParentName(i)
Next i
End Sub

Sub ShowMenuSpy(LeechFrom As Byte)
Dim LF
LF = LeechFrom ' to lazy to write leechfrom

Dim BkColor 'BackGround color
Dim MnuCaption

'assign varibles
With frmMenuedit
    BkColor = .mnusubs(LF).BackColor
    MnuCaption = .mnusubs(LF).Caption
End With

MenuSpy.Show
'Set everything up
With MenuSpy
    .Label3.BackColor = BkColor
    .txtCaption.Text = MnuCaption
End With

End Sub

Sub ShowMenuPSpy(LeechFromP As Byte)
Dim LFP
LFP = LeechFrom 'still lazy

Dim BkC 'Back ground color
Dim MnuC


'assign varibles
With frmMenuedit
    BkC = .mnunames(LFP).BackColor
    MnuC = .mnunames(LFP).Caption
End With

'set everything up
With menuPSpy
    .Show
    .Label4.BackColor = BkC
    .txtCaption = MnuC
End With
End Sub
