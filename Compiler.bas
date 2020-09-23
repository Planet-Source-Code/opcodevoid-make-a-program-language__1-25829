Attribute VB_Name = "Compiler"
Option Explicit
Public NumChild() As Long
Public NumButtons As Long
Public NumObjects As Long
Public NumLabels As Long
Public NumTextBoxes As Long
Public TriangleMemory(3) As String ' Specail memory
Public comment As Boolean 'Is comments ative
Public cObject() As String 'Object array
Public cId() As String 'Id Array
Dim I        'i'm lazy so most of my loops use i
Public NumMenus As Long 'Must keep track of mhow may menus we have
Public NumSubMenus As Long 'Sub menus
Public MenuClick As Long
Public CallBackId() As String 'For loop handling
Public CallBackA() As String 'For loop handling
Public NumCallBack As Long 'How many loops did you create
Public SkipBit As Boolean ' Skip spaces unless you see "
Public MessageGot As String
Public Flag As String 'For logic
Public CFlag As String
Public MessageId() As String 'Message Id array
Dim OneBit As String * 1
Dim Command As String
Dim NewObject As String
Dim ObjectId As String
Dim CurrentAddress As String 'Current address from file

Public Sub RunFile(FIle As String)
Dim b() As Byte
Dim BitTest
CurrentAddress = 1
Dim Running As Boolean
Close #1
    Open FIle For Binary Access Read As #1
        Do
            DoEvents
                Get #1, CurrentAddress, OneBit
                b = OneBit
                For BitTest = 0 To UBound(b)
                    If b(BitTest) = 13 Or b(BitTest) = 10 Then GoTo NextChar 'Removes lines
                        If b(BitTest) = 34 Then ' check for "
                            SkipBit = Not SkipBit 'Toggle if found
                            GoTo NextChar
                        End If
                        If b(BitTest) = 32 And SkipBit = False Then GoTo NextChar 'skip spaces only if told
                Next BitTest
                If OneBit = "/" And comment = False Then Exit Sub 'Fource exit
                If OneBit = "?" Then comment = True 'Fource Comment
                If OneBit = ";" And comment = False Then Synax (Command): Command = "": GoTo NextChar 'Got a command
                If comment = False Then Command = Command & OneBit
                If OneBit = "_" Then comment = False 'End comments
NextChar:
            CurrentAddress = CurrentAddress + 1
        Loop Until CurrentAddress = LOF(1)
    Close #1
    
End Sub


Public Sub Synax(tCommand As String)
'This is were most of everything is process here all commands go we determaine debug them all right here
'so this is basically the engine of the whole program

'Please don't ask me why a come up with names for varibles
Dim f As Long
Dim MidCommand As String
Dim Ob As Boolean
Dim Ot As Boolean
Dim FC As Boolean
Dim MF
Dim ObType As String
Dim TempA As String * 1
Dim CType As String
Dim Command2 As String 'The Command
Dim Output2 As String ' Output of the Command
Dim Ids As String
Dim FoundId As String
Dim RetVal As String
Dim Add1 As Long
Dim MenuId As String
Dim MenuH1 As String
Dim MenuCaptions As String
'simple loop that gets the command type
For f = 1 To Len(Command)
    If Mid(tCommand, f, 1) = ">" Then CType = ">": Ot = True: Exit For
    If Mid(tCommand, f, 1) = "." Then CType = ".": Ot = True: Exit For
    If Mid(tCommand, f, 1) = "*" Then CType = "*": Ot = True: Exit For
    If Mid(tCommand, f, 1) = "~" Then CType = "~": Ot = True: Exit For
    If Mid(tCommand, f, 1) = "[" Then CType = "[": Ot = True: Exit For
    If Mid(tCommand, f, 1) = "!" Then CType = "!": Ot = True
    MidCommand = MidCommand & Mid(tCommand, f, 1)
Next f


' ">" =Point command
' ". =" Object Command
'[ Logical Commands
'~ Compare Command
'NULL = Direct Command

Command2 = ""
        
Select Case CType

Case "!" 'Double Direct mostly backup
            For f = 2 To Len(tCommand)
            If Mid(tCommand, f, 1) = "=" Then Ot = False: GoTo NN
            If Mid(tCommand, f, 1) = ";" Then Exit For
            If Ot = False Then Output2 = Output2 & Mid(tCommand, f, 1)
            If Ot = True Then Command2 = Command2 & Mid(tCommand, f, 1)
NN:
        Next f
        SendCommand Command2, Output2
Case "[" 'Logic
    If Flag = "true" Then
        For f = 3 To Len(tCommand)
            If Mid(tCommand, f, 1) = "=" Then Ot = False: GoTo N1
            If Mid(tCommand, f, 1) = ";" Then Exit For
            If Ot = False Then Output2 = Output2 & Mid(tCommand, f, 1)
            If Ot = True Then Command2 = Command2 & Mid(tCommand, f, 1)
N1:
        Next f
        SendCommand Command2, Output2
    End If


 Case "~" 'Compare
    For f = 3 To Len(tCommand)
        If Mid(tCommand, f, 1) = "," Then Ot = False: GoTo TNextChar
        If Ot = True Then Command2 = Command2 & Mid(tCommand, f, 1)
        If Ot = False And Mid(tCommand, f, 1) = ";" Then Exit For
        If Ot = False Then Output2 = Output2 & Mid(tCommand, f, 1)
TNextChar:
    Next f
    'Output2 = Message
    'Command2 = "Id"
    MF = GetMessageId(Command2)

    If MF = "notfound" Then
        MsgBox "Can't Find Message Id: Address " & CurrentAddress: Exit Sub 'Can't find it
    Else
        MessageId(MF) = Output2 'Get message
        If MessageGot = MessageId(MF) Then 'Check it
            Flag = "true"
        End If
        
    End If
Case "*"
    For f = 1 To Len(tCommand)
        If Mid(tCommand, f, 1) = "&" Then Ot = False: Exit For
        If Ot = True Then Command2 = Command2 & Mid(tCommand, f, 1)
    Next f

    If Command2 = "loop" Then
        For f = 1 To Len(tCommand)
            If Mid(Command2, Len(tCommand) - f + 1, 1) = ";" Then
                If Mid(Command2, Len(tCommand) - f + 1, 1) = "*" Then Exit For
            Else
                Output2 = Output2 & Mid(tCommand, Len(tCommand) - f + 1, 1)
            End If
         Next f
    End If
        TempA = Output2
        ReDim Preserve CallBackA(NumCallBack)
        ReDim Preserve CallBackId(NumCallBack)
        CallBackA(NumCallBack) = CurrentAddress
        CallBackId(NumCallBack) = TempA
        NumCallBack = NumCallBack + 1
Case ">"
           f = 0
            Ob = True
            For f = 1 To Len(Command)
                If f >= Len(Command) Then
                Else
                    If Ot = False Then Exit For
                    If Ob = True Then MenuH1 = MenuH1 & Mid(tCommand, Len(tCommand) - f + 1, 1)
                    If Ob = False Then MenuId = MenuId & Mid(tCommand, Len(Command) - f + 1, 1)
                    If Ob = False And Mid(tCommand, Len(tCommand) - f + 1, 1) = ">" Then Ot = False
                    If Mid(tCommand, Len(tCommand) - f + 1, 1) = ">" Then Ob = False
                End If
            Next f
            RetVal = Clean(MenuH1)
            MenuH1 = RetVal
            RetVal = Clean(MenuId)
            MenuId = RetVal
            Ot = True
    Select Case Trim(MidCommand) 'Pointer Commands

        Case "create"
            f = 0
            Ob = True
            ReDim Preserve cObject(NumObjects)
            ReDim Preserve cId(NumObjects)
            ReDim Preserve MessageId(NumObjects)
            For f = 1 To Len(Command)
                If f >= Len(Command) Then
                Else
                    If Ot = False Then Exit For
                    If Ob = True Then cId(NumObjects) = cId(NumObjects) & Mid(tCommand, Len(tCommand) - f + 1, 1)
                    If Ob = False Then cObject(NumObjects) = cObject(NumObjects) & Mid(tCommand, Len(Command) - f + 1, 1)
                    If Ob = False And Mid(tCommand, Len(tCommand) - f + 1, 1) = ">" Then Ot = False
                    If Mid(tCommand, Len(tCommand) - f + 1, 1) = ">" Then Ob = False
                End If
            Next f
            RetVal = Clean(cObject(NumObjects)) 'Clean the strings because they or reverse (hehe)
            cObject(NumObjects) = RetVal
            RetVal = Clean(cId(NumObjects))
            cId(NumObjects) = RetVal
            fCreateObject cObject(NumObjects), cId(NumObjects) 'Call create object
            NumObjects = NumObjects + 1
        Case "menucaption"
          ProgramWindow.mnuNames(CLng(MenuId)).Caption = MenuH1
        Case "submenucaption"
            ProgramWindow.mnusubs(CInt(MenuId)).Caption = MenuH1
            'ProgramWindow.mnusubs(CInt(MenuId)).Visible = True
        Case "submenucolor"
            If MenuH1 = "green" Then MenuH1 = vbGreen
            If MenuH1 = "red" Then MenuH1 = vbRed
            If MenuH1 = "yellow" Then MenuH1 = vbYellow
            If MenuH1 = "blue" Then MenuH1 = vbBlue
            If MenuH1 = "white" Then MenuH1 = vbWhite
            If MenuH1 = "black" Then MenuH1 = vbBlack
            ProgramWindow.mnusubs(CInt(MenuId)).BackColor = MenuH1
        Case "new"
            Select Case MenuId
                Case "submenu"
                   NumChild(MenuH1) = NumChild(MenuH1) + 1
                    NumSubMenus = NumSubMenus + 1
                    Load ProgramWindow.mnusubs(NumSubMenus)
                    ProgramWindow.mnusubs(NumSubMenus).Top = ProgramWindow.mnusubs(NumSubMenus).Top * NumChild(MenuH1) + ProgramWindow.mnusubs(NumSubMenus).Height
                   
                 If CByte(MenuH1) <> 0 Then ProgramWindow.mnusubs(NumSubMenus).Left = CByte(MenuH1) * ProgramWindow.mnusubs(0).Width
                With ProgramWindow
                   'This is only use for test or debugging
                   ' .mnusubs(NumSubMenus).Visible = True
                   '.mnusubs(NumSubMenus).Caption = "Be kewl"
                End With
            Case "menu"
                With ProgramWindow
                     NumMenus = NumMenus + 1
                    ReDim Preserve NumChild(NumMenus)
                    NumChild(NumMenus) = -1
                    
                   
                    Load .mnuNames(NumMenus)
                    Dim HH 'Hold value 1
                    Dim HH2 'Hold Value 2 not used though hmm wonder why
                    
                    HH = NumMenus * .mnuNames(0).Width
                    .mnuNames(NumMenus).Left = .mnuNames(NumMenus).Left + HH
                End With
            End Select
        Case "add_reg_reg"
     
'Unfinish---------------------------------------------------------------------------------/////////////////////////////////////////////////
            Dim R1, R2
            
            R1 = MenuId 'First Command
            R2 = MenuH1 'Second Command
            
            
            Add_Reg_Reg MenuId, MenuH1
            
           ' If R1 = "dx" Then
             '   If R2 = "bx" Then
            '        dx = dx + bx
            '    End If
           ' End If
            
           ' If R1 = "fx" Then
               ' If R2 = "cx" Then
               ' Fx = Fx + cx
               ' End If
               ' If R2 = "bx" Then
                '    Fx = Fx + bx
               ' End If
                'If R2 = "cx" Then
                '    cx = cx + cx
               ' End If
               ' If R2 = "dx" Then
              '      cx = cx + dx
             '   End If
            'End If
        Case "setpixel"
            Dim TT As String, TT2 As String
            If MenuId = "fx" Then TT = Fx
            If MenuH1 = "dx" Then TT2 = dx
            ProgramWindow.PSet (TT, TT2), ax
        Case "goback"
            Output2 = Right(tCommand, 1)
            Command2 = GoBackTo(Output2)
            If Command2 = "notfound" Then
                If Debuging = True Then
                    'MsgBox "Error Address not found:" & "Error at Address:" & CurrentAddress
                    Close #66
                    Open App.Path & "/debug.gff" For Append As #66
                        Print #1, "Error Address not Found: Error At Address: " & CurrentAddress
                    Close #66
                End If
            Else
                CurrentAddress = CallBackA(Command2)
                I = CurrentAddress
            End If
      End Select 'End Potiner Comamnds
      
    Case "."
        Ob = True
        For f = 1 To Len(tCommand)
             If Ob = True Then Output2 = Output2 & Mid(tCommand, Len(tCommand) - f + 1, 1)
             If Mid(tCommand, Len(tCommand) - f + 1, 1) = "=" Then Ob = False
             If Ob = False Then Command2 = Command2 & Mid(tCommand, Len(tCommand) - f + 1, 1)
             If Ob = False And Mid(tCommand, Len(tCommand) - f + 1, 1) = "." Then Exit For
        Next f
         RetVal = TranslateMessage(Output2)
        Output2 = RetVal
        RetVal = TranslateMessage(Command2)
        Command2 = RetVal
        For f = 1 To Len(tCommand)
                If Mid(tCommand, f, 1) = "." Then Exit For
                Ids = Ids & Mid(tCommand, f, 1)
        Next f
        FoundId = FinDId1(Ids)
        If FoundId = "notfound" Then
            'MsgBox "Error Object Not Define: Address " & CurrentAddress
            If Debuging = True Then
                Close #66
                    Open App.Path & "/debug.gff" For Append As #66
                        Print #1, "Error Object not define: Error At Address: " & CurrentAddress
                    Close #66
            End If
        Else
        End If
        Select Case Command2
            
            Case ".caption"
               ObType = GetObjectType(CLng(FoundId))
               ChangeCaption ObType, CLng(FoundId), Output2
            Case ".left"
                ObType = GetObjectType(CLng(FoundId))
                ChangeLeft ObType, CLng(FoundId), Output2
            Case ".top"
                ObType = GetObjectType(CLng(FoundId))
                ChangeTop ObType, CLng(FoundId), Output2
            Case ".width"
                ObType = GetObjectType(CLng(FoundId))
                ChangeWidth ObType, CLng(FoundId), Output2
            End Select
    Case Else
        'Direct Commands
                
                SendCommand MidCommand, ""
End Select
End Sub


Public Function Clean(StringToClean As String) As String
Dim Bill As String
Dim Temp  As String
Dim Temp2 As String
Dim L
Temp = Replace(StringToClean, ">", "")
For L = 1 To Len(Temp)
    Temp2 = Temp2 & Mid(Temp, Len(Temp) - L + 1, 1)
Next L
Clean = Temp2
End Function
Public Sub fCreateObject(ObjectType As String, ObjectId As String)
Select Case ObjectType
    Case "commandbutton"
        With ProgramWindow
            If NumButtons = 0 Then
                .CBUTTON(0).Visible = True
                .CBUTTON(0).Caption = "Command Button 0"
               NumButtons = NumButtons + 1
            Else
                Load .CBUTTON(NumButtons)
                .CBUTTON(NumButtons).Visible = True
                .CBUTTON(NumButtons).Caption = "Command Button " & NumButtons
                NumButtons = NumButtons + 1
            End If
        End With
    Case "label"
        With ProgramWindow
            If NumLabels = 0 Then
                .CLABEL(0).Visible = True
                .CLABEL(0).Caption = "Label 0"
                NumLabels = NumLabels + 1
            Else
                Load .CLABEL(NumLabels)
                .CLABEL(NumLabels).Visible = True
                .CLABEL(NumLabels).Caption = "Label " & NumLabels
                NumLabels = NumLabels + 1
            End If
        End With
    Case "textbox"
        With ProgramWindow
            If NumTextBoxes = 0 Then
                .CTEXT(0).Visible = True
                .CTEXT(0).Text = "TextBox 0"
                NumTextBoxes = NumTextBoxes + 1
            Else
                Load .CTEXT(NumTextBoxes)
                .CTEXT(NumTextBoxes).Text = "TextBox " & NumTextBoxes
                .CTEXT(NumTextBoxes).Visible = True
                NumTextBoxes = NumTextBoxes + 1
            End If
        End With
    End Select
End Sub
Public Function TranslateMessage(mess As String) As String
Dim Temp, Temp2, Temp3, Temp4 As String
Dim H

H = InStr(mess, ".")
If H <> 0 Then
    Temp = Replace(mess, ".", "")
    Temp2 = Replace(mess, "=", "")
Else
    Temp = Replace(mess, "=", "")
    Temp2 = Temp
End If


For H = 0 To Len(Temp2)
    Temp3 = Temp3 & Mid(Temp2, Len(Temp2) - H + 1, 1)
Next H
 TranslateMessage = Temp3
End Function
Public Function FinDId1(IdName As String) As String
Dim H
Dim F1
For H = 0 To NumObjects
    If Trim(IdName) = cId(H) Then F1 = H:  FinDId1 = F1: Exit Function
Next H
f:

FinDId1 = "notfound"
End Function

Public Function GetObjectType(tObject As Long) As String
Dim H, F1
GetObjectType = cObject(tObject)
End Function
Public Function GoBackTo(What As String) As String
On Error GoTo na
Dim H
For H = 0 To NumCallBack
    If CallBackId(H) = What Then GoBackTo = H: Exit Function
    
Next H
na:
GoBackTo = "notfound"
End Function
Public Function GetMessageId(mess1 As String) As String
Dim H
For H = 0 To NumObjects
    If mess1 = cId(H) Then GetMessageId = H: Exit Function
Next H
End Function
Public Sub SendCommand(SCommand As String, Optional WhatToDO)
Dim Handle As String
Dim Command3 As String
Dim TempHandle  As String
Dim CheapF As String
Dim H
Dim b1() As Byte
Dim Fp
Dim FC As Boolean, SC As Boolean
Dim Reg1, CTo
Dim missing



If WhatToDO = "" Then CheapF = WhatToDO

Select Case SCommand 'Handle Direct Commands
    Case "showwindow"
        ProgramWindow.Show
    Case "msgbox"
        MsgBox WhatToDO
        MessageGot = ""
        Case Else
           For H = 1 To Len(SCommand)
                
                If Mid(SCommand, H, 1) = ";" Then Exit For
                If Mid(SCommand, H, 1) = "'" Then SC = True: GoTo NNN
                If SC = True And Mid(SCommand, H, 1) = "|" Then Fp = True: SC = False: GoTo NNN
                
                If SC = True Then Reg1 = Reg1 & Mid(SCommand, H, 1)
                If Fp = True Then CTo = CTo & Mid(SCommand, H, 1)
                
                If FC = True And Mid(SCommand, H, 1) = "," Then SC = True:  GoTo NNN
                If Mid(SCommand, H, 1) = "," And FC = False Then FC = True: GoTo NNN
                If FC = False Then Handle = Handle & Mid(SCommand, H, 1)
                If FC = True Then CheapF = CheapF & Mid(SCommand, H, 1)
NNN:
            Next H
            'Remove Spaces
            'Debug.Print SCommand
            b1 = Handle
            Handle = ""
                For H = 0 To UBound(b1)
                    Command3 = Hex(b1(H))
                    If Len(Command3) >= 2 Then
                        Handle = Handle & Chr(b1(H))
                End If
            Next H
            Select Case Trim(Handle)
                Case "inputbox"
                    mInputBox (CheapF)
                Case "msgbox"
                    mMsgbox (CheapF)
                Case "cmp"
                    mCompare Reg1, CTo
                 Case "qie"
                    If CFlag = "true" Then End
                Case "flushflags"  'flushflags
                    Flag = ""
                    CFlag = ""
                Case "showide"
                 StartEmu.Show
                Case "hideide"
                 StartEmu.Hide
                Case "idecaption"
                    StartEmu.Caption = CheapF
                Case "initmenu"
                    ProgramWindow.mnuNames(CInt(CheapF)).Visible = True
                Case "bx"
                    bx = CheapF
                Case "cx"
                    cx = CheapF
                Case "dx"
                    dx = CheapF
                Case "fx'"
                    Debug.Print "cheap "; CheapF
                    Fx = CheapF
                Case "px"
                    px = CheapF
                Case "sx"
                    sx = CheapF
                Case "menucompare"
                    If MenuClick = CheapF Then
                        Flag = "true"
                        MenuClick = 3000
                    End If
                Case "end"
                    End
                Case "hx"
                    hx = CheapF
                Case "cmp_window"
                    If CheapF = MessageGot Then Flag = "true"
                Case "ax"
                ax = CheapF
                If ax = "green" Then ax = vbGreen
                If ax = "red" Then ax = vbRed
                If ax = "yellow" Then ax = vbYellow
                If ax = "green" Then ax = vbGreen
                If ax = "black" Then ax = vbBlack
                If ax = "white" Then ax = vbWhite
                Case "delay"
                Dim d
                For d = 0 To CLng(CheapF)
                Next d
            End Select
  End Select
  
  
  
End Sub
Public Sub ChangeCaption(tObject As String, tObjectId As String, ToWhat As String)
Dim Temp3, Temp4 As Long



'2 labels created first so the object array will hold
'Tobject(0) = label
'Tobject(1) = label
'ObId = 0
'ObId = 1
'1 command buttons
'Tobject(2) = commandbutton
'ObId = 3
'1 more label
'Tobject(3) = label
'Obid = 4
'To DO = Find the third label Index which is 2

'---------------------------------'
'They send me Obid 4
'temp4 = ubound(clabel) = 2
'temp3 = numlabels: Numlabels = 3
'Temp4 - Temp4 - Obid: Now I equel -2
'temp4 = temp4 + NumLabels + 1 = 2
'--------------This is one way to get it
'But lets try to get the commandbutton using this way
'They send me Obid = 3
'temp4 = ubound(cbutton) = 0
'if temp4 = 0 then Change Property

If tObject = "commandbutton" Then
    Temp4 = NumButtons - 1
    If tObjectId = "0" Then Temp4 = 0
    If Temp4 = 0 Then ProgramWindow.CBUTTON(CInt(Temp4)).Caption = ToWhat: Exit Sub
    Temp3 = NumButtons 'lets say this is 2
    'we want to find 1
    Temp4 = Temp4 - tObjectId + 2 - 1 'lets say object id is 3
    If Temp4 < 0 Then Temp4 = 0
    ProgramWindow.CBUTTON(CInt(Temp4)).Caption = ToWhat
End If
If tObject = "label" Then
    Temp4 = NumLabels - 1
    If tObjectId = "0" Then Temp4 = 0
    If Temp4 = 0 Then ProgramWindow.CLABEL(CInt(Temp4)).Caption = ToWhat: Exit Sub
    Temp3 = NumLabels
    Temp4 = Temp4 - tObjectId + 2 - 1
    If Temp4 < 0 Then Temp4 = 0
    ProgramWindow.CLABEL(CInt(Temp4)).Caption = ToWhat
End If
If tObject = "textbox" Then
    Temp4 = NumTextBoxes - 1
    If tObjectId = "0" Then Temp4 = 0
    If Temp4 = 0 Then ProgramWindow.CTEXT(CInt(Temp4)).Text = ToWhat: Exit Sub
    Temp3 = NumTextBoxes
    Temp4 = Temp4 - tObjectId + 2 - 1
    If Temp4 < 0 Then Temp4 = 0
    ProgramWindow.CTEXT(CInt(Temp4)).Text = ToWhat
End If

End Sub
Public Sub ChangeLeft(tObject As String, tObjectId As String, ToWhat As String)
Dim Temp3, Temp4
If tObject = "commandbutton" Then
    Temp4 = NumButtons - 1
    If tObjectId = "0" Then Temp4 = 0
    If Temp4 = 0 Then ProgramWindow.CBUTTON(CInt(Temp4)).Left = CLng(ToWhat): Exit Sub
    Temp3 = NumButtons 'lets say this is 2
    'we want to find 1
    Temp4 = Temp4 - tObjectId + 2 - 1 'lets say object id is 3
    If Temp4 < 0 Then Temp4 = 0
    ProgramWindow.CBUTTON(CInt(Temp4)).Left = CLng(ToWhat)
End If
If tObject = "label" Then
     Temp4 = NumLabels - 1
    If tObjectId = "0" Then Temp4 = 0
    If Temp4 = 0 Then ProgramWindow.CLABEL(CInt(Temp4)).Left = ToWhat: Exit Sub
    Temp3 = NumLabels
    Temp4 = Temp4 - tObjectId + 2 - 1
    If Temp4 < 0 Then Temp4 = 0
    ProgramWindow.CLABEL(CInt(Temp4)).Left = CLng(ToWhat)
End If
If tObject = "textbox" Then
    Temp4 = NumTextBoxes - 1
    If tObjectId = "0" Then Temp4 = 0
    If Temp4 = 0 Then ProgramWindow.CTEXT(CInt(Temp4)).Left = ToWhat: Exit Sub
    Temp3 = NumTextBoxes
    Temp4 = Temp4 - tObjectId + 2 - 1
    If Temp4 < 0 Then Temp4 = 0
    ProgramWindow.CTEXT(CInt(Temp4)).Left = ToWhat
End If

End Sub
Public Sub ChangeTop(tObject As String, tObjectId As String, ToWhat As String)
Dim Temp3, Temp4
If tObject = "commandbutton" Then
       Temp4 = NumButtons - 1
    If tObjectId = "0" Then Temp4 = 0
    If Temp4 = 0 Then ProgramWindow.CBUTTON(CInt(Temp4)).Top = CLng(ToWhat): Exit Sub
    Temp3 = NumButtons 'lets say this is 2
    'we want to find 1
    Temp4 = Temp4 - tObjectId + 2 - 1 'lets say object id is 3
    If Temp4 < 0 Then Temp4 = 0
    ProgramWindow.CBUTTON(CInt(Temp4)).Top = CLng(ToWhat)
End If
If tObject = "label" Then
    Temp4 = NumLabels - 1
    If tObjectId = "0" Then Temp4 = 0
    If Temp4 = 0 Then ProgramWindow.CLABEL(CInt(Temp4)).Top = CLng(ToWhat): Exit Sub
    Temp3 = NumLabels
    Temp4 = Temp4 - tObjectId + 2 - 1
    If Temp4 < 0 Then Temp4 = 0
    ProgramWindow.CLABEL(CInt(Temp4)).Top = CLng(ToWhat)
End If
If tObject = "textbox" Then
    Temp4 = NumTextBoxes - 1
    If tObjectId = "0" Then Temp4 = 0
    If Temp4 = 0 Then ProgramWindow.CTEXT(CInt(Temp4)).Top = ToWhat: Exit Sub
    Temp3 = NumTextBoxes
    Temp4 = Temp4 - tObjectId + 2 - 1
    If Temp4 < 0 Then Temp4 = 0
    ProgramWindow.CTEXT(CInt(Temp4)).Top = ToWhat
End If

End Sub
Public Sub ChangeWidth(tObject As String, tObjectId As String, ToWhat As String)
Dim Temp3, Temp4
If tObject = "commandbutton" Then
    Temp4 = NumButtons - 1
    If tObjectId = "0" Then Temp4 = 0
    If Temp4 = 0 Then ProgramWindow.CBUTTON(CInt(Temp4)).Width = CLng(ToWhat): Exit Sub
    Temp3 = NumButtons 'lets say this is 2
    'we want to find 1
    Temp4 = Temp4 - tObjectId + 2 - 1 'lets say object id is 3
    If Temp4 < 0 Then Temp4 = 0
    ProgramWindow.CBUTTON(CInt(Temp4)).Width = CLng(ToWhat)
End If
If tObject = "label" Then
    Temp4 = NumLabels - 1
    If tObjectId = "0" Then Temp4 = 0
    If Temp4 = 0 Then ProgramWindow.CLABEL(CInt(Temp4)).Width = CLng(ToWhat): Exit Sub
    Temp3 = NumLabels
    Temp4 = Temp4 - tObjectId + 2 - 1
    If Temp4 < 0 Then Temp4 = 0
    ProgramWindow.CLABEL(CInt(Temp4)).Width = CLng(ToWhat)
End If
If tObject = "textbox" Then
    Temp4 = NumTextBoxes - 1
    If tObjectId = "0" Then Temp4 = 0
    If Temp4 = 0 Then ProgramWindow.CTEXT(CInt(Temp4)).Width = ToWhat: Exit Sub
    Temp3 = NumTextBoxes
    Temp4 = Temp4 - tObjectId + 2 - 1
    If Temp4 < 0 Then Temp4 = 0
    ProgramWindow.CTEXT(CInt(Temp4)).Width = ToWhat
End If
End Sub
Public Sub mInputBox(tText As String)
ax = InputBox(tText, "Flaming Ice")
If Watching = True Then
    Open App.Path & "/watch.pas" For Append As #88
        Print #88, "Value Change on Register AX to " & ax & " At  " & Date & ":" & Time
    Close #88
End If
        
End Sub
Public Sub mMsgbox(ttText As String)
MsgBox ttText
End Sub
Public Sub mCompare(tReg, wTo)


Select Case tReg
    Case "ax"
        If ax = wTo Then CFlag = "true"
    Case "bx"
        If bx = wTo Then CFlag = "true"
    Case "cx"
        If cx = wTo Then CFlag = "true"
    Case "dx"
        If dx = wTo Then CFlag = "true"
    Case "fx"
        If Fx = wTo Then CFlag = "true"
End Select
End Sub


Public Sub Add_Reg_Reg(Regone As String, RegTwo As String)

 Dim R1, R2
            Dim A, AA
            A = CLng(R1)
            AA = CLng(R2)
            R1 = Regone 'First Command
            R2 = RegTwo 'Second Command
            
            If R1 = "bx" Then A = bx
            If R1 = "cx" Then A = cx
            If R1 = "dx" Then A = dx
            If R1 = "fx" Then A = Fx
            
            
            If R2 = "bx" Then AA = bx
            If R2 = "cx" Then AA = cx
            If R2 = "dx" Then AA = dx
            If R2 = "fx" Then AA = Fx
            
            A = A + AA
            
            If R1 = "dx" Then
                dx = A
                If Watching = True Then
                    Open App.Path & "/watch.pas" For Append As #88
                        Print #88, "Value Change on Register DX to " & dx & " At  " & Date & ":" & Time
                    Close #88
                End If
            End If
            
            If R1 = "fx" Then
                Fx = A
                       If Watching = True Then
                    Open App.Path & "/watch.pas" For Append As #88
                        Print #88, "Value Change on Register FX to " & Fx & " At  " & Date & ":" & Time
                    Close #88
                End If
            End If
            
            
            
            If R1 = "bx" Then
                bx = A
                       If Watching = True Then
                    Open App.Path & "/watch.pas" For Append As #88
                        Print #88, "Value Change on Register BX to " & bx & " At  " & Date & ":" & Time
                    Close #88
                End If
            End If
            
            
            
            If R1 = "cx" Then
                cx = A
                       If Watching = True Then
                    Open App.Path & "/watch.pas" For Append As #88
                        Print #88, "Value Change on Register CX to " & cx & " At  " & Date & ":" & Time
                    Close #88
                End If
            End If
            
            
 
End Sub

