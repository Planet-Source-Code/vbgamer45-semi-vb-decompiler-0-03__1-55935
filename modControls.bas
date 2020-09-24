Attribute VB_Name = "modControls"
'#############################################
'modControls vbgamer45 2004
'##############################################
'Control Sepeartor Constatns
Public Const vbFormNewChildControl = 511 'FF01
Public Const vbFormExistingChildControl = 767 'FF02
Public Const vbFormChildControl = 1023 'FF03
Public Const vbFormEnd = 1279 'FF04
Public Const vbFormMenu = 1535 'FF05
'Control Header
Public Type ControlHeader
   'Length As Integer
    'unknown As Integer
    length As Long
    cId As Byte 'Used To link events
    cName As String
    un2 As Byte
    cType As Byte
End Type
Public Type ControlArrayHeader
    length As Integer
    un1 As Byte
   ' Length As Long
    ArrayFlag As Integer
    cId As Byte
    un2 As Byte
    cName As String
    un3 As Byte
    cType As Byte
End Type
Public Type ControlSize
    clientLeft As Integer
    un1 As Integer
    clientTop As Integer
    un2 As Integer
    clientWidth As Integer
    un3 As Integer
    clientHeight As Integer
    un4 As Integer
End Type
'Used in cType
Public Enum ControlType
    vbPictureBox = 0
    vbLabel = 1
    vbTextBox = 2
    vbFrame = 3
    vbCommandbutton = 4
    vbCheckbox = 5
    vbOptionbutton = 6
    vbComboBox = 7
    vbListbox = 8
    vbHscroll = 9
    vbVscroll = 10
    vbTimer = 11
    vbForm = 13
    vbDriveListbox = 16
    vbDirectoryListbox = 17
    vbFileListbox = 18
    vbMenu = 19
    vbMDIForm = 20
    vbShape = 22
    vbLine = 23
    vbImage = 24
    vbData = 37
    vbOLE = 38
    vbUserControl = 40
    vbPropertyPage = 41
    vbUserDocument = 42
End Enum



'External Controls
Private Type OcxListType
    strGuid As String
    strocxName As String
    strLibName As String
    strName As String
End Type

Global gOcxList() As OcxListType


Public Type FontType
    un1 As Byte
    un2 As Byte
    un3 As Byte
    Action As Byte
    Weight As Integer
    Size As Long
    FontLen As Byte
End Type

'Public Type tControlEventPointer
   'Const1 As Byte      ' 0x00
    'Flag1 As Long       ' 0x01
    'Const2 As Integer   ' 0x05 split up const2 into 2 ints
   ' Const3 As Integer   ' 0x07
    'Const4 As Byte      ' 0x09 changed from const3
   ' aEvent As Long      ' 0x0A
                        ' 0x0E &lt;-- Structure Size
'End Type


Public Type tControlEventLink

    Const1 As Integer        ' 0x00
    CompileType As Byte      ' 0x02 compileType According to Sarge[more info?]
    aEvent As Long           ' 0x03
                             ' 0x07 &lt;-- Structure Size
End Type


Public Type tControlEventPointer
    Const1 As Byte          ' 0x00
    Flag1 As Long           ' 0x01
    Const2 As Integer       ' 0x05
    EventLink As tControlEventLink ' 0x07
                            ' 0x0E &lt;-- Structure Size
End Type

Public Type LineSizeType
    X1 As Long
    op1 As Byte
    Y1 As Long
    op2 As Byte
    X2 As Long
    op3 As Byte
    Y2 As Long
End Type

Public Type DataFormatType
    LCID As Integer
End Type
'##################################
'Begin Subs for Processing Special opcodes and properties for common controls
'##################################
Sub ProccessForm(f As Variant, Opcode As Byte)

End Sub
Sub ProccessPictureBox(f As Variant, Opcode As Byte)

End Sub
Sub ProccessLabel(f As Variant, Opcode As Byte)

End Sub
Sub ProccessTextBox(f As Variant, Opcode As Byte)

End Sub
Sub ProccessFrame(f As Variant, Opcode As Byte)

End Sub
Sub ProccessCommandButton(f As Variant, Opcode As Byte)

End Sub
Sub ProccessCheckBox(f As Variant, Opcode As Byte)

End Sub
Sub ProccessOption(f As Variant, Opcode As Byte)

End Sub
Sub ProccessComboBox(f As Variant, Opcode As Byte)

End Sub
Sub ProccessListBox(f As Variant, Opcode As Byte)

End Sub
Sub ProccessHscroll(f As Variant, Opcode As Byte)

End Sub
Sub ProccessVscroll(f As Variant, Opcode As Byte)

End Sub
Sub ProccessTimer(f As Variant, Opcode As Byte)

End Sub
Sub ProccessDriveListBox(f As Variant, Opcode As Byte)

End Sub
Sub ProccessDirListBox(f As Variant, Opcode As Byte)

End Sub
Sub ProccessFileListBox(f As Variant, Opcode As Byte)

End Sub
Sub ProccessShape(f As Variant, Opcode As Byte)

End Sub
Sub ProccessLine(f As Variant, Opcode As Byte)

End Sub
Sub ProccessImage(f As Variant, Opcode As Byte)

End Sub
'##################################
'End Subs for Processing Special opcodes and properties for common controls
'##################################
Sub GetControlProperties(offset As Long)
'*****************************
'Purpose: Used for Form Editor
'*****************************
        Dim strCurrentForm As String
        Dim fPos As Long 'Holds current location in the file used for controlheader
        Dim cListIndex As Integer ' Used for COM
        Dim cControlHeader As ControlHeader
        Dim lForm As Integer
        Dim FRXAddress As Long
         Dim FileLen As Long
        
        'Unload Old Controls
        If frmMain.txtEditArray.UBound > 0 Then
        For i = 1 To frmMain.txtEditArray.UBound
            Unload frmMain.txtEditArray(i)
            Unload frmMain.lblArrayEdit(i)
        Next
        End If
        Set gVBFile = New clsFile
        Call gVBFile.Setup(SFilePath)
        f = gVBFile.FileNumber
        Seek f, offset + 1
        FRXAddress = 0

        fPos = Loc(f)
        Get #f, , cControlHeader
        frmMain.lblObjectName.Caption = "ObjectName: " & cControlHeader.cName
        Dim tliTypeInfo As TypeInfo 'Used for COM to find information about the properties of the control
        'Used to caculate how much father to go in the control
        'Select what type of control it is

        Select Case cControlHeader.cType
            Case vbPictureBox '= 0
                cListIndex = 22

            Case vbLabel '= 1
                cListIndex = 14
           
            Case vbTextBox ' = 2
                cListIndex = 27
           
            Case vbFrame '= 3
                cListIndex = 10
        
            Case vbCommandbutton '= 4
                cListIndex = 4
            Case vbCheckbox '= 5
                cListIndex = 1
              
            Case vbOptionbutton     ' = 6
                cListIndex = 21
              
            Case vbComboBox     ' = 7
                cListIndex = 3
            Case vbListbox     '= 8
                cListIndex = 17
            
            Case vbHscroll     '= 9
                cListIndex = 12
            
            Case vbVscroll     '= 10
                cListIndex = 32
             
            Case vbTimer     '= 11
                cListIndex = 28
          
            Case vbForm     '= 13
                cListIndex = 9
                strCurrentForm = cControlHeader.cName
               ' MsgBox cControlHeader.cName
                'Call modGlobals.LoadNewFormHolder(cControlHeader.cName)
                'Call AddText("Begin VB.Form " & cControlHeader.cName)
               ' gIdentSpaces = 1
            Case vbDriveListbox     '= 16
                cListIndex = 7
              '  Call AddText("Begin VB.DriveListbox " & cControlHeader.cName)
            Case vbDirectoryListbox     '= 17
                cListIndex = 6
               ' Call AddText("Begin VB.DirectoryListbox " & cControlHeader.cName)
            Case vbFileListbox     '= 18
                cListIndex = 8
               ' Call AddText("Begin VB.FileListBox " & cControlHeader.cName)
            Case vbMenu     '= 19
                cListIndex = 19
              '  Call AddText("Begin VB.Menu " & cControlHeader.cName)
            Case vbMDIForm     '= 20
                cListIndex = 18
                'Call AddText("Begin VB.MDIForm " & cControlHeader.cName)
            Case vbShape     '= 22
                cListIndex = 26
               ' Call AddText("Begin VB.Shape " & cControlHeader.cName)
            Case vbLine     '= 23
                cListIndex = 16
               ' Call AddText("Begin VB.Line " & cControlHeader.cName)
            Case vbImage     '= 24
                cListIndex = 12
              '  Call AddText("Begin VB.Image " & cControlHeader.cName)
            Case vbData     '= 37
                cListIndex = 5
            
            Case vbOLE     '= 38
                cListIndex = 20
              
            Case vbUserControl     '= 40
                cListIndex = 29
            
            Case vbPropertyPage     '= 41
                cListIndex = 24
                
            Case vbUserDocument     '= 42
                cListIndex = 30
              
            Case 255 'external control
             
                'Load the control view COM if its on the computer
                 
                Seek f, fPos + cControlHeader.length ' - 2
                
        End Select
        Set tliTypeInfo = tliTypeLibInfo.GetTypeInfo(Replace(Replace(frmMain.lstTypeInfos.List(cListIndex), "<", ""), ">", ""))
        'Use the ItemData in lstTypeInfos to set the SearchData for lstMembers
        tliTypeLibInfo.GetMembersDirect frmMain.lstTypeInfos.ItemData(cListIndex), frmMain.lstMembers.hwnd, , , True
        FileLen = Loc(f) - fPos
        FileLen = cControlHeader.length - FileLen
        
        Dim bCode As Byte 'holds gui opcode
        Dim varHold As Variant 'Holds the different data types
        Dim strHold As String 'holds the string
        Dim strReturnType As String 'holds the return type
        'Do While FileLen > 1
        Do While Loc(f) < (fPos + cControlHeader.length - 2)
       
         bCode = frmMain.GetOpcode(f) 'Get the guiopcode
        
         FileLen = FileLen - 1
         Dim g As Integer
         For g = 0 To frmMain.lstMembers.ListCount - 1
         
         
            'Control Postion opcode
            If bCode = 4 And cControlHeader.cType = vbDirectoryListbox Then
                Call frmMain.GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 4 And cControlHeader.cType = vbListbox Then
                Call frmMain.GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 4 And cControlHeader.cType = vbDriveListbox Then
                Call frmMain.GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 4 And cControlHeader.cType = vbFileListbox Then
                Call frmMain.GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 4 And cControlHeader.cType = vbTextBox Then
                Call frmMain.GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 4 And cControlHeader.cType = vbCommandbutton Then
                Call frmMain.GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 5 And cControlHeader.cType = vbPictureBox Then
                Call frmMain.GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 5 And cControlHeader.cType = vbComboBox Then

                Call frmMain.GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
   
            End If
            
            If bCode = 5 And cControlHeader.cType = vbOptionbutton Then
                Call frmMain.GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 5 And cControlHeader.cType = vbFrame Then
                Call frmMain.GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 5 And cControlHeader.cType = vbCheckbox Then
                Call frmMain.GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 5 And cControlHeader.cType = vbLabel Then
                Call frmMain.GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 7 And cControlHeader.cType = vbTimer Then
                AddText "Left=" & GetLong(f)
                FileLen = FileLen - 4
                Exit For
            End If
            If bCode = 8 And cControlHeader.cType = vbTimer Then
                AddText "Top=" & GetLong(f)
                FileLen = FileLen - 4
                Exit For
            End If
            If bCode = 2 And cControlHeader.cType = vbHscroll Then
                Call frmMain.GetControlSize(f)
                FileLen = FileLen - 8
                Exit For

            End If
            If bCode = 2 And cControlHeader.cType = vbVscroll Then
                Call frmMain.GetControlSize(f)
                FileLen = FileLen - 8
                Exit For

            End If
            If bCode = 37 And cControlHeader.cType = vbLabel Then
                'Font
                Call frmMain.GetFontProperty(f)
                Exit For
            End If
            If bCode = 64 And cControlHeader.cType = vbForm Then
                'Font
                Call frmMain.GetFontProperty(f)
                Exit For
            End If
            If cControlHeader.cType = vbCommandbutton And bCode = 29 Then
                Call frmMain.GetFontProperty(f)
                Exit For
            End If
            If cControlHeader.cType = vbPictureBox And bCode = 57 Then
                Call frmMain.GetFontProperty(f)
                Exit For
            End If
            If cControlHeader.cType = vbTextBox And bCode = 46 Then
                Call frmMain.GetFontProperty(f)
                Exit For
            End If
            If cControlHeader.cType = vbFrame And bCode = 4 Then
                
                'Call AddText("ForeColor=" & GetLong(f))
                Exit For
            End If
            
            If ReturnGuiOpcode(frmMain.lstTypeInfos.ItemData(cListIndex), tliInvokeKinds, frmMain.lstMembers.List(g)) = bCode Then
              Dim strExtraInfo As String
              Dim strHelp As String

                'Com Hack Check
                strReturnType = Trim(ReturnDataType(frmMain.lstTypeInfos.ItemData(cListIndex), tliInvokeKinds, frmMain.lstMembers.List(g)))
                strHelp = modCOM.ReturnHelpString(frmMain.lstTypeInfos.ItemData(cListIndex), tliInvokeKinds, frmMain.lstMembers.List(g))
                For k = 0 To UBound(gComFix)
                    If frmMain.lstTypeInfos.List(cListIndex) = gComFix(k).ObjectName And frmMain.lstMembers.List(g) = gComFix(k).PropertyName Then
                        strReturnType = gComFix(k).NewType
                        
                        Exit For
                    End If
                Next
                
                If InStr(1, strReturnType, "Byte") Then
                    varHold = GetByte2(f)
                    'Call AddText(lstMembers.List(g) & " = " & varHold & strExtraInfo)
                    Call modGlobals.AddPropertyToTheList(frmMain.lstMembers.List(g), varHold, "Byte", Loc(f) - 1, strHelp)
                    FileLen = FileLen - 1
                    Exit For
                End If
                If InStr(1, strReturnType, "Boolean") Then
                    varHold = GetBoolean(f)
                    
                    If varHold = True Then
                        varHold = False
                        'Call AddText(lstMembers.List(g) & " = " & -1 & strExtraInfo)
                        Call modGlobals.AddPropertyToTheList(frmMain.lstMembers.List(g), varHold, "Boolean", Loc(f) - 2, strHelp)
                    Else
                        varHold = True
                        'Call AddText(lstMembers.List(g) & " = " & 0 & strExtraInfo)
                        Call modGlobals.AddPropertyToTheList(frmMain.lstMembers.List(g), varHold, "Boolean", Loc(f) - 2, strHelp)
                    End If
                    Seek f, Loc(f)
                    FileLen = FileLen - 2
                    Exit For
                End If
                If InStr(1, strReturnType, "Integer") Then
                    varHold = gVBFile.GetInteger(Loc(f))
                    'varHold = GetInteger(f)
                    'Call AddText(lstMembers.List(g) & " = " & varHold & strExtraInfo)
                    Call modGlobals.AddPropertyToTheList(frmMain.lstMembers.List(g), varHold, "Integer", Loc(f) - 2, strHelp)
                    FileLen = FileLen - 2
                    Exit For
                End If
                If InStr(1, strReturnType, "Long") Then
                    varHold = GetLong(f)
                    'Call AddText(lstMembers.List(g) & " = " & varHold & strExtraInfo)
                    Call modGlobals.AddPropertyToTheList(frmMain.lstMembers.List(g), varHold, "Long", Loc(f) - 4, strHelp)
                    FileLen = FileLen - 4
                    Exit For
                End If
                
                If InStr(1, strReturnType, "Single") Then
                    varHold = GetSingle(f)
                    'Call AddText(lstMembers.List(g) & " = " & varHold & strExtraInfo)
                    Call modGlobals.AddPropertyToTheList(frmMain.lstMembers.List(g), varHold, "Single", Loc(f) - 4, strHelp)
                    FileLen = FileLen - 4
                    Exit For
                End If

                If InStr(1, strReturnType, "String") Then
                    
                    ''Seek f, Loc(f) + 3
                    ''strHold = GetUntilNull(f)
                    strHold = GetAllString(f)
                    'Call AddText(lstMembers.List(g) & " = " & Chr(34) & strHold & Chr(34) & strExtraInfo)
                    Call modGlobals.AddPropertyToTheList(frmMain.lstMembers.List(g), strHold, "String", Loc(f), strHelp)
                    FileLen = FileLen - Len(strHold) - 3
                    Exit For
                End If
                If InStr(1, strReturnType, "stdole.Picture") Then
                    
                    varHold = GetLong(f)
                   
                    If varHold <> -1 Then
                
                        If cControlHeader.cName <> strCurrentForm Then
                            Call frmMain.GetStdPicture(f, varHold, strCurrentForm & "." & cControlHeader.cName, strCurrentForm, 0)
                        Else
                            Call frmMain.GetStdPicture(f, varHold, cControlHeader.cName, strCurrentForm, 0)
                        End If
                        
                        
                        'Call AddText(lstMembers.List(g) & "=" & Chr(34) & strCurrentForm & ".frx" & Chr(34) & ":" & frxAddress & strExtraInfo)
                        Seek f, Loc(f)
                        FileLen = FileLen - varHold + 1 ' - 18

                    Else
                        FileLen = FileLen - 4
                    End If
                    Exit For
                End If
               
               
                Exit For

            End If
            
            'Get height width top left
            If bCode = 53 Then
            '53 is the size opcode for form's
            Dim objectSize As ControlSize
                Get f, , objectSize
                FileLen = FileLen - 16
                
                If cControlHeader.cType = vbForm Then
                   ' Call AddText("ClientLeft = " & objectSize.clientLeft)
                   ' Call AddText("ClientTop = " & objectSize.clientTop)
                   ' Call AddText("ClientWidth = " & objectSize.clientWidth)
                   ' Call AddText("ClientHeight = " & objectSize.clientHeight)
                End If
                
                Exit For
            End If
         Next
        Loop
        
        
        'Get the seperator type for the end of the control

        cControlEnd = GetInteger(f)
 
        FileLen = FileLen - 2
Close f
'##########################################
'End of Form/Control Properties Loop
'##########################################
End Sub
