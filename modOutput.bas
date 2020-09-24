Attribute VB_Name = "modOutput"
'#############################################
'modOutput vbgamer45 2004
'##############################################
'Types and Variables for exe changes
Private Type byteChangeType
    bByte As Byte
    offset As Long
End Type
Private Type BooleanChangeType
    bBool As Boolean
    offset As Long
End Type
Private Type IntegerChangeType
    iInt As Integer
    offset As Long
End Type
Private Type LongChangeType
    lLong As Long
    offset As Long
End Type
Private Type SingleChangeType
    sSingle As Single
    offset As Long
End Type
Private Type StringChangeType
    sString As String
    offset As Long
End Type
Global ByteChange() As byteChangeType
Global BooleanChange() As BooleanChangeType
Global IntegerChange() As IntegerChangeType
Global LongChange() As LongChangeType
Global SingleChange() As SingleChangeType
Global StringChange() As StringChangeType

Sub DumpVBExeInfo(Filename As String, FileTitle As String)
'*****************************
'Purpose: Prints a report about the Exe that was decompiled
'*****************************
Dim i As Integer

    Open Filename For Output As #1
        Print #1, "----------------------------------"
        Print #1, FileTitle
        Print #1, "Output made by Semi VB Decompiler by vbgamer45"
        Print #1, "----------------------------------"
        Print #1, "VB Exe Info"
        Print #1, "----------------------------------"
        Print #1, "VBStartOffset " & AppData.VBStartOffset
        Print #1, "FormCount= " & gVBHeader.FormCount
        Print #1, "ModuleCount= " & AppData.AppModuleCount
        Print #1, "CompileType= " & AppData.CompileType
        Print #1, "----------------------------------"
        Print #1, "VB Header Infomation"
        Print #1, "----------------------------------"
        Print #1, "ProjectTitle= " & ProjectTitle
        Print #1, "ProjectName= " & ProjectName
        Print #1, "ExeName= " & ProjectExename
        Print #1, "HelpFile= " & HelpFile
        If gVBHeader.aSubMain <> 0 Then
            Print #1, "SubMain Address= " & gVBHeader.aSubMain + 1 - OptHeader.ImageBase
        End If
        Print #1, "ExternalComponentCount= " & gVBHeader.ExternalComponentCount
        Print #1, "----------------------------------"
        Print #1, "Object List"
        Print #1, "----------------------------------"
        For i = 0 To UBound(gObjectNameArray)
            Print #1, gObjectNameArray(i)
        Next i
     
        Print #1, "----------------------------------"
        Print #1, "External Ocx List"
        Print #1, "----------------------------------"
        If UBound(gOcxList) > 0 Then
            For i = 0 To UBound(gOcxList) - 1
                Print #1, gOcxList(i).strocxName
                Print #1, gOcxList(i).strName
                Print #1, gOcxList(i).strLibName
                Print #1, gOcxList(i).strGuid
                Print #1, ""
            Next i
        End If
        Print #1, "----------------------------------"
        Print #1, "Api List"
        Print #1, "----------------------------------"
        
        For i = 0 To UBound(gApiList) - 1
           Print #1, "Declare " & gApiList(i).strFunctionName & " Lib " & Chr(34) & gApiList(i).strLibraryName & Chr(34)
        Next i
        Print #1, "----------------------------------"
        Print #1, "Controls Guids"
        Print #1, "----------------------------------"
        Print #1, "Parent Form, Control Name, GUID"
        For i = 0 To UBound(gControlNameArray)
        
            Print #1, gControlNameArray(i).strParentForm & " , " & gControlNameArray(i).strControlName & " , " & gControlNameArray(i).strGuid
        Next i
  
    Close #1
    
    
End Sub
Sub WriteVBP(Filename As String)
'*****************************
'Purpose: Writes the visual basic project file
'*****************************
    Open Filename For Output As #3
    
        Print #3, "Type=Exe"
        'If gVBHeader.ExternalComponentCount > 1 Then
           ' Print #3, "Reference="
           ' Print #3, "Object="
       ' End If
        
        For i = 0 To UBound(gObject)
            If gObject(i).ObjectType = 98435 Then
            'Form LooP
                Print #3, "Form=" & gObjectNameArray(i) & ".frm"
            End If
            If gObject(i).ObjectType = 98305 Then
            'Module Loop
                Print #3, "Module=" & gObjectNameArray(i) & "; " & gObjectNameArray(i) & ".bas"
            End If
            If gObject(i).ObjectType = 1146883 Then
            'Class Loop
                Print #3, "Class=" & gObjectNameArray(i) & "; " & gObjectNameArray(i) & ".cls"
            End If
             If gObject(i).ObjectType = 1941507 Or gObject(i).ObjectType = 1943555 Then
            'User Control
                Print #3, "UserControl=" & gObjectNameArray(i) & "; " & gObjectNameArray(i) & ".ctl"
            End If
            
        Next
        If gVBHeader.ExternalComponentCount > 0 Then
            For i = 0 To UBound(gOcxList) - 1
                If gOcxList(i).strGuid <> "" Then
                    Print #3, "Object={" & gOcxList(i).strGuid & "}#1.0#0; " & gOcxList(i).strocxName
                End If
            Next i
        
        End If
        
        'Print #3, "IconForm=" & Chr(34) & DATAHERE & Chr(34)
        If gVBHeader.aSubMain = 0 Then
            Print #3, "Startup=" & Chr(34) & gObjectNameArray(0) & Chr(34)
        End If
        Print #3, "Description=" & Chr(34) & ProjectDescription & Chr(34)
        Print #3, "HelpFile=" & Chr(34) & HelpFile & Chr(34)
        Print #3, "Name=" & Chr(34) & ProjectName & Chr(34)
        Print #3, "Title=" & Chr(34) & ProjectTitle & Chr(34)
        Print #3, "ExeName32=" & Chr(34) & ProjectExename & Chr(34)
        Print #3, "VersionCompanyName=" & Chr(34) & gFileInfo.CompanyName & Chr(34)
    Close #3
    
End Sub
Sub ShowVBPFile()
'*****************************
'Purpose: To Show the VBP File in the textbox
'*****************************
    frmMain.txtCode.Text = ""
    frmMain.txtCode.Text = frmMain.txtCode.Text & "Type=Exe" & vbCrLf
        For i = 0 To UBound(gObject)
            If gObject(i).ObjectType = 98435 Then
            'Form LooP
                frmMain.txtCode.Text = frmMain.txtCode.Text & "Form=" & gObjectNameArray(i) & ".frm" & vbCrLf
            End If
            If gObject(i).ObjectType = 98305 Then
            'Module Loop
                frmMain.txtCode.Text = frmMain.txtCode.Text & "Module=" & gObjectNameArray(i) & "; " & gObjectNameArray(i) & ".bas" & vbCrLf
            End If
            If gObject(i).ObjectType = 1146883 Then
            'Class Loop
                frmMain.txtCode.Text = frmMain.txtCode.Text & "Class=" & gObjectNameArray(i) & "; " & gObjectNameArray(i) & ".cls" & vbCrLf
            End If
             If gObject(i).ObjectType = 1941507 Or gObject(i).ObjectType = 1943555 Then
            'User Control
                frmMain.txtCode.Text = frmMain.txtCode.Text & "UserControl=" & gObjectNameArray(i) & "; " & gObjectNameArray(i) & ".ctl" & vbCrLf
            End If
            
        Next
        'External Components
        If gVBHeader.ExternalComponentCount > 0 Then
            For i = 0 To UBound(gOcxList) - 1
                If gOcxList(i).strGuid <> "" Then
                    frmMain.txtCode.Text = frmMain.txtCode.Text & "Object={" & gOcxList(i).strGuid & "}#1.0#0; " & gOcxList(i).strocxName & vbCrLf
                End If
            Next i
        End If
        
    If gVBHeader.aSubMain = 0 Then
        frmMain.txtCode.Text = frmMain.txtCode.Text & "Startup=" & Chr(34) & gObjectNameArray(0) & Chr(34) & vbCrLf
    Else
        frmMain.txtCode.Text = frmMain.txtCode.Text & "Startup=" & Chr(34) & "Sub Main" & Chr(34) & vbCrLf
    
    End If
    frmMain.txtCode.Text = frmMain.txtCode.Text & "Description=" & Chr(34) & ProjectDescription & Chr(34) & vbCrLf
    frmMain.txtCode.Text = frmMain.txtCode.Text & "HelpFile=" & Chr(34) & HelpFile & Chr(34) & vbCrLf
    frmMain.txtCode.Text = frmMain.txtCode.Text & "Name=" & Chr(34) & ProjectName & Chr(34) & vbCrLf
    frmMain.txtCode.Text = frmMain.txtCode.Text & "Title=" & Chr(34) & ProjectTitle & Chr(34) & vbCrLf
    frmMain.txtCode.Text = frmMain.txtCode.Text & "ExeName32=" & Chr(34) & ProjectExename & Chr(34) & vbCrLf
    frmMain.txtCode.Text = frmMain.txtCode.Text & "VersionCompanyName=" & Chr(34) & gFileInfo.CompanyName & Chr(34) & vbCrLf

End Sub
Sub WriteForms(FilePath As String)
'*****************************
'Purpose: To export the forms to a .frm file
'*****************************
    For i = 0 To frmMain.txtFinal.UBound
        If frmMain.txtFinal(i).Tag <> "" Then
        Open FilePath & frmMain.txtFinal(i).Tag & ".frm" For Output As #4
            Print #4, "VERSION 5.00"
            'Begin Object References
            
            'Begin Form
            Print #4, frmMain.txtFinal(i).Text
            'Print the procedures
            Print #4, "'Generated by Semi VB Decompiler -vbgamer45"
            For nApi = 0 To UBound(gProcedureList)
                If UCase(frmMain.txtFinal(i).Tag) = UCase(gProcedureList(nApi).strParent) Then
                    Print #4, "Sub " & gProcedureList(nApi).strProcedureName & "()"
                    Print #4, "End Sub"
                End If
            Next

        Close #4
        End If
    Next
    
End Sub
Sub WriteFormFrx(FilePath As String, FormName As String)
'*****************************
'Purpose: Write the forms graphic files (.frx)
'*****************************
    Dim pFrxHeader As FRXITEMHDR
    Dim i As Integer
    On Error Resume Next
    Kill FilePath & "\" & FormName & ".frx"
    
    fFile = FreeFile
    Open FilePath & "\" & FormName & ".frx" For Binary Access Write Lock Write As fFile
    For i = 0 To UBound(FrxPreview)
        If FrxPreview(i).ParentForm = FormName Then
            pFrxHeader.dwSizeImage = FrxPreview(i).length
            pFrxHeader.dwSizeImageEx = FrxPreview(i).length + 8
            pFrxHeader.dwKey = &H746C
            Put fFile, , pFrxHeader
            PicFile = FreeFile
            Dim Buffer() As Byte
            Dim bEndByte As Integer
            ReDim Buffer(pFrxHeader.dwSizeImage)
            bEndByte = 2573
            Open App.path & "\dump\" & SFile & "\" & FrxPreview(i).strPath For Binary Access Read Lock Read As PicFile
                Get PicFile, , Buffer
                Put fFile, , Buffer
                Seek fFile, Loc(fFile)
                Put fFile, , bEndByte
            Close PicFile
            
        End If
    Next
    Close fFile

End Sub
Sub WriteModules(Filename As String, ObjectName As String)
'*****************************
'Purpose: To export the modules to a .bas file
'*****************************
    Open Filename For Output As #5
        Print #5, "Attribute VB_Name = " & Chr(34) & ObjectName & Chr(34)
    Close #5
End Sub
Sub WriteClasses(Filename As String, ObjectName As String)
'*****************************
'Purpose: To export the classes to a .cls file
'*****************************
    Open Filename For Output As #6
        Print #6, "VERSION 1.0 CLASS"
        Print #6, "Begin"
        Print #6, "  MultiUse = -1  'True"
        Print #6, "  Persistable = 0  'NotPersistable"
        Print #6, "  DataBindingBehavior = 0  'vbNone"
        Print #6, "  DataSourceBehavior = 0   'vbNone"
        Print #6, "  MTSTransactionMode = 0   'NotAnMTSObject"
        Print #6, "End"
        Print #6, "Attribute VB_Name = " & Chr(34) & ObjectName & Chr(34)
        Print #6, "Attribute VB_GlobalNameSpace = False"
        Print #6, "Attribute VB_Creatable = True"
        Print #6, "Attribute VB_PredeclaredId = False"
        Print #6, "Attribute VB_Exposed = False"
        Print #6, "Attribute VB_Ext_KEY = " & Chr(34) & "SavedWithClassBuilder6" & Chr(34) & "," & Chr(34) & "Yes" & Chr(34)
        Print #6, "Attribute VB_Ext_KEY = " & Chr(34) & "Top_Level" & Chr(34) & " ," & Chr(34) & "No" & Chr(34)
        
        Print #6, "'Generated by Semi VB Decompiler -vbgamer45"
        For nApi = 0 To UBound(gProcedureList)
            If UCase(ObjectName) = UCase(gProcedureList(nApi).strParent) Then
                Print #6, "Sub " & gProcedureList(nApi).strProcedureName & "()"
                Print #6, "End Sub"
            End If
        Next
        
    
    Close #6
End Sub
Sub WriteUserControls(Filename As String)
'*****************************
'Purpose: To export the controls to a .ctl file
'*****************************
Dim DATAHERE As String
    Open Filename For Output As #7
    
    Close #7
End Sub
