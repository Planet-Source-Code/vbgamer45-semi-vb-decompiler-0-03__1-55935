Attribute VB_Name = "modNative"
'#############################################
'modNative vbgamer45 2004
'##############################################
'Module for Processing Native Code
Private Type API_VBDEF
    Rva As Long
    Ordinal As Long
    uName As String
    uDescr As String
End Type
Private exeVB6_APIDEF() As API_VBDEF
Sub Decode(Filename As String)
'*****************************
'Purpose: To Get the procdures of a Native Exe and produce a report
'*****************************

Dim FileNum As Integer
    Open App.path & "\dump\" & SFile & "\NativeOut.txt" For Output As #2
        Print #2, "Semi VB Decompiler vbgamer45"
        Print #2, "Native Output for : " & Filename
        Print #2, "---------------------------------"
       
'        Open SFilePath For Binary Access Read Lock Read As #FileNum
            'Goto The address of Native Code
           ' Seek #FileNum, gProjectInfo.aNativeCode + 1 - OptHeader.ImageBase
            'Begin Code recovery
            'call modDecodeNative.RecoverCode(loc(f))
            
      '  Close #FileNum
       
       ' Do
           ' c = 0
            'For a = 0 To ProcCnt - 1
                'If ProcList(a) <> 0 Then
                   ' Print #2, DecompileProc(ProcList(a))
                   ' ProcList(a) = 0
                 '   c = 1
              '  End If
           ' Next
       ' Loop While c
    Close #2
    
End Sub

Sub VBFunction_Description_Init(ByVal fRes As String)
'*****************************
'Purpose: To load the Msvbvm60.dll api list from a file
'*****************************
Dim lfp As Integer, i As Long
Dim sAdr As String, sOrd As String, sName As String, sDef As String
lfp = FreeFile
Erase exeVB6_APIDEF()

    Open fRes For Input Access Read As #lfp
        i = 0
        Do
        i = i + 1
            Input #lfp, sAdr, sOrd, sName, sDef
            If LCase$(sAdr) <> "eof" Then
                ReDim Preserve exeVB6_APIDEF(1 To i)
                exeVB6_APIDEF(i).Rva = Val("&H" & sAdr)
                exeVB6_APIDEF(i).Ordinal = CLng(sOrd)
                exeVB6_APIDEF(i).uName = sName
                exeVB6_APIDEF(i).uDescr = sDef
            Else
                Exit Do
            End If
        Loop Until EOF(1)
    
    Close #lfp

End Sub
Public Function VBFunction_Description(ByVal inOrdinal As Long, ByVal inAPIname As String, ByRef outRName As String) As String
'*****************************
'Purpose: To return the description of a function
'*****************************
Dim i As Long


If inOrdinal > 0 And inAPIname = "" Then
    'by ordinal :
    For i = 1 To UBound(exeVB6_APIDEF)
        If exeVB6_APIDEF(i).Ordinal = inOrdinal Then
            VBFunction_Description = exeVB6_APIDEF(i).uDescr
            outRName = exeVB6_APIDEF(i).uName
            Exit Function
        End If
    Next i

Else
    'by name:
   
    For i = 1 To UBound(exeVB6_APIDEF)
        If exeVB6_APIDEF(i).uName = inAPIname Then
            VBFunction_Description = exeVB6_APIDEF(i).uDescr
            
            Exit Function
        End If
    Next i
End If

VBFunction_Description = "Error API incorrect or not present in msvbvm60.dll"

End Function

Sub LoadMsvbvm60DllExports()
'Used for VB6
'Note there are multiple versions of Msvbvm60.dll right now i know of about 4

End Sub
Sub LoadMsvbvm50DllExports()
'Used for VB5

End Sub
Sub LoadOLEAUT32DllExports()

End Sub
Sub LoadExternalDll()
'Used for api calls

End Sub

Sub NativeProcessProcedures(f As Variant)
    'Goto Begining of the code
    Seek f, gProjectInfo.aStartOfCode + 1 - OptHeader.ImageBase
End Sub

Function NameFromOrdinal(libname As String, Ordinal As Integer, TypeLibInfo As TypeLibInfo)
'*****************************
'Purpose: To get a function name from an Ordinal
'*****************************
On Error Resume Next
    Dim TypeInfo As TypeInfo
    Dim Member As MemberInfo

    Dim sDLLName As String
    Dim sEntryName As String
    Dim iOrdinal As Integer

    For Each TypeInfo In TypeLibInfo.TypeInfos
        For Each Member In TypeInfo.Members
            Member.GetDllEntry sDLLName, sEntryName, iOrdinal
            If Ordinal = iOrdinal Then
                Set NameFromOrdinal = Member
                Exit Function
            End If
        Next
    Next
End Function

