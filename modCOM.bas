Attribute VB_Name = "modCOM"
'*****************************
'modCom.bas
'Purpose to Retrive the members and variable types of a control
'*****************************
Global tliTypeLibInfo As TypeLibInfo
Public Function GetSearchType(ByVal SearchData As Long) As TliSearchTypes
    'This helper function adapted from Microsoft documentation
    If SearchData And &H80000000 Then
        GetSearchType = ((SearchData And &H7FFFFFFF) \ &H1000000 And &H7F&) Or &H80
    Else
        GetSearchType = SearchData \ &H1000000 And &HFF&
    End If
End Function
Public Function PrototypeMember(ByVal SearchData As Long, _
    ByVal InvokeKinds As InvokeKinds, _
    Optional ByVal MemberName As String) As String
    'This helper function adapted from Microsoft documentation
    On Error GoTo exitFunction
    Dim tliParameterInfo As ParameterInfo
    Dim bFirstParameter As Boolean
    Dim bIsConstant As Boolean
    Dim bByVal As Boolean
    Dim strReturn As String
    Dim ConstVal As Variant
    Dim strTypeName As String
    Dim intVarTypeCur As Integer
    Dim bDefault As Boolean
    Dim bOptional As Boolean
    Dim bParamArray As Boolean
    Dim tliTypeInfo As TypeInfo
    Dim tliResolvedTypeInfo As TypeInfo
    Dim tliTypeKinds As TypeKinds
  
    With tliTypeLibInfo
        
        'First, determine the type of member we're dealing with
        bIsConstant = GetSearchType(SearchData) And tliStConstants
        With .GetMemberInfo(SearchData, InvokeKinds, , MemberName)
            Debug.Print "MemberID: 0x" & Hex(.MemberId - &H10000)
            If bIsConstant Then
                strReturn = "Const "
            ElseIf InvokeKinds = INVOKE_FUNC Or InvokeKinds = INVOKE_EVENTFUNC Then
                Select Case .ReturnType.VarType
                    Case VT_VOID, VT_HRESULT
                        strReturn = "Sub "
                    Case Else
                        strReturn = "Function "
                End Select
            Else
                strReturn = "Property "
            End If
        
            'Now add the name of the member
            strReturn = strReturn & .Name
        
            'Process the member's paramters
            With .Parameters
                If .count Then
                    strReturn = strReturn & " ("
                    bFirstParameter = True
                    bParamArray = .OptionalCount = -1
                    For Each tliParameterInfo In .Me
                        If Not bFirstParameter Then
                            strReturn = strReturn & ", "
                        End If
                        bFirstParameter = False
                        bDefault = tliParameterInfo.Default
                        bOptional = bDefault Or tliParameterInfo.Optional
                        If bOptional Then
                            If bParamArray Then
                                'This will be the only optional parameter
                                strReturn = strReturn & "[ParamArray "
                            Else
                                strReturn = strReturn & "["
                            End If
                        End If
                    
                        With tliParameterInfo.VarTypeInfo
                            Set tliTypeInfo = Nothing
                            Set tliResolvedTypeInfo = Nothing
                            tliTypeKinds = TKIND_MAX
                            intVarTypeCur = .VarType
                            If (intVarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
                                On Error Resume Next
                                Set tliTypeInfo = .TypeInfo
                                If Not tliTypeInfo Is Nothing Then
                                    Set tliResolvedTypeInfo = tliTypeInfo
                                    tliTypeKinds = tliResolvedTypeInfo.TypeKind
                                    Do While tliTypeKinds = TKIND_ALIAS
                                        tliTypeKinds = TKIND_MAX
                                        Set tliResolvedTypeInfo = tliResolvedTypeInfo.ResolvedType
                                        If Err Then
                                            Err.Clear
                                        Else
                                            tliTypeKinds = tliResolvedTypeInfo.TypeKind
                                        End If
                                    Loop
                                End If
                            
                                'Determine whether parameters are ByVal or ByRef
                                Select Case tliTypeKinds
                                    Case TKIND_INTERFACE, TKIND_COCLASS, TKIND_DISPATCH
                                        bByVal = .PointerLevel = 1
                                    Case TKIND_RECORD
                                        'Records not passed ByVal in VB
                                        bByVal = False
                                    Case Else
                                        bByVal = .PointerLevel = 0
                                End Select
                            
                                'Indicate ByVal
                                If bByVal Then
                                    strReturn = strReturn & "ByVal "
                                End If
                            
                                'Display the parameter name
                                strReturn = strReturn & tliParameterInfo.Name
                            
                                If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                    strReturn = strReturn & "()"
                                End If
                                
                                If tliTypeInfo Is Nothing Then 'Information not available
                                    strReturn = strReturn & " As ?"
                                Else
                                    If .IsExternalType Then
                                        strReturn = strReturn & " As " & .TypeLibInfoExternal.Name & "." & tliTypeInfo.Name
                                    Else
                                        strReturn = strReturn & " As " & tliTypeInfo.Name
                                    End If
                                End If
                            
                                'Reset error handling
                                On Error GoTo 0
                            Else
                                If .PointerLevel = 0 Then
                                    strReturn = strReturn & "ByVal "
                                End If
                                    
                                strReturn = strReturn & tliParameterInfo.Name
                                If intVarTypeCur <> vbVariant Then
                                    strTypeName = TypeName(.TypedVariant)
                                    If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                        strReturn = strReturn & "() As " & Left$(strTypeName, Len(strTypeName) - 2)
                                    Else
                                        strReturn = strReturn & " As " & strTypeName
                                    End If
                                End If
                            End If
                                
                            If bOptional Then
                                If bDefault Then
                                    strReturn = strReturn & ProduceDefaultValue(tliParameterInfo.DefaultValue, tliResolvedTypeInfo)
                                    'strReturn = strReturn & " = " & tliParameterInfo.DefaultValue
                                End If
                                strReturn = strReturn & "]"
                            End If
                        End With
                    Next
                    strReturn = strReturn & ")"
                End If
            End With
        
            If bIsConstant Then
                ConstVal = .Value
                strReturn = strReturn & " = " & ConstVal
                Select Case VarType(ConstVal)
                    Case vbInteger, vbLong
                        If ConstVal < 0 Or ConstVal > 15 Then
                            strReturn = strReturn & " (&H" & Hex$(ConstVal) & ")"
                        End If
                End Select
            Else
                With .ReturnType
                    intVarTypeCur = .VarType
                    If intVarTypeCur = 0 Or (intVarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
                        On Error Resume Next
                        If Not .TypeInfo Is Nothing Then
                            If Err Then 'Information not available
                                strReturn = strReturn & " As ?"
                            Else
                                If .IsExternalType Then
                                    strReturn = strReturn & " As " & .TypeLibInfoExternal.Name & "." & .TypeInfo.Name
                                Else
                                    strReturn = strReturn & " As " & .TypeInfo.Name
                                End If
                            End If
                        End If
                        
                        If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                            strReturn = strReturn & "()"
                        End If
                        On Error GoTo 0
                    Else
                        Select Case intVarTypeCur
                            Case VT_VARIANT, VT_VOID, VT_HRESULT
                            Case Else
                                strTypeName = TypeName(.TypedVariant)
                                If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                    strReturn = strReturn & "() As " & Left$(strTypeName, Len(strTypeName) - 2)
                                Else
                                    strReturn = strReturn & " As " & strTypeName
                                End If
                        End Select
                    End If
                End With
            End If
            
            PrototypeMember = strReturn & vbCrLf
            lblMemberOf = "Member of " & tliTypeLibInfo.Name & "." & tliTypeLibInfo.GetTypeInfo(SearchData And &HFFFF&).Name
            lblHelpText = .HelpString
        End With
    End With
exitFunction:
End Function
Public Function getNameFromMemberInfo(mi As MemberInfo) As String
Dim sOutput As String, strTypeName As String, ConstVal As String
Dim lSearchData As Long
Dim bIsConstant As Boolean, bDefault As Boolean, bFirstParameter As Boolean
Dim bParamArray As Boolean, bOptional As Boolean, bByVal As Boolean
Dim tliParameterInfo As ParameterInfo
Dim tliTypeInfo As TypeInfo, tliResolvedTypeInfo As TypeInfo
Dim tliTypeKinds As TypeKinds
Dim intVarTypeCur As Integer
            With mi
                '.VTableOffset
                sOutput = sOutput & "0x" & Hex(.VTableOffset) & ":"
                bIsConstant = GetSearchType(lSearchData) And tliStConstants
                If bIsConstant Then
                    sOutput = sOutput & "Const "
                ElseIf mi.InvokeKind = INVOKE_FUNC Or mi.InvokeKind = INVOKE_EVENTFUNC Then
                    Select Case .ReturnType.VarType
                        Case VT_VOID, VT_HRESULT
                            sOutput = sOutput & "Sub "
                        Case Else
                            sOutput = sOutput & "Function "
                    End Select
                Else
                    sOutput = sOutput & "Property "
                End If
                sOutput = sOutput & .Name
                With .Parameters
                    If .count Then
                        sOutput = sOutput & " ("
                        bFirstParameter = True
                        bParamArray = .OptionalCount = -1
                        For Each tliParameterInfo In .Me
                            If Not bFirstParameter Then
                                sOutput = sOutput & ", "
                            End If
                            bFirstParameter = False
                            bDefault = tliParameterInfo.Default
                            bOptional = bDefault Or tliParameterInfo.Optional
                            If bOptional Then
                                If bParamArray Then
                                    'This will be the only optional parameter
                                    sOutput = sOutput & "[ParamArray "
                                Else
                                    sOutput = sOutput & "["
                                End If
                            End If
                        
                            With tliParameterInfo.VarTypeInfo
                                Set tliTypeInfo = Nothing
                                Set tliResolvedTypeInfo = Nothing
                                tliTypeKinds = TKIND_MAX
                                intVarTypeCur = .VarType
                                If (intVarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
                                    On Error Resume Next
                                    Set tliTypeInfo = .TypeInfo
                                    If Not tliTypeInfo Is Nothing Then
                                        Set tliResolvedTypeInfo = tliTypeInfo
                                        tliTypeKinds = tliResolvedTypeInfo.TypeKind
                                        Do While tliTypeKinds = TKIND_ALIAS
                                            tliTypeKinds = TKIND_MAX
                                            Set tliResolvedTypeInfo = tliResolvedTypeInfo.ResolvedType
                                            If Err Then
                                                Err.Clear
                                            Else
                                                tliTypeKinds = tliResolvedTypeInfo.TypeKind
                                            End If
                                        Loop
                                    End If
                                
                                    'Determine whether parameters are ByVal or ByRef
                                    Select Case tliTypeKinds
                                        Case TKIND_INTERFACE, TKIND_COCLASS, TKIND_DISPATCH
                                            bByVal = .PointerLevel = 1
                                        Case TKIND_RECORD
                                            'Records not passed ByVal in VB
                                            bByVal = False
                                        Case Else
                                            bByVal = .PointerLevel = 0
                                    End Select
                                
                                    'Indicate ByVal
                                    If bByVal Then
                                        sOutput = sOutput & "ByVal "
                                    End If
                                
                                    'Display the parameter name
                                    sOutput = sOutput & tliParameterInfo.Name
                                
                                    If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                        sOutput = sOutput & "()"
                                    End If
                                    
                                    If tliTypeInfo Is Nothing Then 'Information not available
                                        sOutput = sOutput & " As ?"
                                    Else
                                        If .IsExternalType Then
                                            sOutput = sOutput & " As " & .TypeLibInfoExternal.Name & "." & tliTypeInfo.Name
                                        Else
                                            sOutput = sOutput & " As " & tliTypeInfo.Name
                                        End If
                                    End If
                                
                                    'Reset error handling
                                    On Error GoTo 0
                                Else
                                    If .PointerLevel = 0 Then
                                        sOutput = sOutput & "ByVal "
                                    End If
                                        
                                    sOutput = sOutput & tliParameterInfo.Name
                                    If intVarTypeCur <> vbVariant Then
                                        strTypeName = TypeName(.TypedVariant)
                                        If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                            sOutput = sOutput & "() As " & Left$(strTypeName, Len(strTypeName) - 2)
                                        Else
                                            sOutput = sOutput & " As " & strTypeName
                                        End If
                                    End If
                                End If
                                    
                                If bOptional Then
                                    If bDefault Then
                                        sOutput = sOutput & ProduceDefaultValue(tliParameterInfo.DefaultValue, tliResolvedTypeInfo)
                                        'sOutput = sOutput & " = " & tliParameterInfo.DefaultValue
                                    End If
                                    sOutput = sOutput & "]"
                                End If
                            End With
                        Next
                        sOutput = sOutput & ")"
                    End If
                End With
                'return type
                If bIsConstant Then
                    ConstVal = .Value
                    sOutput = sOutput & " = " & ConstVal
                    Select Case VarType(ConstVal)
                        Case vbInteger, vbLong
                            If ConstVal < 0 Or ConstVal > 15 Then
                                sOutput = sOutput & " (&H" & Hex$(ConstVal) & ")"
                            End If
                    End Select
                Else
                    With .ReturnType
                        intVarTypeCur = .VarType
                        If intVarTypeCur = 0 Or (intVarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
                            On Error Resume Next
                            If Not .TypeInfo Is Nothing Then
                                If Err Then 'Information not available
                                    sOutput = sOutput & " As ?"
                                Else
                                    If .IsExternalType Then
                                        sOutput = sOutput & " As " & .TypeLibInfoExternal.Name & "." & .TypeInfo.Name
                                    Else
                                        sOutput = sOutput & " As " & .TypeInfo.Name
                                    End If
                                End If
                            End If
                            
                            If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                sOutput = sOutput & "()"
                            End If
                            On Error GoTo 0
                        Else
                            Select Case intVarTypeCur
                                Case VT_VARIANT, VT_VOID, VT_HRESULT
                                Case Else
                                    strTypeName = TypeName(.TypedVariant)
                                    If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                        sOutput = sOutput & "() As " & Left$(strTypeName, Len(strTypeName) - 2)
                                    Else
                                        sOutput = sOutput & " As " & strTypeName
                                    End If
                            End Select
                        End If
                    End With
                End If
            End With
        getNameFromMemberInfo = sOutput
End Function
Public Function ProduceDefaultValue(DefVal As Variant, ByVal tliTypeInfo As TypeInfo) As String
'This helper function adapted from Microsoft documentation
Dim lngTrackVal As Long
Dim mi As MemberInfo
Dim tliTypeKinds As TypeKinds
    
If tliTypeInfo Is Nothing Then
    Select Case VarType(DefVal)
        Case vbString
            If Len(DefVal) Then
                ProduceDefaultValue = """" & DefVal & """"
            End If
        Case vbBoolean 'Always show for Boolean
            ProduceDefaultValue = DefVal
        Case vbDate
            If DefVal Then
                ProduceDefaultValue = "#" & DefVal & "#"
            End If
        Case Else 'Numeric Values
            If DefVal <> 0 Then
                ProduceDefaultValue = DefVal
            End If
    End Select
Else
    'Resolve constants to their enums
    tliTypeKinds = tliTypeInfo.TypeKind
    Do While tliTypeKinds = TKIND_ALIAS
        tliTypeKinds = TKIND_MAX
        On Error Resume Next
        Set tliTypeInfo = tliTypeInfo.ResolvedType
        If Err = 0 Then
            tliTypeKinds = tliTypeInfo.TypeKind
        End If
        On Error GoTo 0
    Loop
    If tliTypeInfo.TypeKind = TKIND_ENUM Then
        lngTrackVal = DefVal
        For Each mi In tliTypeInfo.Members
            If mi.Value = lngTrackVal Then
                ProduceDefaultValue = " = " & mi.Name
                Exit For
            End If
        Next
    End If
End If
End Function
Public Function getFunctionsFromFile(sFileName As String) As String
'*****************************
'Purpose: Get all functions from a COM file
'*****************************
    'On Error Resume Next
    Dim srT As SearchResults
    Dim srM As SearchResults
    Dim mi As MemberInfo, mi2 As MemberInfo
    Dim lSearchData As Long
    Dim bIsConstant As Boolean
    Dim strReturn As String
    Dim p As Long, m As Long, t As Long
    Dim bFirstParameter As Boolean
    Dim bParamArray As Boolean
    Dim tliParameterInfo As ParameterInfo
    Dim bDefault As Boolean
    Dim bOptional As Boolean
    Dim tliTypeInfo As TypeInfo
    Dim tliResolvedTypeInfo As TypeInfo
    Dim tliTypeKinds As TypeKinds
    Dim intVarTypeCur As Integer
    Dim bByVal As Boolean
    Dim strTypeName As String
    Dim ConstVal As Variant
'txtEntityPrototype = PrototypeMember(lstTypeInfos.ItemData(lstTypeInfos.ListIndex), tliInvokeKinds, lstMembers.[_Default])
frmMain.txtFunctions.Text = ""

With tliTypeLibInfo
.ContainingFile = sFileName
   
    
    
    Set srT = .GetTypes(, tliStAll, False)
    For t = 1 To srT.count
        
        lSearchData = srT(t).SearchData
        frmMain.txtFunctions.Text = frmMain.txtFunctions.Text & "'==================== " & srT(t).Name & "====================" & vbCrLf & vbCrLf
        Set srM = tliTypeLibInfo.GetMembers(lSearchData)
        
        
        For m = 1 To srM.count
            
            'Text1.Text = Text1.Text & "guid:" & srM(m).Guid & vbCrLf
            DoEvents
            Set mi = tliTypeLibInfo.GetMemberInfo(lSearchData, srM(m).InvokeKinds, srM(m).MemberId, srM(m).Name)
            frmMain.txtFunctions.Text = frmMain.txtFunctions.Text & getNameFromMemberInfo(mi) & vbCrLf
        Next m
    Next t
End With
MsgBox "all done"

End Function
Public Function ReturnGuiOpcode(ByVal SearchData As Long, _
    ByVal InvokeKinds As InvokeKinds, _
    Optional ByVal MemberName As String) As Integer
'*****************************
'Purpose: To return the opcode of a property used in form decompiling
'*****************************
On Error GoTo exitFunction
    Dim tliTypeInfo As TypeInfo
    Dim num As Integer
    With tliTypeLibInfo
        
        With .GetMemberInfo(SearchData, InvokeKinds, , MemberName)
            'Debug.Print "MemberID: 0x" & Hex(.MemberId - &H10000)
        
            num = (.MemberId - 65536)
        End With
     End With
     If num > 255 Then
        num = -1
     End If
     ReturnGuiOpcode = num
     Exit Function
exitFunction:
    ReturnGuiOpcode = -1
Exit Function
End Function
Public Function ReturnDataType(ByVal SearchData As Long, _
    ByVal InvokeKinds As InvokeKinds, _
    Optional ByVal MemberName As String) As String
'*****************************
'Purpose: To return the data type of a property
'*****************************
    On Error GoTo exitFunction
    Dim tliParameterInfo As ParameterInfo
    Dim bFirstParameter As Boolean
    Dim bIsConstant As Boolean
    Dim bByVal As Boolean
    Dim strReturn As String
    Dim ConstVal As Variant
    Dim strTypeName As String
    Dim intVarTypeCur As Integer
    Dim bDefault As Boolean
    Dim bOptional As Boolean
    Dim bParamArray As Boolean
    Dim tliTypeInfo As TypeInfo
    Dim tliResolvedTypeInfo As TypeInfo
    Dim tliTypeKinds As TypeKinds
  
    With tliTypeLibInfo
        
        'First, determine the type of member we're dealing with
        bIsConstant = GetSearchType(SearchData) And tliStConstants
        With .GetMemberInfo(SearchData, InvokeKinds, , MemberName)

        
            If bIsConstant Then
                ConstVal = .Value
                strReturn = strReturn & " = " & ConstVal
                Select Case VarType(ConstVal)
                    Case vbInteger, vbLong
                        If ConstVal < 0 Or ConstVal > 15 Then
                            strReturn = strReturn & " (&H" & Hex$(ConstVal) & ")"
                        End If
                End Select
            Else
                With .ReturnType
                    intVarTypeCur = .VarType
                    If intVarTypeCur = 0 Or (intVarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
                        On Error Resume Next
                        If Not .TypeInfo Is Nothing Then
                            If Err Then 'Information not available
                                strReturn = strReturn & " As ?"
                            Else
                                If .IsExternalType Then
                                    strReturn = strReturn & .TypeLibInfoExternal.Name & "." & .TypeInfo.Name
                                Else
                                    strReturn = strReturn & .TypeInfo.Name
                                End If
                            End If
                        End If
                        
                        If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                            strReturn = strReturn & "()"
                        End If
                        On Error GoTo 0
                    Else
                        Select Case intVarTypeCur
                            Case VT_VARIANT, VT_VOID, VT_HRESULT
                            Case Else
                                strTypeName = TypeName(.TypedVariant)
                                If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                    strReturn = strReturn & "() As " & Left$(strTypeName, Len(strTypeName) - 2)
                                Else
                                    strReturn = strReturn & strTypeName
                                End If
                        End Select
                    End If
                End With
            End If
            
            ReturnDataType = strReturn & vbCrLf

        End With
    End With
exitFunction:
    
End Function

Public Sub ProcessTypeLibrary()
'*****************************
'Purpose: Procces the Type Libary
'*****************************
    'Clear lists
    frmMain.lstTypeInfos.Clear
    frmMain.lstMembers.Clear
    
    'Display members for type library
    tliTypeLibInfo.GetTypesDirect frmMain.lstTypeInfos.hwnd, , tliStAll
End Sub
Public Function getEventsFromFile(sFileName As String) As String
'*****************************
'Purpose: Get all events from file
'*****************************
    'On Error Resume Next
    Dim srT As SearchResults
    Dim srM As SearchResults
    Dim mi As MemberInfo, mi2 As MemberInfo
    Dim lSearchData As Long
    Dim p As Long, m As Long, t As Long
    Dim LibInfoGuid As TypeLibInfo

    Dim tliTypeInfo As TypeInfo
    With tliTypeLibInfo
    
        .ContainingFile = sFileName

         Set srT = .GetTypes(, tliStEvents, False)
        For t = 1 To srT.count
        
            lSearchData = srT(t).SearchData

              Set tliTypeInfo = tliTypeLibInfo.GetTypeInfo(Replace(Replace(srT(t).Name, "<", ""), ">", ""))
              
           
            frmMain.txtCode.Text = frmMain.txtCode.Text & "'==================== " & srT(t).Name & "====================" & tliTypeLibInfo.GUID & vbCrLf & vbCrLf
     
            Set srM = tliTypeLibInfo.GetMembers(lSearchData)
            
            For m = 1 To srM.count
            
                DoEvents
                
                Set mi = .GetMemberInfo(lSearchData, srM(m).InvokeKinds, srM(m).MemberId, srM(m).Name)
                
                frmMain.txtCode.Text = frmMain.txtCode.Text & getEventInfo(mi, srT(t).Name, True) & vbCrLf
             '   Call getEventInfo(mi, srT(t).Name)
            Next m
        Next t
End With

End Function
Public Function getEventInfo(mi As MemberInfo, ObjectName As String, ShowOpcode As Boolean) As String
'*****************************
'Purpose: Get a specific event information
'*****************************
Dim sOutput As String, strTypeName As String, ConstVal As String
Dim lSearchData As Long
Dim bIsConstant As Boolean, bDefault As Boolean, bFirstParameter As Boolean
Dim bParamArray As Boolean, bOptional As Boolean, bByVal As Boolean
Dim tliParameterInfo As ParameterInfo
Dim tliTypeInfo As TypeInfo, tliResolvedTypeInfo As TypeInfo
Dim tliTypeKinds As TypeKinds
Dim intVarTypeCur As Integer
            With mi
                If ShowOpcode = True Then
                sOutput = sOutput & .VTableOffset
                
                End If
                bIsConstant = GetSearchType(lSearchData) And tliStConstants
                
                sOutput = sOutput & .Name
                With .Parameters
                    If .count Then
                        sOutput = sOutput & " ("
                        bFirstParameter = True
                        bParamArray = .OptionalCount = -1
                        For Each tliParameterInfo In .Me
                            If Not bFirstParameter Then
                                sOutput = sOutput & ", "
                            End If
                            bFirstParameter = False
                            bDefault = tliParameterInfo.Default
                            bOptional = bDefault Or tliParameterInfo.Optional
                            If bOptional Then
                                If bParamArray Then
                                    'This will be the only optional parameter
                                    sOutput = sOutput & "[ParamArray "
                                Else
                                    sOutput = sOutput & "["
                                End If
                            End If
                        
                            With tliParameterInfo.VarTypeInfo
                                Set tliTypeInfo = Nothing
                                Set tliResolvedTypeInfo = Nothing
                                tliTypeKinds = TKIND_MAX
                                intVarTypeCur = .VarType
                                If (intVarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
                                    On Error Resume Next
                                    Set tliTypeInfo = .TypeInfo
                                    If Not tliTypeInfo Is Nothing Then
                                        Set tliResolvedTypeInfo = tliTypeInfo
                                        tliTypeKinds = tliResolvedTypeInfo.TypeKind
                                        Do While tliTypeKinds = TKIND_ALIAS
                                            tliTypeKinds = TKIND_MAX
                                            Set tliResolvedTypeInfo = tliResolvedTypeInfo.ResolvedType
                                            If Err Then
                                                Err.Clear
                                            Else
                                                tliTypeKinds = tliResolvedTypeInfo.TypeKind
                                            End If
                                        Loop
                                    End If
                                
                                    'Determine whether parameters are ByVal or ByRef
                                    Select Case tliTypeKinds
                                        Case TKIND_INTERFACE, TKIND_COCLASS, TKIND_DISPATCH
                                            bByVal = .PointerLevel = 1
                                        Case TKIND_RECORD
                                            'Records not passed ByVal in VB
                                            bByVal = False
                                        Case Else
                                            bByVal = .PointerLevel = 0
                                    End Select
                                
                                    'Indicate ByVal
                                    If bByVal Then
                                        sOutput = sOutput & "ByVal "
                                    End If
                                
                                    'Display the parameter name
                                    sOutput = sOutput & tliParameterInfo.Name
                                
                                    If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                        sOutput = sOutput & "()"
                                    End If
                                    
                                    If tliTypeInfo Is Nothing Then 'Information not available
                                        sOutput = sOutput & " As ?"
                                    Else
                                        If .IsExternalType Then
                                            sOutput = sOutput & " As " & .TypeLibInfoExternal.Name & "." & tliTypeInfo.Name
                                        Else
                                            sOutput = sOutput & " As " & tliTypeInfo.Name
                                        End If
                                    End If
                                
                                    'Reset error handling
                                    On Error GoTo 0
                                Else
                                    If .PointerLevel = 0 Then
                                        sOutput = sOutput & "ByVal "
                                    End If
                                        
                                    sOutput = sOutput & tliParameterInfo.Name
                                    If intVarTypeCur <> vbVariant Then
                                        strTypeName = TypeName(.TypedVariant)
                                        If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                            sOutput = sOutput & "() As " & Left$(strTypeName, Len(strTypeName) - 2)
                                        Else
                                            sOutput = sOutput & " As " & strTypeName
                                        End If
                                    End If
                                End If
                                    
                                If bOptional Then
                                    If bDefault Then
                                        sOutput = sOutput & ProduceDefaultValue(tliParameterInfo.DefaultValue, tliResolvedTypeInfo)
                                        'sOutput = sOutput & " = " & tliParameterInfo.DefaultValue
                                    End If
                                    sOutput = sOutput & "]"
                                End If
                            End With
                        Next
                        sOutput = sOutput & ")"
                    End If
                End With
                'return type
                If bIsConstant Then
                    ConstVal = .Value
                    sOutput = sOutput & " = " & ConstVal
                    Select Case VarType(ConstVal)
                        Case vbInteger, vbLong
                            If ConstVal < 0 Or ConstVal > 15 Then
                                sOutput = sOutput & " (&H" & Hex$(ConstVal) & ")"
                            End If
                    End Select
                Else
                    With .ReturnType
                        intVarTypeCur = .VarType
                        If intVarTypeCur = 0 Or (intVarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
                            On Error Resume Next
                            If Not .TypeInfo Is Nothing Then
                                If Err Then 'Information not available
                                    sOutput = sOutput & " As ?"
                                Else
                                    If .IsExternalType Then
                                        sOutput = sOutput & " As " & .TypeLibInfoExternal.Name & "." & .TypeInfo.Name
                                    Else
                                        sOutput = sOutput & " As " & .TypeInfo.Name
                                    End If
                                End If
                            End If
                            
                            If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                sOutput = sOutput & "()"
                            End If
                            On Error GoTo 0
                        Else
                            Select Case intVarTypeCur
                                Case VT_VARIANT, VT_VOID, VT_HRESULT
                                Case Else
                                    strTypeName = TypeName(.TypedVariant)
                                    If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                        sOutput = sOutput & "() As " & Left$(strTypeName, Len(strTypeName) - 2)
                                    Else
                                        sOutput = sOutput & " As " & strTypeName
                                    End If
                            End Select
                        End If
                    End With
                End If
            End With
        getEventInfo = sOutput
End Function

Public Function ReturnHelpString(ByVal SearchData As Long, ByVal InvokeKinds As InvokeKinds, Optional ByVal MemberName As String) As String
'*****************************
'Purpose: To return the help string used on textbox in form editor to describe function
'*****************************
    With tliTypeLibInfo
        'First, determine the type of member we're dealing with
        With .GetMemberInfo(SearchData, InvokeKinds, , MemberName)
            ReturnHelpString = .HelpString
        End With
    End With

End Function
Public Function getEventComplete(sFileName As String, strGuid As String, EventNum As Integer) As String
'*****************************
'Purpose: To return all the events from a filename by COM
'*****************************
    'On Error Resume Next
    Dim srT As SearchResults
    Dim srM As SearchResults
    Dim mi As MemberInfo, mi2 As MemberInfo
    Dim lSearchData As Long
    Dim p As Long, m As Long, t As Long
    Dim LibInfoGuid As TypeLibInfo

    Dim tliTypeInfo As TypeInfo
    With tliTypeLibInfo
    
        .ContainingFile = sFileName

         Set srT = .GetTypes(, tliStEvents, False)
        For t = 1 To srT.count
        
            lSearchData = srT(t).SearchData

            Set tliTypeInfo = tliTypeLibInfo.GetTypeInfo(Replace(Replace(srT(t).Name, "<", ""), ">", ""))

            'frmMain.txtCode.Text = frmMain.txtCode.Text & "'==================== " & srT(t).Name & "====================" & tliTypeLibInfo.GUID & vbCrLf & vbCrLf
            If tliTypeInfo.GUID = strGuid Then
               MsgBox "GuidFound " & srT(t).Name
               Set srM = tliTypeLibInfo.GetMembers(lSearchData)
            
            For m = 1 To srM.count
            
                DoEvents
                If m = EventNum Then
                Set mi = .GetMemberInfo(lSearchData, srM(m).InvokeKinds, srM(m).MemberId, srM(m).Name)
                
                'frmMain.txtCode.Text = frmMain.txtCode.Text & getEventInfo(mi, srT(t).Name, False) & vbCrLf
                getEventComplete = getEventInfo(mi, srT(t).Name, False)
                
                Exit Function
                End If
                Next m
           End If
            '
        Next t
End With

End Function

