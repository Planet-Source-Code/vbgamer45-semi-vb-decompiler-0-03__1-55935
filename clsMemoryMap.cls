VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMemoryMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'Purpose outputs an html view of the exe with important strutures highlighted
Option Explicit
Private Type tSector
    SectorID As Integer
    SectorName As String
    SectorBegin As Long
    SectorEnd As Long
    SectorLength As Long
    CollisionID As Long
    Collision As Boolean
End Type

Private mSectors() As tSector
Private mNumSectors As Long
Private mForm As frmMain

Private Sub SwitchSectors(ByRef sec1 As tSector, ByRef sec2 As tSector)
    Dim tSectorID As Integer, tSectorName As String, tSectorBegin As Long
    Dim tSectorEnd As Long, tCollisionID As Long, tCollision As Boolean, tSecLen As Long
    With sec1
        'save sector 1
        tCollision = .Collision
        tCollisionID = .CollisionID
        tSectorBegin = .SectorBegin
        tSecLen = .SectorLength
        tSectorEnd = .SectorEnd
        tSectorID = .SectorID
        tSectorName = .SectorName
        'set sector 1 to sector 2 values
        .Collision = sec2.Collision
        .CollisionID = sec2.CollisionID
        .SectorBegin = sec2.SectorBegin
        .SectorLength = sec2.SectorLength
        .SectorEnd = sec2.SectorEnd
        .SectorID = sec2.SectorID
        .SectorName = sec2.SectorName
    End With
    With sec2
        'set sector 2 to saved sector 1 values
        .Collision = tCollision
        .CollisionID = tCollisionID
        .SectorBegin = tSectorBegin
        .SectorLength = tSecLen
        .SectorEnd = tSectorEnd
        .SectorID = tSectorID
        .SectorName = tSectorName
    End With
End Sub

Public Sub RenameSector(OldSector As String, NewSector As String)


    Dim i As Integer
    For i = 0 To mNumSectors - 1: DoEvents
        If mSectors(i).SectorName = OldSector Then
            mSectors(i).SectorName = NewSector
        End If
    Next i
End Sub

Private Function SectorByID(id As Integer) As tSector
    Dim i As Integer
    SectorByID.SectorID = -1
    For i = 0 To mNumSectors - 1: DoEvents
        If mSectors(i).SectorID = id Then
            SectorByID = mSectors(i)
            Exit Function
        End If
    Next i
End Function
Private Function SortSectors() As Integer
Dim i As Integer, j As Integer, numMoves As Integer
For i = 0 To mNumSectors - 1: DoEvents
    For j = i To mNumSectors - 1: DoEvents
        If mSectors(j).SectorBegin < mSectors(i).SectorBegin Then
            Call SwitchSectors(mSectors(j), mSectors(i))
            numMoves = numMoves + 1
        End If
    Next j
Next i
SortSectors = numMoves
End Function
Public Sub ShowForm()
  Set mForm = New frmMain
    Dim iSector As Integer
    Dim lAddress As Long
    'Sectors are not sorted
    'this function sorts and then displays
  '  gVBFile.Display.Verbose "Sorting Memory Map Sectors: " & SortSectors & " moves"
    For iSector = 0 To mNumSectors - 1: DoEvents
        With mSectors(iSector)
            'add sector to form
            If .SectorBegin > lAddress Then
            '    Call mForm.AddItem(-1, "", .SectorBegin - lAddress, lAddress, .SectorBegin - 1, True, 0)
            End If
            'Call mForm.AddItem(.SectorID, .SectorName, .SectorLength, .SectorBegin, .SectorEnd, .Collision, .CollisionID)
            lAddress = .SectorEnd + 1
        End With
    Next iSector
End Sub


'use this to map out memory areas

Public Function AddSector(ByVal Begin As Long, ByVal Size As Long, Optional Name As String = "Undefined") As Boolean
'function will return true if there is a collision, else false
ReDim Preserve mSectors(mNumSectors) 'add new sector
Dim iSector As Integer

With mSectors(mNumSectors)
    .SectorID = mNumSectors  'AutoInt
    .SectorName = Name
    .SectorBegin = Begin
    .SectorLength = Size
    .SectorEnd = Begin + Size
    
    For iSector = 0 To mNumSectors - 1: DoEvents
        'search for a collision
        
        If (.SectorBegin <= mSectors(iSector).SectorEnd - 1 And .SectorBegin > mSectors(iSector).SectorBegin) Or _
           (.SectorEnd <= mSectors(iSector).SectorEnd And .SectorEnd > mSectors(iSector).SectorBegin) Then
           'collision detected mark it on both sectors
           .Collision = True
           .CollisionID = iSector
           mSectors(iSector).Collision = True
           mSectors(iSector).CollisionID = .SectorID
          ' gVBFile.Display.Warning "Collision Detected: " & mSectors(iSector).SectorName & "(" & mSectors(iSector).SectorBegin & "-" & mSectors(iSector).SectorEnd & ") & " & .SectorName & "(" & .SectorBegin & "-" & .SectorEnd & ")", "clsMemoryMap.AddSector"
        End If
        
    Next iSector
    AddSector = .Collision
End With
'prepare for next time
mNumSectors = mNumSectors + 1
End Function
Property Get NumSectors() As String
NumSectors = mNumSectors
End Property


Private Sub writeit(FileNum As Integer, Stuff As String)
Put #FileNum, , Stuff
DoEvents
End Sub

Public Sub ExportToHTML()
Dim sOutString As String, sLine As String, sColumns As String, sLineBin As String, sStatus As String, sNewStatus As String
Dim lOffset As Long
Dim bAdd As Byte
Dim iColumn As Integer, iSector As Integer, iColumnCount As Integer

'Headers
sOutString = _
    "<html><head><style></style>" & vbCrLf & _
    "<style type=""text/css"">" & vbCrLf & _
    "<!--" & vbCrLf & _
    "TD,BODY,TR,HTML{font-face:verdana;font-size:12 px}" & vbCrLf & _
    ".tooltiptitle{COLOR: #FFFFFF; TEXT-DECORATION: none; CURSOR: Default; font-family: arial; font-weight: bold; font-size: 8pt}" & vbCrLf & _
    ".tooltipcontent{COLOR: #000000; TEXT-DECORATION: none; CURSOR: Default; font-family: arial; font-size: 8pt}" & vbCrLf & _
    "#p{position:absolute; width: 100px; top: 0px; left: 0px; z-index:4; visibility:hidden;}" & vbCrLf & _
    "-->" & vbCrLf & "</style>" & vbCrLf & _
    "<script language = ""javascript"">" & vbCrLf & _
    "<!--" & vbCrLf & _
    "var ie = document.all ? 1 : 0" & vbCrLf & _
    "var ns = document.layers ? 1 : 0" & vbCrLf & _
    "if(ns){doc = ""document.""; sty = """"}" & vbCrLf & _
    "if(ie){doc = ""document.all.""; sty = "".style""}" & vbCrLf & _
    "var initialize = 0" & vbCrLf & _
    "var ex, ey, topColor, subColor, ContentInfo" & vbCrLf & _
    "if(ie){" & vbCrLf & _
    "ex = ""event.x""" & vbCrLf & _
    "ey = ""event.y""" & vbCrLf & _
    "topColor = ""#808080""" & vbCrLf & _
    "subColor = ""#C0C0C0""" & vbCrLf & _
    "}" & vbCrLf & _
    "if(ns){" & vbCrLf & _
    "ex = ""e.pageX""" & vbCrLf
'more headers
sOutString = sOutString & _
    "ey = ""e.pageY""" & vbCrLf & _
    "window.captureEvents(Event.MOUSEMOVE)" & vbCrLf & _
    "window.onmousemove = overhere" & vbCrLf & _
    "topColor = ""#808080""" & vbCrLf & _
    "subColor = ""#C0C0C0""" & vbCrLf & _
    "}" & vbCrLf & _
    "function MoveToolTip(layerName, FromTop, FromLeft, e){" & vbCrLf & _
    "if(ie){eval(doc + layerName + sty + "".top = ""  + (eval(FromTop) + document.body.scrollTop))}" & vbCrLf & _
    "if(ns){eval(doc + layerName + sty + "".top = ""  +  eval(FromTop))}" & vbCrLf & _
    "eval (doc + layerName + sty + "".left = "" + (eval(FromLeft) + 15))" & vbCrLf & _
    "}" & vbCrLf & _
    "function ReplaceContent(layerName){if(ie){eval('document.all.'+layerName+'.innerHTML = ContentInfo')}if(ns){with(document.layers[layerName].document){open();write(ContentInfo);close();}}}" & vbCrLf & _
    "function Activate(){initialize=1}" & vbCrLf & _
    "function deActivate(){initialize=0}" & vbCrLf & _
    "function overhere(e){if(initialize){MoveToolTip(""p"", ey, ex, e);eval (doc + ""p"" + sty + "".visibility = 'visible'"");}else{MoveToolTip(""p"", 0, 0);eval (doc + ""p"" + sty + "".visibility = 'hidden'"");}}" & vbCrLf & _
    "function tt(TContent){ContentInfo = '<table border=""0"" width=""150"" cellspacing=""0"" cellpadding=""0"">'+'<tr><td width=""100%"" bgcolor=""#000000"">'+'<table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""0"">'+'<tr><td width=""100%"" bgcolor='+topColor+'>'+'<table border=""0"" width=""90%"" cellspacing=""0"" cellpadding=""0"" align=""center"">'+'<tr><td width=""100%"">'+'<font class=""tooltiptitle"">&nbsp;Info:</font>'+'</td></tr>'+'</table>'+'</td></tr>'+'<tr><td width=""100%"" bgcolor='+subColor+'>'+'<table border=""0"" width=""90%"" cellpadding=""0"" cellspacing=""1"" align=""center"">'+'<tr><td width=""100%"">'+'<font class=""tooltipcontent"">'+TContent+'</font>'+'</td></tr>'+'</table>'+'</td></tr>'+'</table>'+'</td></tr>'+'</table>';ReplaceContent ('p'); Activate();}" & vbCrLf & _
    "//-->" & vbCrLf & _
    "</script>" & vbCrLf & _
    "<Title>" & gVBFile.ShortFileName & " Memory Map</Title>" & vbCrLf & _
    "</head><body onmousemove=""overhere()""><div id=""p""></div>" & vbCrLf & _
    "<TABLE cellSpacing=""1"" cellPadding=""1"" border=""0"">"
'Add rows
'    <td align=right><b>00000000</b></td>
'    <td colspan=2> </td>
'    <td onMouseover="tt('MZ Header');" onMouseout="deActivate()" bgcolor="#000000">4D5A</td>
'    <td onMouseover="tt('MZ Header');" onMouseout="deActivate()" bgcolor="#000000">9000</td>

'    <td onMouseover="tt('MZ Header');" onMouseout="deActivate()" bgcolor="#000000">0300</td>
'    <td onMouseover="tt('MZ Header');" onMouseout="deActivate()" bgcolor="#000000">0000</td>

'    <td onMouseover="tt('MZ Header');" onMouseout="deActivate()" bgcolor="#000000">0400</td>
'    <td onMouseover="tt('MZ Header');" onMouseout="deActivate()" bgcolor="#000000">0000</td>

'    <td onMouseover="tt('MZ Header');" onMouseout="deActivate()" bgcolor="#000000">FFFF</td>
'    <td onMouseover="tt('MZ Header');" onMouseout="deActivate()" bgcolor="#000000">0000</td>
'    <td> </td>
'    <td>MZê.......ˇˇ..</td>
  '  gVBFile.Display.Verbose "Sorting Memory Map Sectors: " & SortSectors & " moves"
    'iColumn = 16 'init so that it will run line Offset
     
    sStatus = SectorStatus(lOffset)
    sLine = "<td align=right><b>00000000 &nbsp; &nbsp; </b></td>"

    For lOffset = 0 To gVBFile.length ': DoEvents
      '  gVBFile.Display.Verbose "Exporting Memory Map Offset: 0x" & PadHex(VBA.Hex(lOffset), 5)
        'Debug.Assert lOffset < &H1158
       ' DoEvents
       ' MsgBox "AGA"
        sNewStatus = SectorStatus(lOffset)
        If sNewStatus <> sStatus Or iColumn = 16 Then 'breakpoints
            'break points
            '  16th column
            '  change of status
            
            sLine = sLine & "<td colspan=" & iColumnCount
            If sStatus <> "" Then
                sLine = sLine & " onMouseover=""tt('" & sStatus & "');"" onMouseout=""deActivate()"" bgcolor=""#"
                If sStatus <> sNewStatus Then
                    sLine = sLine & SectorColorByOffset(lOffset - 1) 'get previous sector offset
                Else
                    sLine = sLine & SectorColorByOffset(lOffset) 'get current sector offset
                End If
                sLine = sLine & """"
            End If
            If Right(sColumns, 8) = " &nbsp; " Then 'trim trialing &nbsp;
                sColumns = Left(sColumns, Len(sColumns) - 8)
                If Right(sColumns, 8) = " &nbsp; " Then 'trim trialing &nbsp;
                    sColumns = Left(sColumns, Len(sColumns) - 8)
                End If
            End If
            sLine = sLine & ">" & sColumns & "</td>"
            sColumns = ""
            sStatus = sNewStatus
            iColumnCount = 0
        End If
        If iColumn = 16 Then 'flush last line
            sLine = sLine & "<td> </td><td>" & sLineBin & "<td>"
            sOutString = sOutString & "<tr>" & sLine & "</tr>" & vbCrLf
            sLine = "<td align=right><b>" & PadHex(VBA.Hex(lOffset)) & " &nbsp; &nbsp; </b></td>"
            sLineBin = ""
            iColumn = 0
        End If
        iColumn = iColumn + 1
        iColumnCount = iColumnCount + 1
        bAdd = gVBFile.GetByte(lOffset)
       ' MsgBox bAdd
        sColumns = sColumns & PadHex(VBA.Hex(bAdd), 2)
        If bAdd = 0 Then bAdd = 46
        sLineBin = sLineBin & Chr$(bAdd)
        If iColumn Mod 4 = 0 Then 'add spacing
            If iColumnCount <> 0 Then sColumns = sColumns & " &nbsp; "
        ElseIf iColumn Mod 2 = 0 Then
            If iColumnCount <> 0 Then sColumns = sColumns & " "
        End If
    Next lOffset

    If iColumn <> 16 Then
        iColumnCount = iColumnCount + (16 - iColumn)
        sLine = sLine & "<td colspan=" & iColumnCount & ">" & sColumns & "</td>" & _
                        "<td> </td>" & _
                        "<td>" & sLineBin & "<td>"
        sOutString = sOutString & "<tr>" & sLine & "</tr>" & vbCrLf
    End If
            
            
'tail
sOutString = sOutString & "</table></body></html>"
If Len(Dir(gVBFile.ShortFileName & ".html")) Then Call Kill(gVBFile.ShortFileName & ".html")
Open (App.path & "\" & gVBFile.ShortFileName & ".html") For Binary As #10
    Put #10, , sOutString
Close #10
'gVBFile.Display.Verbose ""
End Sub
Private Function SectorColorByOffset(lOffset As Long) As String
    SectorColorByOffset = SectorColorByName(SectorByOffset(lOffset).SectorName)
End Function
Private Function SectorColorByName(ByVal SectorName As String) As String
    SectorName = LCase(SectorName)
    If InStr(1, SectorName, "vb header") = 1 Or InStr(1, SectorName, "vbheader") = 1 Then
        SectorColorByName = "ff0000" 'red
    ElseIf InStr(1, SectorName, "optional") = 1 Or InStr(1, SectorName, "mz") = 1 Or InStr(1, SectorName, "pe") = 1 Or InStr(1, SectorName, "section") = 1 Then
        SectorColorByName = "002222"
    ElseIf InStr(1, SectorName, "moduleinfo") = 1 Then
        SectorColorByName = "ffff00"
    ElseIf InStr(1, SectorName, "forminfo") = 1 Then
        SectorColorByName = "ffff00"
    ElseIf InStr(1, SectorName, "control") = 1 Then
        SectorColorByName = "00cc00"
    ElseIf InStr(1, SectorName, ".eventtable.ptr") Then
        SectorColorByName = "C0C0C0"
    ElseIf InStr(1, SectorName, ".eventtable") Then
        SectorColorByName = "808080"
    ElseIf InStr(1, SectorName, "project info") = 1 Then
        SectorColorByName = "ff3300"
    ElseIf InStr(1, SectorName, "import table") = 1 Then
        SectorColorByName = "ff6600"
    ElseIf InStr(1, SectorName, "objecttable") = 1 Then
        SectorColorByName = "ffcc00"
    ElseIf InStr(1, SectorName, "object") = 1 Then
        SectorColorByName = "ff9900"
    ElseIf InStr(1, SectorName, "form source") = 1 Then
        SectorColorByName = "0099ff"
    ElseIf InStr(1, SectorName, "guiforminfo") = 1 Then
        SectorColorByName = "0066ff"
    ElseIf InStr(1, SectorName, "guitable") = 1 Then
        SectorColorByName = "0033ff"
    
    ElseIf InStr(1, SectorName, "code") = 1 Then
        SectorColorByName = "00ccff"
    Else
      '  gVBFile.Display.Warning "Color for SectorName not found:" & SectorName, "clsMemoryMap.SectorColorByName"
        SectorColorByName = "000000"
    End If
End Function
Private Function SectorByOffset(lOffset As Long) As tSector
    Dim iSector As Integer
    For iSector = 0 To mNumSectors - 1: DoEvents
        With mSectors(iSector)
            If .SectorBegin <= lOffset And .SectorEnd > lOffset Then 'offset is in sector
                SectorByOffset = mSectors(iSector)
                Exit Function
            End If
        End With
    Next iSector
End Function
Private Function SectorIndexByID(id As Long) As Long  'returns array index from ID
    Dim iSector As Integer
    For iSector = 0 To mNumSectors - 1: DoEvents
        With mSectors(iSector)
            If .SectorID = id Then 'offset is in sector
                SectorIndexByID = iSector
                Exit Function
            End If
        End With
    Next iSector
    SectorIndexByID = -1
End Function
Private Function SectorStatus(lOffset As Long) As String
    Dim iSector As Integer
    Dim sCollision As tSector
    Dim lRangeHi As Long
    Dim lRangeLo As Long
    'iSector = SectorByOffset(lOffset)
    'If iSector = -1 Then Exit Function
    With SectorByOffset(lOffset)
        If .CollisionID = 0 And .SectorID = 0 And .SectorLength = 0 Then Exit Function
        If .SectorBegin <= lOffset And .SectorEnd > lOffset Then 'offset is in sector
            SectorStatus = .SectorName & "[0x" & VBA.Hex(.SectorLength) & "]<br>Start: 0x" & VBA.Hex(.SectorBegin) & "<br>End: 0x" & VBA.Hex(.SectorEnd)
            If .Collision Then
                sCollision = SectorByID(CInt(.CollisionID))
                If sCollision.SectorBegin < .SectorBegin And sCollision.SectorEnd >= .SectorBegin Then '. is start col is finish
                    lRangeLo = .SectorBegin
                    lRangeHi = sCollision.SectorEnd
                ElseIf .SectorBegin < sCollision.SectorBegin And .SectorEnd >= sCollision.SectorBegin Then
                    lRangeLo = sCollision.SectorBegin
                    lRangeHi = .SectorEnd
                End If
                If lOffset > lRangeLo And lOffset <= lRangeHi Then
                    SectorStatus = SectorStatus & "<br><font color=red>Collision with " & sCollision.SectorName & "[0x" & VBA.Hex(sCollision.SectorLength) & "]" & _
                             "<br>Start: 0x" & VBA.Hex(sCollision.SectorBegin) & _
                             "<br>End: 0x" & VBA.Hex(sCollision.SectorEnd) & "<br>" & _
                             "<br>Col Start: 0x" & VBA.Hex(lRangeLo) & _
                             "<br>Col End:   0x" & VBA.Hex(lRangeHi) & _
                             "<br>OverLap: 0x" & VBA.Hex(lRangeHi - lRangeLo) & "</font>"
                End If
            End If
            Exit Function
        End If
    End With
    
    'SectorStatus = ""
End Function

