VERSION 5.00
Begin VB.Form frmPcode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P-Code Procedure Decompile View"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox txtView 
      Height          =   3570
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   360
      Width           =   4935
   End
   Begin VB.ListBox lstProcedures 
      Height          =   3570
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Just click on a procedure in the list to decompile it."
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4080
      Width           =   4575
   End
   Begin VB.Label lblTitle 
      Caption         =   "Procedure List:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmPcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#############################################
'frmPcode vbgamer45 2004
'##############################################
Option Explicit
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'*****************************
'Purpose: To load all events into the listbox
'*****************************
    Dim ProcAddr() As Long
    Dim g As Integer, i As Integer

    Call modPCode4.LoadPE(SFilePath)
    Call modPCode4.LoadPcode

    'Get all procedures
    lstProcedures.Clear
    Open SFilePath For Binary Access Read As #24
    For i = 0 To UBound(gObjectInfoHolder)
        If gObjectInfoHolder(i).NumberOfProcs > 0 Then
        ReDim ProcAddr(gObjectInfoHolder(i).NumberOfProcs - 1)
        Seek #24, gObjectInfoHolder(i).aProcTable + 1 - OptHeader.ImageBase
        Get #24, , ProcAddr
        For g = 0 To UBound(ProcAddr)
            If ProcAddr(g) <> 0 And ProcAddr(g) <> -1 Then
                If ProcAddr(g) < UBound(SubName) And ProcAddr(g) > LBound(SubName) Then
                    SubName(ProcAddr(g)) = gObjectNameArray(i) & ".Proc" & ProcAddr(g)
                    lstProcedures.AddItem ProcAddr(g)
                End If
            End If
        Next
        End If
    Next
        Dim addrSubMain As Long
        If gVBHeader.aSubMain <> 0 Then
            Seek #24, gVBHeader.aSubMain + 2 - OptHeader.ImageBase
            Get #24, , addrSubMain
          Dim sTemp
            sTemp = Split(SubName(addrSubMain), ".")
            SubName(addrSubMain) = sTemp(0) & ".Sub Main"
        End If
    Close #24
    'Add Event ProcLists
    For i = 0 To UBound(EventProcList) - 1
        If EventProcList(i) <> 0 Then
            lstProcedures.AddItem EventProcList(i)
        End If
    Next
    For i = 0 To UBound(SubNamelist) - 1
        If SubNamelist(i).offset < UBound(SubName) Then
            SubName(SubNamelist(i).offset) = SubNamelist(i).strName
        End If
    Next
    
End Sub

Private Sub lstProcedures_Click()
    If lstProcedures.Text <> "" Then
        txtView.Text = modPCode4.DecompileProc(lstProcedures.Text)
    End If
End Sub
