VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Tag             =   "                                  vbgamer45"
   Begin VB.CheckBox chkPCODE 
      Caption         =   "Disable P-Code Decompile"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   2895
   End
   Begin VB.CheckBox chkShowColors 
      Caption         =   "Show Colors"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CheckBox chkShowOffsets 
      Caption         =   "Show Offests and Gui Opcodes"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CheckBox chkDumpControls 
      Caption         =   "Dump Control/Form raw binary data"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   3015
   End
   Begin VB.CheckBox chkSkipCOM 
      Caption         =   "Skip COM and Control/Form Property Processing"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   2760
      Width           =   1815
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#############################################
'frmOptions vbgamer45 2004
'##############################################
Private Sub chkDumpControls_Click()
    If chkDumpControls.Value = vbChecked Then
        gDumpData = True
    Else
        gDumpData = False
    End If
End Sub

Private Sub chkPCODE_Click()
    If chkPCODE.Value = vbChecked Then
        gPcodeDecompile = False
    Else
        gPcodeDecompile = True
    End If
End Sub

Private Sub chkShowColors_Click()
    If chkShowColors.Value = vbChecked Then
        gShowColors = True
    Else
        gShowColors = False
    End If
End Sub

Private Sub chkShowOffsets_Click()
    If chkShowOffsets.Value = vbChecked Then
        gShowOffsets = True
    Else
        gShowOffsets = False
    End If
End Sub

Private Sub chkSkipCOM_Click()
    If chkSkipCOM.Value = vbChecked Then
        gSkipCom = True
    Else
        gSkipCom = False
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If gSkipCom = True Then chkSkipCOM.Value = vbChecked
    If gDumpData = True Then Me.chkDumpControls.Value = vbChecked
    If gShowOffsets = True Then Me.chkShowOffsets.Value = vbChecked
    If gShowColors = True Then Me.chkShowColors.Value = vbChecked
    If gPcodeDecompile = False Then Me.chkShowColors.Value = vbChecked
End Sub
