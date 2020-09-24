VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Semi VB Decompiler by vbgamer45"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7950
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1                   vbgamer45"
   ScaleHeight     =   6495
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Tag             =   "                                   v b g a m e r 4 5"
   Begin VB.Frame FrameStatus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Form Generating Status"
      Height          =   3135
      Left            =   1680
      TabIndex        =   19
      Top             =   1560
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   255
         Left            =   3600
         TabIndex        =   21
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txtStatus 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   240
         Width           =   4335
      End
   End
   Begin RichTextLib.RichTextBox txtBuffer 
      Height          =   855
      Left            =   8640
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":27A2
   End
   Begin VB.ListBox lstMembers 
      Height          =   2400
      Left            =   7920
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.ListBox lstTypeInfos 
      Height          =   2400
      Left            =   7920
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   3975
   End
   Begin RichTextLib.RichTextBox buffCodeAp 
      Height          =   1935
      Left            =   8760
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   3413
      _Version        =   393217
      TextRTF         =   $"frmMain.frx":2824
   End
   Begin RichTextLib.RichTextBox buffCodeAv 
      Height          =   1575
      Left            =   8040
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2778
      _Version        =   393217
      TextRTF         =   $"frmMain.frx":28AF
   End
   Begin VB.TextBox txtFinal 
      Height          =   1695
      Index           =   0
      Left            =   10080
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
   End
   Begin RichTextLib.RichTextBox txtFunctions 
      Height          =   615
      Left            =   8400
      TabIndex        =   5
      Top             =   5640
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":293A
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   4
      Top             =   6225
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   476
      Style           =   1
      SimpleText      =   "Credits: VB Decompiling community... Sarge, Mr. Unleaded, Moogman, _aLfa_, Alex Ionescu, Warning, and others..."
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   5400
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imglistControl 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   45
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F56
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4542
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":48DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":592E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C80
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5FD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6324
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6676
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":69C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7070
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":73C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7714
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A66
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7DB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":810A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":845C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":87AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8B00
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8E54
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":91A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":94F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":984A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9B9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9EEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A240
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AAB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B324
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BB96
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C408
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CC7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D4EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D83E
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E0B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E922
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F194
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F4E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F83A
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FDD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1036E
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10908
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10EA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10FFC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvProject 
      Height          =   6075
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   10716
      _Version        =   393217
      Indentation     =   617
      LabelEdit       =   1
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "imglistControl"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin TabDlg.SSTab sstViewFile 
      Height          =   6075
      Left            =   3480
      TabIndex        =   1
      Tag             =   "T{20/21/}"
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   10716
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Code"
      TabPicture(0)   =   "frmMain.frx":1134E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtCode"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "imlIcons"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Properties"
      TabPicture(1)   =   "frmMain.frx":1136A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fxgEXEInfo"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Preview"
      TabPicture(2)   =   "frmMain.frx":11386
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picPreview"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Edit Object"
      TabPicture(3)   =   "frmMain.frx":113A2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtEditArray(0)"
      Tab(3).Control(1)=   "lblObjectName"
      Tab(3).Control(2)=   "lblArrayEdit(0)"
      Tab(3).ControlCount=   3
      Begin VB.TextBox txtEditArray 
         Height          =   285
         Index           =   0
         Left            =   -73200
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.PictureBox picPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5415
         Left            =   -74880
         ScaleHeight     =   5385
         ScaleWidth      =   3945
         TabIndex        =   14
         Top             =   480
         Width           =   3975
      End
      Begin MSComctlLib.ImageList imlIcons 
         Left            =   240
         Top             =   1020
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":113BE
               Key             =   "COCLASS"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":11810
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":11C62
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":120B4
               Key             =   "i4"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":12506
               Key             =   "i2"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":12660
               Key             =   "i0"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":127BA
               Key             =   "i1"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":12914
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":12A6E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtCode 
         Height          =   5535
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   9763
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmMain.frx":13008
      End
      Begin MSFlexGridLib.MSFlexGrid fxgEXEInfo 
         Height          =   5535
         Left            =   -74940
         TabIndex        =   2
         Top             =   480
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   9763
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   -2147483627
         ForeColorFixed  =   12829635
         GridColorFixed  =   8421504
         ScrollTrack     =   -1  'True
         GridLinesFixed  =   1
         AllowUserResizing=   3
      End
      Begin MSComDlg.CommonDialog cdlShow 
         Left            =   -74940
         Top             =   7590
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin RichTextLib.RichTextBox txtResult 
         Height          =   675
         Left            =   -74940
         TabIndex        =   3
         Top             =   8190
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   1191
         _Version        =   393217
         BackColor       =   12632256
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":1308A
      End
      Begin VB.Label lblObjectName 
         Caption         =   "ObjectName:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   17
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label lblArrayEdit 
         Caption         =   "Property Name"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   15
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Members:"
      Height          =   195
      Left            =   7920
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TypeInfos:"
      Height          =   195
      Left            =   7920
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileDebugProcess 
         Caption         =   "&Debug VB Process"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileGenerate 
         Caption         =   "&Generate vbp"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuFileSaveExe 
         Caption         =   "&Save Exe Changes"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileExportMemoryMap 
         Caption         =   "&Export Memory Map"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFileAntiDecompiler 
         Caption         =   "&Anti VB Decompiler Protect"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent1 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent2 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent3 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent4 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsPCodeProcedure 
         Caption         =   "&P-Code Procedure Decompile"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###########################################
'#VB Semi Decompiler
'#By vbgamer45
'#Credits:
'#Some code from decompiler.theautomaters.com  The VB Decompiling Community
'#Sarge for the PE Skeleton
'#Mr. Unleaded for MemoryMap
'#Moogman for TypeViewer
'#Brad Martinez for parts of modFrx
'#modAsm from vbAnaylzer
'#Alex Ionescu for his help for COM and strutures
'#And from Warning for treeview
'#Send back your information that you have added.
'#For Contact or request docs email gmdecompiler@yahoo.com
'###########################################

'The following is used for the browse for folder dialog
Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
    End Type
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

'For Syntax Coloring
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const EM_GETLINE = &HC4
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_LINESCROLL = &HB6


'Used for syntax highlighting
Dim prevCountLine As Long
Dim LinesCheck() As String

Private Sub cmdCancel_Click()
    CancelDecompile = True
End Sub

Private Sub Form_Load()
'*****************************
'Purpose: To set all our decompiler and load any functions that need to be loaded.
'*****************************
    Me.Caption = "Semi VB Decompiler by vbgamer45 Version: " & Version
    Call PrintReadMe
    'Setup Variables
    gSkipCom = False
    gDumpData = False
    gShowOffsets = True
    gShowColors = True
    gPcodeDecompile = True
    CancelDecompile = False
    'Get the recent file list
    Dim Recent1Title As String
    Dim Recent2Title As String
    Dim Recent3Title As String
    Dim Recent4Title As String
    
    Recent1Title = GetSetting("VB Decompiler", "Options", "Recent1FileTitle", "")
    Recent2Title = GetSetting("VB Decompiler", "Options", "Recent2FileTitle", "")
    Recent3Title = GetSetting("VB Decompiler", "Options", "Recent3FileTitle", "")
    Recent4Title = GetSetting("VB Decompiler", "Options", "Recent4FileTitle", "")
    
    If Recent1Title <> "" Then
        mnuFileRecent1.Visible = True
        mnuFileSep1.Visible = True
        mnuFileRecent1.Caption = Recent1Title
    End If
    If Recent2Title <> "" Then
        mnuFileRecent2.Visible = True
        mnuFileRecent2.Caption = Recent2Title
    End If
    If Recent3Title <> "" Then
        mnuFileRecent3.Visible = True
        mnuFileRecent3.Caption = Recent3Title
    End If
    If Recent4Title <> "" Then
        mnuFileRecent4.Visible = True
        mnuFileRecent4.Caption = Recent4Title
    End If
    
    'Setup the COM Functions
    Set tliTypeLibInfo = New TypeLibInfo
    'GUID for vb6.olb used to find the gui opcodes of the standard controls
    tliTypeLibInfo.LoadRegTypeLib "{FCFB3D2E-A0FA-1068-A738-08002B3371B5}", 6, 0, 9
    Call ProcessTypeLibrary
    tliTypeLibInfo.AppObjString = "<Global>"
    'Load the functions
  '  Call getFunctionsFromFile("C:\Program Files\Microsoft Visual Studio\VB98\VB6.OLB")
    'Load Com Hacks
    Call modGlobals.LoadCOMFIX
    'Load Events Opcodes for standard controls
    'Call getEventsFromFile(App.Path & "\VB6.OLB")
    
    'Load the vb Function list
    Call modNative.VBFunction_Description_Init(App.path & "\VB60_APIDEF.txt")
    'Init the Asm Engine
    Call modAsm.Init_unASM

    
    ReDim LinesCheck(0)
    LinesCheck(0) = txtCode
    gUpdateText = False
End Sub

Private Sub Form_Resize()
'*****************************
'Purpose: When the form is resized adjust all our controls.
'*****************************
    On Error Resume Next
    tvProject.Height = Me.Height - StatusBar1.Height - 700
    sstViewFile.Height = Me.Height - StatusBar1.Height - 700
    txtCode.Height = sstViewFile.Height - 420
    Me.fxgEXEInfo.Height = sstViewFile.Height - 600
    sstViewFile.Width = Me.Width - tvProject.Width - 200 ' - sstViewFile.Width
    txtCode.Width = sstViewFile.Width - 200
    fxgEXEInfo.Width = sstViewFile.Width - 200
    picPreview.Width = sstViewFile.Width - 200
    picPreview.Height = sstViewFile.Height - 600
End Sub

Private Sub lstMembers_Click()
'Not used(Debug Only)
 Dim tliInvokeKinds As InvokeKinds
    tliInvokeKinds = lstMembers.ItemData(lstMembers.ListIndex)
 
    If lstTypeInfos.ListIndex <> -1 Then
    MsgBox ReturnDataType(lstTypeInfos.ItemData(lstTypeInfos.ListIndex), tliInvokeKinds, lstMembers.[_Default])
    End If
End Sub

Private Sub lstTypeInfos_Click()
'Not Used(Debug Only)
    Dim tliTypeInfo As TypeInfo
    Set tliTypeInfo = tliTypeLibInfo.GetTypeInfo(Replace(Replace(lstTypeInfos.List(lstTypeInfos.ListIndex), "<", ""), ">", ""))
    'Use the ItemData in lstTypeInfos to set the SearchData for lstMembers
    tliTypeLibInfo.GetMembersDirect lstTypeInfos.ItemData(lstTypeInfos.ListIndex), lstMembers.hwnd, , , True
    
End Sub

Private Sub mnuFileAntiDecompiler_Click()
'*****************************
'Purpose: Show save dialog and encypt the current exe
'*****************************
    Cd1.Filename = ""
    Cd1.DialogTitle = "Save File As"
    Cd1.Filter = "Exe Files(*.exe)|*.exe"
    
    Cd1.ShowSave
    
    If Cd1.Filename = "" Then Exit Sub
    
    Call modAntiDecompiler.LoadCrypter
    Call modAntiDecompiler.EncryptExe(SFilePath, Cd1.Filename)
    
End Sub

Private Sub mnuFileExit_Click()
'*****************************
'Purpose: To exit the decompiler and  clear any used memory
'*****************************
    End
End Sub

Private Sub mnuFileExportMemoryMap_Click()
'*****************************
'Purpose: To generate a Memory Map of the current exe file.
'*****************************
    Set gVBFile = Nothing
    Set gVBFile = New clsFile
    Call gVBFile.Setup(SFilePath)
    Dim strTitle As String
    strTitle = Me.Caption
    Me.Caption = "Generating Memory Map...Please Wait..."
    
    Set gMemoryMap = New clsMemoryMap
 
    'hascollision = gMemoryMap.AddSector(0, Len(DosHeader), "mz")
    hascollision = gMemoryMap.AddSector(AppData.PeHeaderOffset, Len(PEHeader), "pe")
    hascollision = gMemoryMap.AddSector(VBStartHeader.PushStartAddress - OptHeader.ImageBase, 102, "vb header")
    hascollision = gMemoryMap.AddSector(gVBHeader.aProjectInfo - OptHeader.ImageBase, 572, "project info")
    hascollision = gMemoryMap.AddSector(gProjectInfo.aObjectTable - OptHeader.ImageBase, 84, "objecttable")
    hascollision = gMemoryMap.AddSector(gVBHeader.aComRegisterData - OptHeader.ImageBase, Len(modGlobals.gCOMRegData), "ComRegisterData")

    Dim i As Integer
    For i = 0 To gObjectTable.ObjectCount1
    
    Next
    
    gMemoryMap.ExportToHTML 'exports to File.Name & ".html"
    Me.Caption = strTitle
    MsgBox "Memory Map Created!"

End Sub

Private Sub mnuFileGenerate_Click()
'*****************************
'Purpose: To generate all the vb files from the decompiled exe.
'*****************************
    Dim sPath As String
    Dim structFolder As BROWSEINFO
    Dim iNull As Integer
    Dim ret As Long
    structFolder.hOwner = Me.hwnd
    structFolder.lpszTitle = "Browse for folder"
    structFolder.ulFlags = BIF_NEWDIALOGSTYLE  'To create make new folder option
    'BIF_RETURNONLYFSDIRS &
   'structFolder.ulFlags = &H40
    
    
    ret = SHBrowseForFolder(structFolder)
    If ret Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList ret, sPath
        'free the block of memory
        CoTaskMemFree ret
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    
    If sPath = "" Then Exit Sub
    
    'Write The Project File
    Call WriteVBP(sPath & "\" & ProjectName & ".vbp")
    'Write the forms
    Call WriteForms(sPath & "\")
    'Write Forms frx files
    For i = 0 To UBound(gObject)
        If gObject(i).ObjectType = 98435 Then
            Call modOutput.WriteFormFrx(sPath, gObjectNameArray(i))
        End If
    Next
    'Write the modules
    For i = 0 To UBound(gObject)
        If gObject(i).ObjectType = 98305 Then
           Call modOutput.WriteModules(sPath & "\" & gObjectNameArray(i) & ".bas", gObjectNameArray(i))
        End If
    Next
    'Write the classes
    For i = 0 To UBound(gObject)
        If gObject(i).ObjectType = 1146883 Then
            Call modOutput.WriteClasses(sPath & "\" & gObjectNameArray(i) & ".cls", gObjectNameArray(i))
        End If
    Next
    'Write the user controls
    
    MsgBox "Done"
End Sub

Private Sub mnuFileOpen_Click()
'*****************************
'Purpose: Show Open Dialog and then call OpenVBExe
'*****************************
    Cd1.Filename = ""
    Cd1.DialogTitle = "Select VB5/VB6 exe"
    Cd1.Filter = "VB Files(*.exe,*.ocx,*.dll)|*.exe;*.ocx;*.dll|All Files(*.*)|*.*;"
    Cd1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
    Cd1.ShowOpen
    
    If Cd1.Filename = "" Then Exit Sub
    
    If FileExists(Cd1.Filename) = True Then
        Call OpenVBExe(Cd1.Filename, Cd1.FileTitle)
    Else
        MsgBox "File Does not exist"
    End If
End Sub

Sub OpenVBExe(FilePath As String, FileTitle As String)
'################################################
'Purpose: Main function that gets all VB Sturtures
'#################################################
 Dim bFormEndUsed As Boolean
 Dim i As Integer 'Loop Var
 Dim k As Integer 'Loop Var
 Dim addr As Integer 'Loop Var
 Dim StartOffset As Long 'Holds Address of first VB Struture
 Dim f As Integer 'FileNumber holder
 
    'Erase existing data
    bFormEndUsed = False
    For i = 0 To txtFinal.UBound
        txtFinal(i).Text = ""
        txtFinal(i).Tag = ""
    Next
    mnuFileGenerate.Enabled = False
    mnuFileExportMemoryMap.Enabled = False
    mnuFileAntiDecompiler.Enabled = False
    SFilePath = ""
    SFile = ""
    ReDim gControlNameArray(0) 'Treeveiw control list
    ReDim gProcedureList(0)
    ReDim gOcxList(0)
    ReDim FrxPreview(0)
    'Reset Change Types
    ReDim ByteChange(0)
    ReDim BooleanChange(0)
    ReDim IntegerChange(0)
    ReDim LongChange(0)
    ReDim SingleChange(0)
    ReDim StringChange(0)
    'Pcode
    ReDim EventProcList(0)
    ReDim SubNamelist(0)
    'clear the nodes
    tvProject.Nodes.Clear
    'Save name and path
    SFilePath = FilePath
    SFile = FileTitle
    
    'Reset the error flag
    ErrorFlag = False
    CancelDecompile = False
    'Get a file handle
    InFileNumber = FreeFile
    
    'Check for error
    'On Error GoTo AnalyzeError
    
    'Access the file
    Open SFilePath For Binary As #InFileNumber
       
    'Is it a VB6 file?
    If CheckHeader() = True Then
        'Good file
        
        Close #InFileNumber
    Else
       'Bad file
        MsgBox "Not a VB6 file.", vbOKOnly Or vbCritical Or vbApplicationModal, "Bad file!"
        Close #InFileNumber
        Exit Sub
    End If

    StartOffset = VBStartHeader.PushStartAddress - OptHeader.ImageBase
  
    MakeDir (App.path & "\dump")
    MakeDir (App.path & "\dump\" & FileTitle)

   'Setup the VB File class
    Set gVBFile = New clsFile
    Call gVBFile.Setup(SFilePath)
    f = gVBFile.FileNumber
        'Goto begining of vb header
        Seek f, StartOffset + 1
        'Get the vb header
        Get #f, , gVBHeader
        
        AppData.FormTableAddress = gVBHeader.aGUITable
        'GetHelpFile
        Seek #f, StartOffset + 1 + gVBHeader.oHelpFile 'Loc(f) + gVBHeader.oHelpFile + 1
      
        HelpFile = GetUntilNull(f)
       
        'Get Project Name
        Seek #f, StartOffset + 1 + gVBHeader.oProjectName
        ProjectName = GetUntilNull(f)
        'Project Title
        Seek #f, StartOffset + 1 + gVBHeader.oProjectTitle
        ProjectTitle = GetUntilNull(f)
        'ExeName
        Seek #f, StartOffset + 1 + gVBHeader.oProjectExename
        ProjectExename = GetUntilNull(f)
        'Get ComRegisterData
        Seek #f, gVBHeader.aComRegisterData + 1 - OptHeader.ImageBase
        Get #f, , gCOMRegData
        Get #f, , gCOMRegInfo
        
        'Get ProjectDescription
        Seek #f, gVBHeader.aComRegisterData + 1 + gCOMRegData.oNTSProjectDescription - OptHeader.ImageBase
        ProjectDescription = GetUntilNull(f)

        
        'Get External Componetns
        '##########
        If gVBHeader.ExternalComponentCount > 0 Then
        Seek f, gVBHeader.aExternalComponentTable + 1 - OptHeader.ImageBase
            'MsgBox gVBHeader.aExternalComponentTable + 1 - OptHeader.ImageBase
            ReDim gOcxList(0)
            Dim AexternEnd As Long
            Dim bExternEnd As Long
            For i = 1 To gVBHeader.ExternalComponentCount
               bExternEnd = Loc(f)
               Dim cOcx As tComponent
               Get f, , cOcx
               AexternEnd = bExternEnd + 1 + cOcx.StructLength
               If cOcx.GUIDlength = 72 Then
                    Seek f, bExternEnd + 1 + cOcx.GUIDoffset
                    gOcxList(UBound(gOcxList)).strGuid = UCase(GetUnicodeString(f, 36))
               End If
               Seek f, bExternEnd + 1 + cOcx.FileNameOffset
               gOcxList(UBound(gOcxList)).strocxName = GetUntilNull(f)
               Seek f, bExternEnd + 1 + cOcx.SourceOffset
               gOcxList(UBound(gOcxList)).strLibName = GetUntilNull(f)
               Seek f, bExternEnd + 1 + cOcx.NameOffset
               gOcxList(UBound(gOcxList)).strName = GetUntilNull(f)
               ReDim Preserve gOcxList(UBound(gOcxList) + 1)
               Seek f, AexternEnd
            Next
        End If
        
        'Get Project Info Table
        Seek f, gVBHeader.aProjectInfo + 1 - OptHeader.ImageBase
        Get #f, , gProjectInfo
        
        'Begin Main Loop to get api list
        Dim nApi As Integer
        ReDim gApiList(0)
        For nApi = 0 To gProjectInfo.ExternalCount - 1
          'Get External Table 'Number of Api Calls
           Seek f, gProjectInfo.aExternalTable + 1 + (nApi * 8) - OptHeader.ImageBase
          Get #f, , gExternalTable
     
          'Get External Library
          If gProjectInfo.ExternalCount > 0 And gExternalTable.flag <> 6 Then
              Seek f, gExternalTable.aExternalLibrary + 1 - OptHeader.ImageBase
              Get #f, , gExternalLibrary
              
              If gExternalLibrary.aLibraryFunction <> 0 Then
                Seek f, gExternalLibrary.aLibraryFunction + 1 - OptHeader.ImageBase
                gApiList(UBound(gApiList)).strFunctionName = GetUntilNull(f)
                Seek f, gExternalLibrary.aLibraryName + 1 - OptHeader.ImageBase
                gApiList(UBound(gApiList)).strLibraryName = GetUntilNull(f)
                ReDim Preserve gApiList(UBound(gApiList) + 1)
              End If
          End If
          
        Next nApi 'End Api List Loop
    
        'Get Object Table
        Seek f, gProjectInfo.aObjectTable + 1 - OptHeader.ImageBase
        Get #f, , gObjectTable
        
        'Resize for the number of objects...(forms,modules,classes)
        ReDim gObject(gObjectTable.ObjectCount1 - 1)
        ReDim gObjectNameArray(gObjectTable.ObjectCount1 - 1)
        ReDim gObjectProcCountArray(gObjectTable.ObjectCount1 - 1)
        ReDim gObjectInfoHolder(gObjectTable.ObjectCount1 - 1)
        'Get Object
        Seek f, gObjectTable.aObject + 1 - OptHeader.ImageBase
        Get #f, , gObject
       
        
        Dim loopC As Integer
        For loopC = 0 To UBound(gObject)
        'Get ObjectName
        Seek f, gObject(loopC).aObjectName + 1 - OptHeader.ImageBase
        gObjectNameArray(loopC) = GetUntilNull(f)
        gObjectProcCountArray(loopC) = gObject(loopC).ProcCount
        
        'Get Object Info
        Seek f, gObject(loopC).aObjectInfo + 1 - OptHeader.ImageBase
        Get #f, , gObjectInfo
        'Save the information for later on
        gObjectInfoHolder(loopC).aConstantPool = gObjectInfo.aConstantPool
        gObjectInfoHolder(loopC).aObject = gObjectInfo.aObject
        gObjectInfoHolder(loopC).aObjectTable = gObjectInfo.aObjectTable
        gObjectInfoHolder(loopC).aProcTable = gObjectInfo.aProcTable
        gObjectInfoHolder(loopC).aSmallRecord = gObjectInfo.aSmallRecord
        gObjectInfoHolder(loopC).Const1 = gObjectInfo.Const1
        gObjectInfoHolder(loopC).Flag1 = gObjectInfo.Flag1
        gObjectInfoHolder(loopC).iConstantsCount = gObjectInfo.iConstantsCount
        gObjectInfoHolder(loopC).iMaxConstants = gObjectInfo.iMaxConstants
        gObjectInfoHolder(loopC).Flag5 = gObjectInfo.Flag5
        gObjectInfoHolder(loopC).Flag6 = gObjectInfo.Flag6
        gObjectInfoHolder(loopC).Flag7 = gObjectInfo.Flag7
        gObjectInfoHolder(loopC).Null1 = gObjectInfo.Null1
        gObjectInfoHolder(loopC).Null2 = gObjectInfo.Null2
        gObjectInfoHolder(loopC).NumberOfProcs = gObjectInfo.NumberOfProcs
        gObjectInfoHolder(loopC).ObjectIndex = gObjectInfo.ObjectIndex
        gObjectInfoHolder(loopC).RunTimeLoaded = gObjectInfo.RunTimeLoaded
        
        'If gObjectInfo.aProcTable - OptHeader.ImageBase > 0 Then
            'Dim ProcCodeInfo As tCodeInfo
            'Seek f, gObjectInfo.aProcTable + 1 - OptHeader.ImageBase
            'Get f, , ProcCodeInfo
        'End If
        'If gObjectInfo.aConstantPool <> 0 Then
            'Seek f, gObjectInfo.aConstantPool + 1 - OptHeader.ImageBase
       ' End If
        
         'Get Optional Object Info
        Seek f, gObject(loopC).aObjectInfo + 57 - OptHeader.ImageBase
        
        'Decide if to get Optional Info or not
        If ((gObject(loopC).ObjectType And &H80) = &H80) Then
            
            Get #f, , gOptionalObjectInfo
            'Dim testLink() As tEventLink
            Dim LinkPCode() As MethodLinkPCode
            Dim LinkNative() As MethodLinkNative
            
            'Resize the Arrays
            ReDim LinkPCode(gOptionalObjectInfo.iEventCount - 1)
            ReDim LinkNative(gOptionalObjectInfo.iEventCount - 1)
            
            'MsgBox gOptionalObjectInfo.iEventCount
            If gOptionalObjectInfo.aEventLinkArray <> 0 And gOptionalObjectInfo.aEventLinkArray <> -1 Then
                If gOptionalObjectInfo.aEventLinkArray + 1 - OptHeader.ImageBase > 0 Then
                    Seek f, gOptionalObjectInfo.aEventLinkArray + 1 - OptHeader.ImageBase
                    If gProjectInfo.aNativeCode = 0 Then
                    'P-Code
                        Get f, , LinkPCode
                    Else
                    'Native
                        Get f, , LinkNative
                    End If
                    
                    
                    'For i = 0 To UBound(LinkPCode)
                       ' MsgBox LinkPCode(i).movAddress '+ 1 - OptHeader.ImageBase
                    'Next
                End If
            End If
        End If
        'Address PublicBytes
        'Notes aPublicBytes points to a structure of 2 integers (iStringBytes and iVarBytes) and this structure tells how many pointers will be in memory at aModulePublic.
        If gObject(loopC).aPublicBytes <> 0 Then
            Seek #f, gObject(loopC).aPublicBytes + 1 - OptHeader.ImageBase
            Dim iStringBytes As Integer, iVarBytes As Integer
            Get f, , iStringBytes
            Get f, , iVarBytes
            'MsgBox "StringBytes: " & iStringBytes & " VarBytes: " & iVarBytes
            If gObject(loopC).aModulePublic <> 0 Then
                Seek #f, gObject(loopC).aModulePublic + 1 - OptHeader.ImageBase
              '  MsgBox gObject(loopC).aModulePublic + 1 - OptHeader.ImageBase
            End If
        End If
        
        'Resize the control array
        'Check if its a form
        If gObject(loopC).ObjectType = 98435 And gOptionalObjectInfo.ControlCount < 5000 And gOptionalObjectInfo.ControlCount <> 0 Then
            ReDim gControl(gOptionalObjectInfo.ControlCount - 1)
            'Get Control Array
            Seek f, gOptionalObjectInfo.aControlArray + 1 - OptHeader.ImageBase
            Get #f, , gControl
            'Resize Event Table array
            ReDim gEventTable(UBound(gControl))

            Dim ControlName As String
            
            For i = 0 To UBound(gControl)
                'Get Event Table
               Seek f, gControl(i).aEventTable + 1 - OptHeader.ImageBase
               ' ReDim gEventTable(i).aEventPointer(gControl(i).EventCount - 1)
                ReDim taEventPointer(gControl(i).EventCount - 1)
                'MsgBox gOptionalObjectInfo.iEventCount & " " & gControl(i).EventCount
                Get #f, , gEventTable(i)
                Get #f, , taEventPointer
      
                If gControl(i).aName + 1 - OptHeader.ImageBase > 0 Then
                 Seek f, gControl(i).aName + 1 - OptHeader.ImageBase
                 ControlName = GetUntilNull(f)
                 Dim strGuid As String
                 Seek f, gControl(i).aGUID + 1 - OptHeader.ImageBase
                 strGuid = modGlobals.ReturnGuid(f)

                For k = 0 To UBound(taEventPointer)
                    If taEventPointer(k) <> 0 Then
                    '  MsgBox "Good:" & ControlName & " " & taEventPointer(k) + 1 - OptHeader.ImageBase & " #" & k
                       'MsgBox "Offset: " & taEventPointer(k) + 1 - OptHeader.ImageBase
                        Dim pointerAevent As tEventPointer
                        Seek f, taEventPointer(k) + 1 - OptHeader.ImageBase
                        Get f, , pointerAevent
                        If pointerAevent.aEvent <> 0 Then
                       ' MsgBox getEventComplete(App.path & "\VB6.OLB", strGuid, Int(k) + 1)
                            SubNamelist(UBound(SubNamelist)).strName = gObjectNameArray(loopC) & "." & ControlName & "_Event"
                            SubNamelist(UBound(SubNamelist)).offset = pointerAevent.aEvent
                            ReDim Preserve SubNamelist(UBound(SubNamelist) + 1)
                            EventProcList(UBound(EventProcList)) = pointerAevent.aEvent 'taEventPointer(k)
                            ReDim Preserve EventProcList(UBound(EventProcList) + 1)
                        End If
                    End If

                Next
                 

                 'Save the control information for the treeview
                 ReDim Preserve gControlNameArray(UBound(gControlNameArray) + 1)
                 gControlNameArray(UBound(gControlNameArray)).strControlName = ControlName
                 gControlNameArray(UBound(gControlNameArray)).strParentForm = gObjectNameArray(loopC)
                 gControlNameArray(UBound(gControlNameArray)).strGuid = strGuid
                End If
            Next
        
        End If
        
        If gObject(loopC).ProcCount <> 0 Then

            If gObject(loopC).aProcNamesArray <> 0 Then
            Dim AddressProcNamesArray() As Long
            ReDim AddressProcNamesArray(gObject(loopC).ProcCount - 1)
            
            
            Seek f, gObject(loopC).aProcNamesArray + 1 - OptHeader.ImageBase
            Get f, , AddressProcNamesArray
   
                For addr = 0 To UBound(AddressProcNamesArray)
                   
                    If AddressProcNamesArray(addr) = 0 Then
                    
                    Else
                        If (AddressProcNamesArray(addr) - OptHeader.ImageBase) < 0 Then
                            
                        Else
                            Seek f, AddressProcNamesArray(addr) + 1 - OptHeader.ImageBase
                      
                           
                            gProcedureList(UBound(gProcedureList)).strProcedureName = GetUntilNull(f)
                            gProcedureList(UBound(gProcedureList)).strParent = gObjectNameArray(loopC)
                            SubNamelist(UBound(SubNamelist)).strName = gProcedureList(UBound(gProcedureList)).strProcedureName
                            SubNamelist(UBound(SubNamelist)).offset = AddressProcNamesArray(addr)
                            ReDim Preserve SubNamelist(UBound(SubNamelist) + 1)
                            
                            ReDim Preserve gProcedureList(UBound(gProcedureList) + 1)
                        End If
                    End If
                Next
         
            End If

        
        End If
        Next loopC

        'Main Loop to Get all Form's Properties
        FrameStatus.Visible = True
        txtStatus.Text = ""
        Call ProccessControls(f)
        
    Close f

    
    'Set the compile type either pcode or ncode
    If gProjectInfo.aNativeCode <> 0 Then
        AppData.CompileType = "Native"
        'Begin Native Decompile
        Call modNative.Decode(SFilePath)
    Else
        AppData.CompileType = "PCode"
        'Begin Pcode Decompile
        txtStatus.Text = txtStatus.Text & "Begin PCode Decompile" & vbCrLf
        Call modPCode4.Init
        txtStatus.Text = txtStatus.Text & "End PCode Decompile" & vbCrLf
        'Decompile the file
        If gPcodeDecompile = True Then
            Call modPCode4.Decode(SFilePath)
        End If
    End If
    
    

    mnuFileGenerate.Enabled = True
    mnuFileExportMemoryMap.Enabled = True
   ' mnuFileAntiDecompiler.Enabled = True
    'Get FileVersion Info
    gFileInfo = modGlobals.FileInfo(SFilePath)
    
    'Hide Form Generation Status
    FrameStatus.Visible = False
    
    Call SetupTreeView
    Call modOutput.DumpVBExeInfo(App.path & "\dump\" & FileTitle & "\FileReport.txt", FileTitle)
    
    'Add to recent files
    Call AddToRecentList(SFilePath, SFile)

    'Clear current data
    DDirPath = ""
    DFile = ""

  Exit Sub
    
AnalyzeError:

    MsgBox "Analyze error", vbCritical Or vbOKOnly, "Source file error"

End Sub
Sub AddToRecentList(Filename As String, FileTitle As String)
'*****************************
'Purpose: Add a Filename to the recently access list via the registry
'*****************************
    Dim Recent1File As String
    Dim Recent1Title As String
    Dim Recent2File As String
    Dim Recent2Title As String
    Dim Recent3File As String
    Dim Recent3Title As String
    
    mnuFileSep1.Visible = True
    mnuFileRecent1.Visible = True
    
    Recent1Title = GetSetting("VB Decompiler", "Options", "Recent1FileTitle", "")
    Recent2Title = GetSetting("VB Decompiler", "Options", "Recent2FileTitle", "")
    Recent3Title = GetSetting("VB Decompiler", "Options", "Recent3FileTitle", "")
    Recent1File = GetSetting("VB Decompiler", "Options", "Recent1File", "")
    Recent2File = GetSetting("VB Decompiler", "Options", "Recent2File", "")
    Recent3File = GetSetting("VB Decompiler", "Options", "Recent3File", "")
    
   
    
    If Recent1Title <> "" Then
        mnuFileRecent2.Visible = True
    End If
    If Recent2Title <> "" Then
        mnuFileRecent3.Visible = True
    End If
    If Recent3Title <> "" Then
        mnuFileRecent4.Visible = True
    End If


    Call SaveSetting("VB Decompiler", "Options", "Recent4File", Recent3File)
    Call SaveSetting("VB Decompiler", "Options", "Recent4FileTitle", Recent3Title)
    Call SaveSetting("VB Decompiler", "Options", "Recent3File", Recent2File)
    Call SaveSetting("VB Decompiler", "Options", "Recent3FileTitle", Recent2Title)
    Call SaveSetting("VB Decompiler", "Options", "Recent2File", Recent1File)
    Call SaveSetting("VB Decompiler", "Options", "Recent2FileTitle", Recent1Title)


    Call SaveSetting("VB Decompiler", "Options", "Recent1File", Filename)
    Call SaveSetting("VB Decompiler", "Options", "Recent1FileTitle", FileTitle)
    
    
    
    mnuFileRecent4.Caption = mnuFileRecent3.Caption
    mnuFileRecent3.Caption = mnuFileRecent2.Caption
    mnuFileRecent2.Caption = mnuFileRecent1.Caption
    mnuFileRecent1.Caption = FileTitle
    

End Sub
Sub MakeDir(path As String)
'*****************************
'Purpose: To make a dir without erroring
'*****************************

On Error Resume Next
    MkDir (path)

End Sub

Private Sub mnuFileRecent1_Click()
'*****************************
'Purpose: To load a recent file if it exists
'*****************************
Dim RecentTitle As String
Dim RecentFile As String
    RecentTitle = GetSetting("VB Decompiler", "Options", "Recent1FileTitle", "")
    RecentFile = GetSetting("VB Decompiler", "Options", "Recent1File", "")
    If FileExists(RecentFile) = True Then
        Call OpenVBExe(RecentFile, RecentTitle)
    Else
        MsgBox "File no longer exists!", vbExclamation
    End If
End Sub

Private Sub mnuFileRecent2_Click()
'*****************************
'Purpose: To load a recent file if it exists
'*****************************
Dim RecentTitle As String
Dim RecentFile As String
    RecentTitle = GetSetting("VB Decompiler", "Options", "Recent2FileTitle", "")
    RecentFile = GetSetting("VB Decompiler", "Options", "Recent2File", "")
    If FileExists(RecentFile) = True Then
        Call OpenVBExe(RecentFile, RecentTitle)
    Else
        MsgBox "File no longer exists!", vbExclamation
    End If
End Sub

Private Sub mnuFileRecent3_Click()
'*****************************
'Purpose: To load a recent file if it exists
'*****************************
Dim RecentTitle As String
Dim RecentFile As String
    RecentTitle = GetSetting("VB Decompiler", "Options", "Recent3FileTitle", "")
    RecentFile = GetSetting("VB Decompiler", "Options", "Recent3File", "")
    If FileExists(RecentFile) = True Then
        Call OpenVBExe(RecentFile, RecentTitle)
    Else
        MsgBox "File no longer exists!", vbExclamation
    End If
End Sub

Private Sub mnuFileRecent4_Click()
'*****************************
'Purpose: To load a recent file if it exists
'*****************************
Dim RecentTitle As String
Dim RecentFile As String
    RecentTitle = GetSetting("VB Decompiler", "Options", "Recent4FileTitle", "")
    RecentFile = GetSetting("VB Decompiler", "Options", "Recent4File", "")
    If FileExists(RecentFile) = True Then
        Call OpenVBExe(RecentFile, RecentTitle)
    Else
        MsgBox "File no longer exists!", vbExclamation
    End If
End Sub

Private Sub mnuFileSaveExe_Click()
'#####################################
'Purpose: Save Changes to the Form's Gui
'And generates a Patch Report
'#####################################
    Cd1.DialogTitle = "Save As"
    Cd1.Filename = ""
    Cd1.Filter = "VB Files(*.exe,*.ocx,*.dll)|*.exe;*.ocx;*.dll"
    Cd1.ShowSave
    
    If Cd1.Filename = "" Then Exit Sub
    On Error Resume Next
    'Copy the exe to the temp directory
    FileCopy SFilePath, App.path & "\dump\" & SFile & "\" & SFile
    
    'Make the changes
    fFile = FreeFile
    Dim i As Integer
    Dim NewByte As Byte
    Open App.path & "\dump\" & SFile & "\" & SFile For Binary Access Write Lock Write As fFile
        If UBound(StringChange) > 0 Then
            For i = 1 To UBound(StringChange)
                Seek fFile, StringChange(i).offset '+ 1
                Dim bArray() As Byte
                ReDim bArray(Len(StringChange(i).sString))
                For g = 0 To Len(StringChange(i).sString)
                    bArray(g) = Asc(Mid(StringChange(i).sString, 1 + g, 1))
                Next g
                Put fFile, , bArray
                'Put fFile, , StringChange(I).sString
            Next
        End If
        If UBound(ByteChange) > 0 Then
            For i = 1 To UBound(ByteChange)
                Seek fFile, ByteChange(i).offset
                Put fFile, , ByteChange(i).bByte
            Next
        End If
        
        If UBound(BooleanChange) > 0 Then
            For i = 1 To UBound(BooleanChange)
                Seek fFile, BooleanChange(i).offset
                If BooleanChange(i).bBool = True Then
                    NewByte = 255
                    Put fFile, , NewByte
                Else
                    NewByte = 0
                    Put fFile, , NewByte
                End If
                'Put fFile, , ByteChange(i).bByte
            Next i
        End If
        If UBound(IntegerChange) > 0 Then
            For i = 1 To UBound(IntegerChange)
                Seek fFile, IntegerChange(i).offset
                Put fFile, , IntegerChange(i).iInt
            Next
        End If
        If UBound(LongChange) > 0 Then
            For i = 1 To UBound(LongChange)
                Seek fFile, LongChange(i).offset
                Put fFile, , LongChange(i).lLong
            Next
        End If
        If UBound(SingleChange) > 0 Then
            For i = 1 To UBound(SingleChange)
                Seek fFile, SingleChange(i).offset
                Put fFile, , SingleChange(i).sSingle
            Next
        End If
        
    Close fFile
    
    'Save the file
    FileCopy App.path & "\dump\" & SFile & "\" & SFile, Cd1.Filename
    'Kill the temp file
    Kill App.path & "\dump\" & SFile & "\" & SFile
    
    'Write Patch Report
    
    fFile = FreeFile
   
    Open App.path & "\dump\" & SFile & "\PatchReport.txt" For Output As fFile
        Print #fFile, "File Patch Report from Semi VB Decompiler by vbgamer45"
        Print #fFile, "------------------------------------------------------"
        Print #fFile, "Filename=" & SFile
        Print #fFile, ""
        Print #fFile, "Byte Changes"
        For i = 0 To UBound(ByteChange)
                If i <> 0 Then
                    Print #fFile, "Offset:" & ByteChange(i).offset & " Changed to: " & ByteChange(i).bByte
                End If
        Next i
        Print #fFile, ""
        Print #fFile, "Boolean Changes"
        For i = 0 To UBound(BooleanChange)
                If i <> 0 Then
                    Print #fFile, "Offset:" & BooleanChange(i).offset & " Changed to: " & BooleanChange(i).bBool
                End If
        Next i
        Print #fFile, ""
        Print #fFile, "Integer Changes"
        For i = 0 To UBound(IntegerChange)
                If i <> 0 Then
                    Print #fFile, "Offset:" & IntegerChange(i).offset & " Changed to: " & IntegerChange(i).iInt
                End If
        Next i
        Print #fFile, ""
        Print #fFile, "Long Changes"
        For i = 0 To UBound(LongChange)
                If i <> 0 Then
                    Print #fFile, "Offset:" & LongChange(i).offset & " Changed to: " & LongChange(i).lLong
                End If
        Next i
        Print #fFile, ""
        Print #fFile, "Single Changes"
        For i = 0 To UBound(SingleChange)
                If i <> 0 Then
                    Print #fFile, "Offset:" & SingleChange(i).offset & " Changed to: " & SingleChange(i).sSingle
                End If
        Next i
        Print #fFile, ""
        Print #fFile, "String Changes"
        For i = 0 To UBound(StringChange)
                If i <> 0 Then
                    Print #fFile, "Offset:" & StringChange(i).offset & " Changed to: " & StringChange(i).sString
                End If
        Next i
        
    Close fFile
    
    MsgBox "Done"
End Sub

Private Sub mnuHelpAbout_Click()
'*****************************
'Purpose: Show my Cool about screen.
'*****************************
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuOptions_Click()
'*****************************
'Purpose: Show the options form
'*****************************
    frmOptions.Show vbModal, Me
    
End Sub
Private Sub mnuToolsPCodeProcedure_Click()
'*****************************
'Purpose: Show the Procedure Decompile View
'*****************************
    If SFilePath = "" Then
        MsgBox "No File Loaded"
        Exit Sub
    End If
    If modGlobals.gProjectInfo.aNativeCode = 0 Then
        frmPcode.Show vbModal, Me
    Else
        MsgBox "This is a Native compiled exe! Not a P-Code one!", vbExclamation
    End If
End Sub

Private Sub tvProject_NodeClick(ByVal Node As MSComctlLib.Node)
'*****************************
'Purpose: To show the contents of each struture and textbox data
'*****************************
On Error Resume Next
 Dim ParentObject As Node
    Dim LenTab As Long
    Dim i As Long, o As Long
    Dim strCode As String
    
    Dim tblPath() As String
    txtCode.SelStart = 0
    txtCode.SelColor = vbBlack
    
    If CurrentItem <> tvProject.SelectedItem.Key Then
        tblPath = Split(tvProject.SelectedItem.Key, "/")
        CurrentItem = tvProject.SelectedItem.Key

        Select Case tblPath(1)
            Case "VERSIONINFO"
                sstViewFile.TabVisible(1) = True
                sstViewFile.TabVisible(0) = False
                sstViewFile.TabVisible(2) = False
                sstViewFile.TabVisible(3) = False
                fxgEXEInfo.Visible = True
                fxgEXEInfo.ColAlignment(1) = 0
                fxgEXEInfo.Clear
                fxgEXEInfo.Rows = 2
                fxgEXEInfo.ColWidth(0) = 2500
                fxgEXEInfo.TextArray(0) = "Name"
                fxgEXEInfo.TextArray(1) = "Value"
                        fxgEXEInfo.ColWidth(0) = 2000
                        fxgEXEInfo.ColWidth(1) = 2500
                        fxgEXEInfo.TextArray(2) = "CompanyName"
                        fxgEXEInfo.TextArray(3) = gFileInfo.CompanyName
                        fxgEXEInfo.AddItem "FileDescription"
                        fxgEXEInfo.TextArray(5) = gFileInfo.FileDescription
                        fxgEXEInfo.AddItem "FileVersion"
                        fxgEXEInfo.TextArray(7) = gFileInfo.FileVersion
                        fxgEXEInfo.AddItem "InternalName"
                        fxgEXEInfo.TextArray(9) = gFileInfo.InternalName
                        fxgEXEInfo.AddItem "LanguageID"
                        fxgEXEInfo.TextArray(11) = gFileInfo.LanguageID
                        fxgEXEInfo.AddItem "LegalCopyright"
                        fxgEXEInfo.TextArray(13) = gFileInfo.LegalCopyright
                        fxgEXEInfo.AddItem "OrigionalFileName"
                        fxgEXEInfo.TextArray(15) = gFileInfo.OrigionalFileName
                        fxgEXEInfo.AddItem "ProductName"
                        fxgEXEInfo.TextArray(17) = gFileInfo.ProductName
                        fxgEXEInfo.AddItem "ProductVersion"
                        fxgEXEInfo.TextArray(19) = gFileInfo.ProductVersion
            Case "STRUCT"
                sstViewFile.TabVisible(1) = True
                sstViewFile.TabVisible(0) = False
                sstViewFile.TabVisible(2) = False
                sstViewFile.TabVisible(3) = False
                fxgEXEInfo.Visible = True
                fxgEXEInfo.ColAlignment(1) = 0
                fxgEXEInfo.Clear
                fxgEXEInfo.Rows = 2
                fxgEXEInfo.ColWidth(0) = 2500
                fxgEXEInfo.TextArray(0) = "Name"
                fxgEXEInfo.TextArray(1) = "Value"
                Select Case tblPath(2)
                    Case "", "VBHEADER"
                        fxgEXEInfo.ColWidth(0) = 2500
                        fxgEXEInfo.TextArray(2) = "Signature"
                        fxgEXEInfo.TextArray(3) = gVBHeader.Signature
                        fxgEXEInfo.AddItem "Address SubMain"
                        fxgEXEInfo.TextArray(5) = gVBHeader.aSubMain
                        fxgEXEInfo.AddItem "Address ExternalComponentTable"
                        fxgEXEInfo.TextArray(7) = gVBHeader.aExternalComponentTable
                        fxgEXEInfo.AddItem "Address GUITable"
                        fxgEXEInfo.TextArray(9) = gVBHeader.aGUITable
                        fxgEXEInfo.AddItem "Address ComRegisterData"
                        fxgEXEInfo.TextArray(11) = gVBHeader.aComRegisterData
                        fxgEXEInfo.AddItem "Address ProjectInfo"
                        fxgEXEInfo.TextArray(13) = gVBHeader.aProjectInfo
                        fxgEXEInfo.AddItem "BackupLanguageDLL"
                        fxgEXEInfo.TextArray(15) = gVBHeader.BackupLanguageDLL
                        fxgEXEInfo.AddItem "BackupLanguageID"
                        fxgEXEInfo.TextArray(17) = gVBHeader.BackupLanguageID
                        fxgEXEInfo.AddItem "ExternalComponentCount"
                        fxgEXEInfo.TextArray(19) = gVBHeader.ExternalComponentCount
                        fxgEXEInfo.AddItem "Flag MDLIntObjs"
                        fxgEXEInfo.TextArray(21) = gVBHeader.fMDLIntObjs
                        fxgEXEInfo.AddItem "Flag MDLIntObjs2"
                        fxgEXEInfo.TextArray(23) = gVBHeader.fMDLIntObjs2
                        fxgEXEInfo.AddItem "FormCount"
                        fxgEXEInfo.TextArray(25) = gVBHeader.FormCount
                        fxgEXEInfo.AddItem "LanguageDLL"
                        fxgEXEInfo.TextArray(27) = gVBHeader.LanguageDLL
                        fxgEXEInfo.AddItem "LanguageID"
                        fxgEXEInfo.TextArray(29) = gVBHeader.LanguageID
                        fxgEXEInfo.AddItem "Offset HelpFile"
                        fxgEXEInfo.TextArray(31) = gVBHeader.oHelpFile
                        fxgEXEInfo.AddItem "Offset ProjectExename"
                        fxgEXEInfo.TextArray(33) = gVBHeader.oProjectExename
                        fxgEXEInfo.AddItem "Offset ProjectName"
                        fxgEXEInfo.TextArray(35) = gVBHeader.oProjectName
                        fxgEXEInfo.AddItem "Offset ProjectTitle"
                        fxgEXEInfo.TextArray(37) = gVBHeader.oProjectTitle
                        fxgEXEInfo.AddItem "RuntimeDLLVersion"
                        fxgEXEInfo.TextArray(39) = gVBHeader.RuntimeDLLVersion
                        fxgEXEInfo.AddItem "RuntimeBuild"
                        fxgEXEInfo.TextArray(41) = gVBHeader.RuntimeBuild
                        fxgEXEInfo.AddItem "ThreadCount"
                        fxgEXEInfo.TextArray(43) = gVBHeader.ThreadCount
                        fxgEXEInfo.AddItem "ThreadFlags"
                        fxgEXEInfo.TextArray(45) = gVBHeader.ThreadFlags
                        fxgEXEInfo.AddItem "ThunkCount"
                        fxgEXEInfo.TextArray(47) = gVBHeader.ThunkCount
            
                    Case "VBPROJECTINFO"
                        fxgEXEInfo.ColWidth(0) = 2500
                        fxgEXEInfo.TextArray(2) = "Address EndOfCode"
                        fxgEXEInfo.TextArray(3) = gProjectInfo.aEndOfCode
                        fxgEXEInfo.AddItem "Address ExternalTable"
                        fxgEXEInfo.TextArray(5) = gProjectInfo.aExternalTable
                        fxgEXEInfo.AddItem "Address NativeCode"
                        fxgEXEInfo.TextArray(7) = gProjectInfo.aNativeCode
                        fxgEXEInfo.AddItem "Address ObjectTable"
                        fxgEXEInfo.TextArray(9) = gProjectInfo.aObjectTable
                        fxgEXEInfo.AddItem "Address StartOfCode"
                        fxgEXEInfo.TextArray(11) = gProjectInfo.aStartOfCode
                        fxgEXEInfo.AddItem "Address VBAExceptionhandler"
                        fxgEXEInfo.TextArray(13) = gProjectInfo.aVBAExceptionhandler
                        fxgEXEInfo.AddItem "ExternalCount"
                        fxgEXEInfo.TextArray(15) = gProjectInfo.ExternalCount
                        fxgEXEInfo.AddItem "Flag1"
                        fxgEXEInfo.TextArray(17) = gProjectInfo.Flag1
                        fxgEXEInfo.AddItem "Flag2"
                        fxgEXEInfo.TextArray(19) = gProjectInfo.Flag2
                        fxgEXEInfo.AddItem "Flag3"
                        fxgEXEInfo.TextArray(21) = gProjectInfo.Flag3
                        fxgEXEInfo.AddItem "Null1"
                        fxgEXEInfo.TextArray(23) = gProjectInfo.Null1
                        fxgEXEInfo.AddItem "NullSpacer"
                        fxgEXEInfo.TextArray(25) = gProjectInfo.NullSpacer
                        fxgEXEInfo.AddItem "oProjectLocation"
                        fxgEXEInfo.TextArray(27) = gProjectInfo.oProjectLocation
                        fxgEXEInfo.AddItem "OriginalPathName"
                        fxgEXEInfo.TextArray(29) = gProjectInfo.OriginalPathName
                        fxgEXEInfo.AddItem "Signature"
                        fxgEXEInfo.TextArray(31) = gProjectInfo.Signature
                        fxgEXEInfo.AddItem "ThreadSpace"
                        fxgEXEInfo.TextArray(33) = gProjectInfo.ThreadSpace
                    Case "VBCOMREGDATA"
                        fxgEXEInfo.ColWidth(0) = 2500
                        fxgEXEInfo.TextArray(2) = "TlbVerMajor"
                        fxgEXEInfo.TextArray(3) = gCOMRegData.iTlbVerMajor
                        fxgEXEInfo.AddItem "iTlbVerMinor"
                        fxgEXEInfo.TextArray(5) = gCOMRegData.iTlbVerMinor
                        fxgEXEInfo.AddItem "Padding1"
                        fxgEXEInfo.TextArray(7) = gCOMRegData.iPadding1
                        fxgEXEInfo.AddItem "Padding2"
                        fxgEXEInfo.TextArray(9) = gCOMRegData.iPadding2
                        fxgEXEInfo.AddItem "Padding3"
                        fxgEXEInfo.TextArray(11) = gCOMRegData.lPadding3
                        fxgEXEInfo.AddItem "lTlbLcid"
                        fxgEXEInfo.TextArray(13) = gCOMRegData.lTlbLcid
                        fxgEXEInfo.AddItem "Offset to NTSHelpDirectory"
                        fxgEXEInfo.TextArray(15) = gCOMRegData.oNTSHelpDirectory
                        fxgEXEInfo.AddItem "Offset to NTSProjectDescription"
                        fxgEXEInfo.TextArray(17) = gCOMRegData.oNTSProjectDescription
                        fxgEXEInfo.AddItem "Offset to NTSProjectName"
                        fxgEXEInfo.TextArray(19) = gCOMRegData.oNTSProjectName
                        fxgEXEInfo.AddItem "Offset to RegInfo"
                        fxgEXEInfo.TextArray(21) = gCOMRegData.oRegInfo
                        fxgEXEInfo.AddItem "uuidProjectClsId"
                        For i = 0 To UBound(gCOMRegData.uuidProjectClsId)
                            fxgEXEInfo.TextArray(23) = fxgEXEInfo.TextArray(23) & gCOMRegData.uuidProjectClsId(i)
                        Next
                    Case "VBCOMREGINFO"
                        fxgEXEInfo.ColWidth(0) = 2500
                        fxgEXEInfo.TextArray(2) = "fClassType"
                        fxgEXEInfo.TextArray(3) = gCOMRegInfo.fClassType
                        fxgEXEInfo.AddItem "fIsControl"
                        fxgEXEInfo.TextArray(5) = gCOMRegInfo.fIsControl
                        fxgEXEInfo.AddItem "fIsDesigner"
                        fxgEXEInfo.TextArray(7) = gCOMRegInfo.fIsDesigner
                        fxgEXEInfo.AddItem "fIsInterface"
                        fxgEXEInfo.TextArray(9) = gCOMRegInfo.fIsInterface
                        fxgEXEInfo.AddItem "fObjectType"
                        fxgEXEInfo.TextArray(11) = gCOMRegInfo.fObjectType
                        fxgEXEInfo.AddItem "iDefaultIcon"
                        fxgEXEInfo.TextArray(13) = gCOMRegInfo.iDefaultIcon
                        fxgEXEInfo.AddItem "iToolboxBitmap32"
                        fxgEXEInfo.TextArray(15) = gCOMRegInfo.iToolboxBitmap32
                        fxgEXEInfo.AddItem "lInstancing"
                        fxgEXEInfo.TextArray(17) = gCOMRegInfo.lInstancing
                        fxgEXEInfo.AddItem "lMiscStatus"
                        fxgEXEInfo.TextArray(19) = gCOMRegInfo.lMiscStatus
                        fxgEXEInfo.AddItem "lObjectID"
                        fxgEXEInfo.TextArray(21) = gCOMRegInfo.lObjectID
                        fxgEXEInfo.AddItem "Offset to ControlClsID"
                        fxgEXEInfo.TextArray(23) = gCOMRegInfo.oControlClsID
                        fxgEXEInfo.AddItem "Offset to DesignerData"
                        fxgEXEInfo.TextArray(25) = gCOMRegInfo.oDesignerData
                        fxgEXEInfo.AddItem "Offset to NextObject"
                        fxgEXEInfo.TextArray(27) = gCOMRegInfo.oNextObject
                        fxgEXEInfo.AddItem "Offset to ObjectClsID"
                        fxgEXEInfo.TextArray(29) = gCOMRegInfo.oObjectClsID
                        fxgEXEInfo.AddItem "Offset to ObjectDescription"
                        fxgEXEInfo.TextArray(31) = gCOMRegInfo.oObjectDescription
                        fxgEXEInfo.AddItem "Offset to ObjectName"
                        fxgEXEInfo.TextArray(33) = gCOMRegInfo.oObjectName
                        fxgEXEInfo.AddItem "uuidObjectClsID"
                        For i = 0 To UBound(gCOMRegInfo.uuidObjectClsID)
                            fxgEXEInfo.TextArray(35) = fxgEXEInfo.TextArray(35) & gCOMRegInfo.uuidObjectClsID(i)
                        Next
                    Case "VBOBJECTABLE"
                        fxgEXEInfo.ColWidth(0) = 2500
                        fxgEXEInfo.TextArray(2) = "Address of ExecProj"
                        fxgEXEInfo.TextArray(3) = gObjectTable.aExecProj
                        fxgEXEInfo.AddItem "Address of ProjectInfo2"
                        fxgEXEInfo.TextArray(5) = gObjectTable.aProjectInfo2
                        fxgEXEInfo.AddItem "Address of ProjectObject Size"
                        fxgEXEInfo.TextArray(7) = gObjectTable.lpProjectObject
                        fxgEXEInfo.AddItem "Address of First Object"
                        fxgEXEInfo.TextArray(9) = gObjectTable.aObject
                        fxgEXEInfo.AddItem "Address of ProjectName"
                        fxgEXEInfo.TextArray(11) = gObjectTable.aProjectName
                        fxgEXEInfo.AddItem "Const1"
                        fxgEXEInfo.TextArray(13) = gObjectTable.Const1
                        fxgEXEInfo.AddItem "Flag CompileType"
                        fxgEXEInfo.TextArray(15) = gObjectTable.fCompileType
                        fxgEXEInfo.AddItem "Const3"
                        fxgEXEInfo.TextArray(17) = gObjectTable.Const3
                        fxgEXEInfo.AddItem "Flag1"
                        fxgEXEInfo.TextArray(19) = gObjectTable.Flag1
                        fxgEXEInfo.AddItem "Flag2"
                        fxgEXEInfo.TextArray(21) = gObjectTable.Flag2
                        fxgEXEInfo.AddItem "Flag3"
                        fxgEXEInfo.TextArray(23) = gObjectTable.Flag3
                        fxgEXEInfo.AddItem "Flag4"
                        fxgEXEInfo.TextArray(25) = gObjectTable.Flag4
                        fxgEXEInfo.AddItem "LangID1"
                        fxgEXEInfo.TextArray(27) = gObjectTable.LangID1
                        fxgEXEInfo.AddItem "LangID2"
                        fxgEXEInfo.TextArray(29) = gObjectTable.LangID2
                        fxgEXEInfo.AddItem "Null1"
                        fxgEXEInfo.TextArray(31) = gObjectTable.lNull1
                        fxgEXEInfo.AddItem "Null2"
                        fxgEXEInfo.TextArray(33) = gObjectTable.Null2
                        fxgEXEInfo.AddItem "Null3"
                        fxgEXEInfo.TextArray(35) = gObjectTable.Null3
                        fxgEXEInfo.AddItem "Null4"
                        fxgEXEInfo.TextArray(37) = gObjectTable.Null4
                        fxgEXEInfo.AddItem "Null5"
                        fxgEXEInfo.TextArray(39) = gObjectTable.Null5
                        fxgEXEInfo.AddItem "Null6"
                        fxgEXEInfo.TextArray(41) = gObjectTable.Null6
                        fxgEXEInfo.AddItem "ObjectCount1"
                        fxgEXEInfo.TextArray(43) = gObjectTable.ObjectCount1
                        fxgEXEInfo.AddItem "CompiledObjects"
                        fxgEXEInfo.TextArray(45) = gObjectTable.iCompiledObjects
                        fxgEXEInfo.AddItem "ObjectsInUse"
                        fxgEXEInfo.TextArray(47) = gObjectTable.iObjectsInUse
                Case "VBOBJECTS"
                        If tblPath(3) <> "" And UBound(tblPath) = 4 Then

                            Dim objSel As Long
                            objSel = Val(tblPath(3))
                        
                            fxgEXEInfo.ColWidth(0) = 2500
                            fxgEXEInfo.TextArray(2) = "Address of ModulePublic"
                            fxgEXEInfo.TextArray(3) = gObject(objSel).aModulePublic
                            fxgEXEInfo.AddItem "Address of ModuleStatic"
                            fxgEXEInfo.TextArray(5) = gObject(objSel).aModuleStatic
                            fxgEXEInfo.AddItem "Address of ObjectInfo"
                            fxgEXEInfo.TextArray(7) = gObject(objSel).aObjectInfo
                            fxgEXEInfo.AddItem "Address of ObjectName"
                            fxgEXEInfo.TextArray(9) = gObject(objSel).aObjectName
                            fxgEXEInfo.AddItem "Address Proc Name Array"
                            fxgEXEInfo.TextArray(11) = gObject(objSel).aProcNamesArray
                            fxgEXEInfo.AddItem "Const1"
                            fxgEXEInfo.TextArray(13) = gObject(objSel).Const1
                            fxgEXEInfo.AddItem "Address of PublicBytes"
                            fxgEXEInfo.TextArray(15) = gObject(objSel).aPublicBytes
                            fxgEXEInfo.AddItem "Address of StaticBytes"
                            fxgEXEInfo.TextArray(17) = gObject(objSel).aStaticBytes
                            fxgEXEInfo.AddItem "Offset of StaticVars"
                            fxgEXEInfo.TextArray(19) = gObject(objSel).oStaticVars
                            fxgEXEInfo.AddItem "Null3"
                            fxgEXEInfo.TextArray(21) = gObject(objSel).Null3
                            fxgEXEInfo.AddItem "ObjectType"
                            fxgEXEInfo.TextArray(23) = gObject(objSel).ObjectType
                            fxgEXEInfo.AddItem "ProcCount"
                            fxgEXEInfo.TextArray(25) = gObject(objSel).ProcCount


                        End If
                        If UBound(tblPath) = 5 Then
                            Dim objInfosel As Long
                            objInfosel = Val(tblPath(4))
                            fxgEXEInfo.ColWidth(0) = 2500
                            fxgEXEInfo.TextArray(2) = "Address of ConstantPool"
                            fxgEXEInfo.TextArray(3) = gObjectInfoHolder(objInfosel).aConstantPool
                            fxgEXEInfo.AddItem "Address of Object"
                            fxgEXEInfo.TextArray(5) = gObjectInfoHolder(objInfosel).aObject
                            fxgEXEInfo.AddItem "Address of ObjectTable"
                            fxgEXEInfo.TextArray(7) = gObjectInfoHolder(objInfosel).aObjectTable
                            fxgEXEInfo.AddItem "Address of ProcTable"
                            fxgEXEInfo.TextArray(9) = gObjectInfoHolder(objInfosel).aProcTable
                            fxgEXEInfo.AddItem "Address of SmallRecord"
                            fxgEXEInfo.TextArray(11) = gObjectInfoHolder(objInfosel).aSmallRecord
                            fxgEXEInfo.AddItem "Const1"
                            fxgEXEInfo.TextArray(13) = gObjectInfoHolder(objInfosel).Const1
                            fxgEXEInfo.AddItem "Flag1"
                            fxgEXEInfo.TextArray(15) = gObjectInfoHolder(objInfosel).Flag1
                            fxgEXEInfo.AddItem "iConstantsCount"
                            fxgEXEInfo.TextArray(17) = gObjectInfoHolder(objInfosel).iConstantsCount
                            fxgEXEInfo.AddItem "iMaxConstants"
                            fxgEXEInfo.TextArray(19) = gObjectInfoHolder(objInfosel).iMaxConstants
                            fxgEXEInfo.AddItem "Flag5"
                            fxgEXEInfo.TextArray(21) = gObjectInfoHolder(objInfosel).Flag5
                            fxgEXEInfo.AddItem "Flag6"
                            fxgEXEInfo.TextArray(23) = gObjectInfoHolder(objInfosel).Flag6
                            fxgEXEInfo.AddItem "Flag7"
                            fxgEXEInfo.TextArray(25) = gObjectInfoHolder(objInfosel).Flag7
                            fxgEXEInfo.AddItem "Null1"
                            fxgEXEInfo.TextArray(27) = gObjectInfoHolder(objInfosel).Null1
                            fxgEXEInfo.AddItem "Null2"
                            fxgEXEInfo.TextArray(29) = gObjectInfoHolder(objInfosel).Null2
                            fxgEXEInfo.AddItem "NumberOfProcs"
                            fxgEXEInfo.TextArray(31) = gObjectInfoHolder(objInfosel).NumberOfProcs
                            fxgEXEInfo.AddItem "ObjectIndex"
                            fxgEXEInfo.TextArray(33) = gObjectInfoHolder(objInfosel).ObjectIndex
                            fxgEXEInfo.AddItem "RunTimeLoaded"
                            fxgEXEInfo.TextArray(35) = gObjectInfoHolder(objInfosel).RunTimeLoaded
                        End If
                End Select
            Case "EXEDATA"  '#####################################################'
                sstViewFile.TabVisible(1) = True
                sstViewFile.TabVisible(0) = False
                sstViewFile.TabVisible(2) = False
                fxgEXEInfo.Visible = True
                
                fxgEXEInfo.ColAlignment(1) = 0
                fxgEXEInfo.Clear
                fxgEXEInfo.Rows = 2
                fxgEXEInfo.ColWidth(0) = 2000
                fxgEXEInfo.TextArray(0) = "Name"
                fxgEXEInfo.TextArray(1) = "Value"
                
                Select Case tblPath(2)
                    Case "", "EXEHEADER"
                        fxgEXEInfo.ColWidth(0) = 1500
                        fxgEXEInfo.TextArray(2) = "Signature"
                        fxgEXEInfo.TextArray(3) = DosHeader.Magic
                        fxgEXEInfo.AddItem "Extra Bytes"
                        fxgEXEInfo.TextArray(5) = DosHeader.NumBytesLastPage
                        fxgEXEInfo.AddItem "Pages"
                        fxgEXEInfo.TextArray(7) = DosHeader.NumPages
                        fxgEXEInfo.AddItem "Reloc Items"
                        fxgEXEInfo.TextArray(9) = DosHeader.NumRelocates
                        fxgEXEInfo.AddItem "Header Size"
                        fxgEXEInfo.TextArray(11) = DosHeader.NumHeaderBlks
                        fxgEXEInfo.AddItem "Min Alloc"
                        fxgEXEInfo.TextArray(13) = DosHeader.ReservedW8
                        fxgEXEInfo.AddItem "Max Alloc"
                        fxgEXEInfo.TextArray(15) = DosHeader.ReservedW9
                        fxgEXEInfo.AddItem "Initial SS"
                        fxgEXEInfo.TextArray(17) = DosHeader.SSPointer
                        fxgEXEInfo.AddItem "Initial SP"
                        fxgEXEInfo.TextArray(19) = DosHeader.SPPointer
                        fxgEXEInfo.AddItem "Check Sum"
                        fxgEXEInfo.TextArray(21) = DosHeader.Checksum
                        fxgEXEInfo.AddItem "Initial IP"
                        fxgEXEInfo.TextArray(23) = DosHeader.IPPointer
                        fxgEXEInfo.AddItem "Initial CS"
                        fxgEXEInfo.TextArray(25) = DosHeader.CurrentSeg
                        fxgEXEInfo.AddItem "Reloc Table"
                        fxgEXEInfo.TextArray(27) = DosHeader.RelocTablePointer
                        fxgEXEInfo.AddItem "Overlay"
                        fxgEXEInfo.TextArray(29) = DosHeader.Overlay
                    Case "COFFHEADER"
                        fxgEXEInfo.ColWidth(0) = 2000
                        fxgEXEInfo.TextArray(2) = "Signature"
                        fxgEXEInfo.TextArray(3) = PEHeader.Magic
                        fxgEXEInfo.AddItem "Machine"
                        fxgEXEInfo.TextArray(5) = PEHeader.Machine
                        fxgEXEInfo.AddItem "Number Of Sections"
                        fxgEXEInfo.TextArray(7) = PEHeader.NumSections
                        fxgEXEInfo.AddItem "Time Date Stamp"
                        fxgEXEInfo.TextArray(9) = PEHeader.TimeDate
                        fxgEXEInfo.AddItem "Pointer To Symbol Table"
                        fxgEXEInfo.TextArray(11) = PEHeader.SymbolTablePointer
                        fxgEXEInfo.AddItem "Number Of Symbols"
                        fxgEXEInfo.TextArray(13) = PEHeader.NumSymbols
                        fxgEXEInfo.AddItem "Optional Header Size"
                        fxgEXEInfo.TextArray(15) = PEHeader.OptionalHdrSize
                        fxgEXEInfo.AddItem "Characteristics"
                        fxgEXEInfo.TextArray(17) = PEHeader.Properties
                    Case "OPTIONALHEADER"
                        
                        fxgEXEInfo.ColWidth(0) = 2500
                        fxgEXEInfo.TextArray(2) = "Magic"
                        fxgEXEInfo.TextArray(3) = modPeSkeleton.OptHeader.Magic
                        fxgEXEInfo.AddItem "Linker Major Version"
                        fxgEXEInfo.TextArray(5) = modPeSkeleton.OptHeader.MajLinkerVer
                        fxgEXEInfo.AddItem "Linker Minor Version"
                        fxgEXEInfo.TextArray(7) = modPeSkeleton.OptHeader.MinLinkerVer
                        fxgEXEInfo.AddItem "Size Of Code Section"
                        fxgEXEInfo.TextArray(9) = modPeSkeleton.OptHeader.CodeSize
                        fxgEXEInfo.AddItem "Initialized DataSize"
                        fxgEXEInfo.TextArray(11) = modPeSkeleton.OptHeader.InitDataSize
                        fxgEXEInfo.AddItem "Uninitialized DataSize"
                        fxgEXEInfo.TextArray(13) = modPeSkeleton.OptHeader.UninitDataSize
                        fxgEXEInfo.AddItem "Entry Point RVA"
                        fxgEXEInfo.TextArray(15) = modPeSkeleton.OptHeader.entrypoint
                        fxgEXEInfo.AddItem "Base Of Code"
                        fxgEXEInfo.TextArray(17) = modPeSkeleton.OptHeader.CodeBase
                        fxgEXEInfo.AddItem "Base Of Data"
                        fxgEXEInfo.TextArray(19) = modPeSkeleton.OptHeader.DataBase
                        fxgEXEInfo.AddItem "Image Base"
                        fxgEXEInfo.TextArray(21) = modPeSkeleton.OptHeader.ImageBase
                        fxgEXEInfo.AddItem "Section Alignement"
                        fxgEXEInfo.TextArray(23) = modPeSkeleton.OptHeader.SectionAlignment
                        fxgEXEInfo.AddItem "File Alignement"
                        fxgEXEInfo.TextArray(25) = modPeSkeleton.OptHeader.FileAlignment
                        fxgEXEInfo.AddItem "OS Major Version"
                        fxgEXEInfo.TextArray(27) = modPeSkeleton.OptHeader.MajOSVer
                        fxgEXEInfo.AddItem "OS Minor Version"
                        fxgEXEInfo.TextArray(29) = modPeSkeleton.OptHeader.MinOSVer
                        fxgEXEInfo.AddItem "User Major Version" 'bad
                        fxgEXEInfo.TextArray(31) = modPeSkeleton.OptHeader.MajImageVer
                        fxgEXEInfo.AddItem "User Minor Version" 'bad
                        fxgEXEInfo.TextArray(33) = modPeSkeleton.OptHeader.MinImageVer
                        fxgEXEInfo.AddItem "Sub Sys Major Version"
                        fxgEXEInfo.TextArray(35) = modPeSkeleton.OptHeader.MajSSysVer
                        fxgEXEInfo.AddItem "Sub Sys Minor Version"
                        fxgEXEInfo.TextArray(37) = modPeSkeleton.OptHeader.MinSSysVer
                        fxgEXEInfo.AddItem "Reserved" 'bad
                        fxgEXEInfo.TextArray(39) = modPeSkeleton.OptHeader.SSizeRes
                        fxgEXEInfo.AddItem "Image Size"
                        fxgEXEInfo.TextArray(41) = modPeSkeleton.OptHeader.SizeImage
                        fxgEXEInfo.AddItem "Header Size"
                        fxgEXEInfo.TextArray(43) = modPeSkeleton.OptHeader.SizeHeader
                        fxgEXEInfo.AddItem "File Checksum"
                        fxgEXEInfo.TextArray(45) = modPeSkeleton.OptHeader.Checksum
                        fxgEXEInfo.AddItem "Sub System"
                        fxgEXEInfo.TextArray(47) = modPeSkeleton.OptHeader.SSystem
                        fxgEXEInfo.AddItem "DLL Flags" 'bad
                        fxgEXEInfo.TextArray(49) = modPeSkeleton.OptHeader.LFlags
                        fxgEXEInfo.AddItem "Stack Reserved Size"
                        fxgEXEInfo.TextArray(51) = modPeSkeleton.OptHeader.SSizeRes
                        fxgEXEInfo.AddItem "Stack Commit Size"
                        fxgEXEInfo.TextArray(53) = modPeSkeleton.OptHeader.SSizeCom
                        fxgEXEInfo.AddItem "Heap Reserved Size"
                        fxgEXEInfo.TextArray(55) = modPeSkeleton.OptHeader.HSizeRes
                        fxgEXEInfo.AddItem "Heap Commit Size"
                        fxgEXEInfo.TextArray(57) = modPeSkeleton.OptHeader.HSizeCom
                        fxgEXEInfo.AddItem "Loader Flags"
                        fxgEXEInfo.TextArray(59) = modPeSkeleton.OptHeader.LFlags
                    Case "SECTIONHEADER"
                        If tblPath(3) <> "" Then
                            Dim SelSection As Long
                            SelSection = Val(tblPath(3))
                            
                            fxgEXEInfo.ColWidth(0) = 2000
                            fxgEXEInfo.TextArray(2) = "Section Name"
                            fxgEXEInfo.TextArray(3) = modPeSkeleton.SecHeader(SelSection).SecName
                           
                            fxgEXEInfo.AddItem "Virtual Size"
                            fxgEXEInfo.TextArray(5) = modPeSkeleton.SecHeader(SelSection).Properties
                            fxgEXEInfo.AddItem "RVA Offset"
                            fxgEXEInfo.TextArray(7) = modPeSkeleton.SecHeader(SelSection).Address
                            fxgEXEInfo.AddItem "Size Of Raw Data"
                            fxgEXEInfo.TextArray(9) = modPeSkeleton.SecHeader(SelSection).SizeRawData
                            fxgEXEInfo.AddItem "Pointer To Raw Data"
                            fxgEXEInfo.TextArray(11) = modPeSkeleton.SecHeader(SelSection).RawDataPointer
                            fxgEXEInfo.AddItem "Pointer To Relocs"
                            fxgEXEInfo.TextArray(13) = modPeSkeleton.SecHeader(SelSection).RelocationPointer
                            fxgEXEInfo.AddItem "Pointer To Line Numbers"
                            fxgEXEInfo.TextArray(15) = modPeSkeleton.SecHeader(SelSection).LineNumPointer
                            fxgEXEInfo.AddItem "Number Of Relocs"
                            fxgEXEInfo.TextArray(17) = modPeSkeleton.SecHeader(SelSection).NumRelocations
                            fxgEXEInfo.AddItem "Number Of Line Numbers"
                            fxgEXEInfo.TextArray(19) = modPeSkeleton.SecHeader(SelSection).NumLineNumbers
                            fxgEXEInfo.AddItem "Section Flags"
                            fxgEXEInfo.TextArray(21) = modPeSkeleton.SecHeader(SelSection).Misc
                        Else
                        'SecHeader(0).Address
                            'fxgEXEInfo.TextArray(2) = " " & ExtString(SecHeader(1).SecName)
                           ' fxgEXEInfo.TextArray(3) = AddChar(Hex(SecHeader(1).Address), 8)
                           For i = 1 To PEHeader.NumSections
                               fxgEXEInfo.AddItem " " & ExtString(SecHeader(i).SecName)
                               fxgEXEInfo.TextArray(3 + i * 2) = AddChar(Hex(SecHeader(i).Address), 8)
                            Next i
                        End If
                End Select
            Case "PROJECT"  '#####################################################'
                sstViewFile.TabVisible(0) = True
                sstViewFile.TabVisible(1) = False
                sstViewFile.TabVisible(2) = False
                sstViewFile.TabVisible(3) = False
                fxgEXEInfo.Visible = False
                Call modOutput.ShowVBPFile
                    gUpdateText = True
                    txtCode_Change
                    gUpdateText = False

            Case "CODE"     '#####################################################'
                sstViewFile.TabVisible(0) = True
                sstViewFile.TabVisible(1) = False
                sstViewFile.TabVisible(2) = False
                sstViewFile.TabVisible(3) = False
                fxgEXEInfo.Visible = False
                Select Case tblPath(2)
                    Case "", "API"
                    sstViewFile.TabVisible(0) = True
                    sstViewFile.TabVisible(1) = False
                    sstViewFile.TabVisible(2) = False
                    sstViewFile.TabVisible(3) = False
                    fxgEXEInfo.Visible = False
                    Call modGlobals.WriteApiList
                    gUpdateText = True
                    txtCode_Change
                    gUpdateText = False
                    Case "", "PCODE"
                        sstViewFile.TabVisible(0) = True
                        sstViewFile.TabVisible(1) = False
                        sstViewFile.TabVisible(2) = False
                        sstViewFile.TabVisible(3) = False
                        fxgEXEInfo.Visible = False
                        txtCode.LoadFile (App.path & "\dump\" & SFile & "\PcodeOut.txt")
                    
                    Case "", "ASM"
                        If gVBHeader.aSubMain <> 0 Then
                            txtCode.Text = ""
                            fp = FreeFile
                            Open SFilePath For Binary Access Read As #fp
                                modAsm.FileDeAsm gVBHeader.aSubMain + 1 - OptHeader.ImageBase, fp, LOF(fp), gVBHeader.aSubMain, True
                            Close #fp
                             txtCode.Text = txtCode.Text & "Disassembly of Sub Main()" & vbCrLf
    
                            For i = 1 To UBound(modAsm.StrDEASM())
                                txtCode.Text = txtCode.Text & modAsm.StrDEASM(i) & vbCrLf
                            Next i
                        Else
                            MsgBox "There is no Sub Main in this program"
                        End If
                    
                End Select
                
            Case "FORMS"
                If tblPath(2) <> "" Then
                    sstViewFile.TabVisible(0) = True
                    sstViewFile.TabVisible(3) = True
                    sstViewFile.TabVisible(1) = False
                    sstViewFile.TabVisible(2) = False
                    fxgEXEInfo.Visible = False
                    
                    For i = 0 To txtFinal.UBound
                        If UCase(txtFinal(i).Tag) = tblPath(2) Then
                            txtCode.Text = txtFinal(i).Text
                            txtCode.Text = txtCode.Text & "'BEGIN Code List" & vbCrLf
                            For nApi = 0 To UBound(gProcedureList)
                                If UCase(tblPath(2)) = UCase(gProcedureList(nApi).strParent) Then
                                    txtCode.Text = txtCode.Text & "Sub " & gProcedureList(nApi).strProcedureName & "()" & vbCrLf
                                    txtCode.Text = txtCode.Text & "End Sub" & vbCrLf
                                End If
                            Next
                            Exit For
                            

                        End If
                         Next
                        
                            For g = 0 To UBound(gObjectOffsetArray)
                                If UCase(gObjectOffsetArray(g).ObjectName) = UCase(tblPath(2)) Then
                                    sstViewFile.Tab = 3
                                    'MsgBox gObjectOffsetArray(g).Address
                                    Call modControls.GetControlProperties(gObjectOffsetArray(g).Address)
                                    Exit For
                                End If
                                
                            Next
                   
                    gUpdateText = True
                    txtCode_Change
                    gUpdateText = False
                End If
            Case "USERCONTROL"
                If tblPath(2) <> "" Then
                    sstViewFile.TabVisible(0) = True
                    sstViewFile.TabVisible(3) = True
                    sstViewFile.TabVisible(1) = False
                    sstViewFile.TabVisible(2) = False
                    fxgEXEInfo.Visible = False
                    
                    For i = 0 To txtFinal.UBound
                        If UCase(txtFinal(i).Tag) = tblPath(2) Then
                            txtCode.Text = txtFinal(i).Text
                            txtCode.Text = txtCode.Text & "'BEGIN Code List" & vbCrLf
                            For nApi = 0 To UBound(gProcedureList)
                                If UCase(tblPath(2)) = UCase(gProcedureList(nApi).strParent) Then
                                    txtCode.Text = txtCode.Text & "Sub " & gProcedureList(nApi).strProcedureName & "()" & vbCrLf
                                    txtCode.Text = txtCode.Text & "End Sub" & vbCrLf
                                End If
                            Next
                            Exit For
                            

                        End If
                         Next
                        
                   
                    gUpdateText = True
                    txtCode_Change
                    gUpdateText = False
                End If
                
                Case "MODS"
                If tblPath(2) <> "" Then
                    sstViewFile.TabVisible(0) = True
                    sstViewFile.TabVisible(1) = False
                    sstViewFile.TabVisible(2) = False
                    sstViewFile.TabVisible(3) = False
                    fxgEXEInfo.Visible = False
                    txtCode.Text = ""
                    txtCode.Text = txtCode.Text & "'BEGIN Code List" & vbCrLf
                    For nApi = 0 To UBound(gProcedureList)
                        If UCase(tblPath(2)) = UCase(gProcedureList(nApi).strParent) Then
                            txtCode.Text = txtCode.Text & "Sub " & gProcedureList(nApi).strProcedureName & "()" & vbCrLf
                            txtCode.Text = txtCode.Text & "End Sub" & vbCrLf
                        End If
                    Next
                    gUpdateText = True
                    txtCode_Change
                    gUpdateText = False
                  End If
                Case "CLASS"
                If tblPath(2) <> "" Then
                    sstViewFile.TabVisible(0) = True
                    sstViewFile.TabVisible(1) = False
                    sstViewFile.TabVisible(2) = False
                    sstViewFile.TabVisible(3) = False
                    fxgEXEInfo.Visible = False
                    txtCode.Text = ""
                    txtCode.Text = txtCode.Text & "'BEGIN Code List" & vbCrLf
                    For nApi = 0 To UBound(gProcedureList)
                        If UCase(tblPath(2)) = UCase(gProcedureList(nApi).strParent) Then
                            txtCode.Text = txtCode.Text & "Sub " & gProcedureList(nApi).strProcedureName & "()" & vbCrLf
                            txtCode.Text = txtCode.Text & "End Sub" & vbCrLf
                        End If
                    Next
                    gUpdateText = True
                    txtCode_Change
                    gUpdateText = False
                End If
                Case "IMAGES"
                'Image Preview
                If tblPath(2) <> "" Then
                sstViewFile.TabVisible(2) = True
                    sstViewFile.TabVisible(0) = False
                    sstViewFile.TabVisible(1) = False
                    sstViewFile.TabVisible(3) = False
                    
                    fxgEXEInfo.Visible = False
             
                    For i = 0 To UBound(FrxPreview) - 1
                        If UCase(tblPath(2)) = UCase(FrxPreview(i).strPath) Then
                            On Error Resume Next
                            picPreview.Picture = LoadPicture(App.path & "\dump\" & SFile & "\" & FrxPreview(i).strPath)
                        End If
                    Next i
                End If
                  
        End Select
    
     
    End If
End Sub
Private Sub SetupTreeView()
'*****************************
'Purpose: Sets up all the nodes in the Treeview control
'*****************************
    Dim Parent(0 To &HFF) As String, LenTab As Long, IsMenu As Boolean
    Dim i As Long, o As Long, e As Long
    Filename = SFile
    Call tvProject.Nodes.Add(, , "ROOT/PROJECT/" & Filename, Mid(Filename, InStrRev(Filename, "\") + 1), 34)

    tvProject.Nodes(1).Selected = True
    tvProject.Nodes(1).Expanded = True
    tvProject_NodeClick tvProject.Nodes(1)
    
    '####################   Information about the exe  ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/EXEDATA/", "PE Header", 1)
    Parent(0) = "ROOT/EXEDATA/"
    tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/EXEDATA/EXEHEADER/", "EXE Header", 2
    tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/EXEDATA/COFFHEADER/", "Coff Header", 2
    tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/EXEDATA/OPTIONALHEADER/", "Optional Header", 2
    tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/EXEDATA/SECTIONHEADER/", "Section Header", 3
    
   For i = 1 To PEHeader.NumSections
     tvProject.Nodes.Add "ROOT/EXEDATA/SECTIONHEADER/", tvwChild, "ROOT/EXEDATA/SECTIONHEADER/" & i & "/", SecHeader(i).SecName, 2
    Next i
    '####################   VB Strutures       ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/STRUCT/", "VB Structures", 1)
    Parent(0) = "ROOT/STRUCT/"
    tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/STRUCT/VBHEADER/", "VB Header", 2
    tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/STRUCT/VBPROJECTINFO/", "VB Project Information", 2
    tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/STRUCT/VBCOMREGDATA/", "VB Com Registration Data", 2
    tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/STRUCT/VBCOMREGINFO/", "VB COM Registration Info", 2
    tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/STRUCT/VBOBJECTABLE/", "VB Object Table", 2
    tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/STRUCT/VBOBJECTS/", "VB Object's", 2
    For i = 0 To UBound(gObject)
     tvProject.Nodes.Add "ROOT/STRUCT/VBOBJECTS/", tvwChild, "ROOT/STRUCT/VBOBJECTS/" & i & "/", "Object: " & gObjectNameArray(i), 2
     tvProject.Nodes.Add "ROOT/STRUCT/VBOBJECTS/" & i & "/", tvwChild, "ROOT/STRUCT/VBOBJECTS/OBJINFO/" & i & "/", "ObjectInfo", 2
    Next i
    
    '####################   VB Forms       ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/FORMS/", "Forms", 1)
    Parent(0) = "ROOT/FORMS/"
    For i = 0 To UBound(gObject)
        If gObject(i).ObjectType = 98435 Then
            tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/FORMS/" & UCase(gObjectNameArray(i)) & "/", gObjectNameArray(i), 10
            tvProject.Nodes.Add Parent(0) & UCase(gObjectNameArray(i)) & "/", 4, "ROOT/FORMS/" & UCase(gObjectNameArray(i)) & "/SUBCOUNT/", "Methods Count: " & gObjectProcCountArray(i), 45
        End If
    Next
    For i = 1 To UBound(gControlNameArray)
        If gControlNameArray(i).strControlName <> "" And gControlNameArray(i).strControlName <> "Form" Then
            On Error Resume Next
            'tvProject.Nodes.Add "ROOT/FORMS/" & UCase(gControlNameArray(i).strParentForm) & "/", tvwChild, "ROOT/FORMS/" & UCase(gControlNameArray(i).strParentForm) & "/" & i & "/", gControlNameArray(i).strControlName, 2
            tvProject.Nodes.Add "ROOT/FORMS/" & UCase(gControlNameArray(i).strParentForm) & "/", tvwChild, "ROOT/FORMS/" & UCase(gControlNameArray(i).strParentForm) & "/" & i & "/", gControlNameArray(i).strControlName, 2
            
        End If
    Next
    '####################   VB Modules       ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/MODS/", "Modules", 1)
    Parent(0) = "ROOT/MODS/"
    AppData.AppModuleCount = 0
    For i = 0 To UBound(gObject)
        If gObject(i).ObjectType = 98305 Then
            tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/MODS/" & UCase(gObjectNameArray(i)) & "/", gObjectNameArray(i), 40
            tvProject.Nodes.Add Parent(0) & UCase(gObjectNameArray(i)) & "/", 4, "ROOT/MODS/" & UCase(gObjectNameArray(i)) & "/SUBCOUNT/", "Methods Count: " & gObjectProcCountArray(i), 45
            AppData.AppModuleCount = AppData.AppModuleCount + 1
        End If
    Next
    '####################   VB Classes       ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/CLASS/", "Classes", 1)
    Parent(0) = "ROOT/CLASS/"
    For i = 0 To UBound(gObject)
        If gObject(i).ObjectType = 1146883 Then
            tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/CLASS/" & UCase(gObjectNameArray(i)) & "/", gObjectNameArray(i), 41
             tvProject.Nodes.Add Parent(0) & UCase(gObjectNameArray(i)) & "/", 4, "ROOT/CLASSS/" & UCase(gObjectNameArray(i)) & "/SUBCOUNT/", "Methods Count: " & gObjectProcCountArray(i), 45
        End If
    Next
    '####################   User Controls       ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/USERCONTROL/", "User Controls", 1)
    Parent(0) = "ROOT/USERCONTROL/"

    For i = 0 To UBound(gObject)
        If gObject(i).ObjectType = 1941507 Or gObject(i).ObjectType = 1943555 Then
            tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/USERCONTROL/" & UCase(gObjectNameArray(i)) & "/", gObjectNameArray(i), 43
             tvProject.Nodes.Add Parent(0) & UCase(gObjectNameArray(i)) & "/", 4, "ROOT/USERCONTROL/" & UCase(gObjectNameArray(i)) & "/SUBCOUNT/", "Methods Count: " & gObjectProcCountArray(i), 45
        End If
    Next
    '####################   Property Pages       ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/PROPERTYPAGE/", "Property Pages", 1)
    Parent(0) = "ROOT/PROPERTYPAGE/"
    For i = 0 To UBound(gObject)
        If gObject(i).ObjectType = 1409027 Then
            tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/CLASS/" & UCase(gObjectNameArray(i)) & "/", gObjectNameArray(i), 42
        End If
    Next
    '####################   Procedures - Code       ####################'

    Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/CODE/", "Procedures - Code", 1)
    Parent(0) = "ROOT/CODE/"
    'Add P-Code View
    If gProjectInfo.aNativeCode = 0 Then
        tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/CODE/" & "PCODE", "View P-Code", 4
    End If
    
    tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/CODE/" & "API", "API List", 4
    tvProject.Nodes.Add(Parent(0), tvwChild, "ROOT/CODE/" & "ASM", "Code assembly for Sub Main", 4).Tag = -2
    
    '####################   Images     ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/IMAGES/", "Images", 1)
    Parent(0) = "ROOT/IMAGES/"
    For i = 0 To UBound(FrxPreview) - 1
            tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/IMAGES/" & UCase(FrxPreview(i).strPath) & "/", FrxPreview(i).strPath, 6
    Next
    '####################   File Version Information    ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/VERSIONINFO/", "File Version Information", 1)

    '####################   Other Information    ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/OTHER/", "Import Information", 1)
    Parent(0) = "ROOT/OTHER/"
    
        Dim TDs As String, ouR As String
        j = 1
          
            tvProject.Nodes.Add Parent(0), 4, LCase$(ImportList(0).strName), ImportList(0).strName, 44
                k = UBound(exeIMPORT_APINAME())
                Do While j <= k
                
                    tvProject.Nodes.Add LCase$(ImportList(0).strName), 4, exeIMPORT_APINAME(j).ApiName, exeIMPORT_APINAME(j).ApiName, 44
                    tvProject.Nodes.Add exeIMPORT_APINAME(j).ApiName, 4, , "Offset " & Hex(exeIMPORT_APINAME(j).Address) & "h", 44
                    
                    If Left$(LCase$(ImportList(0).strName), 8) = "msvbvm60" Then
                    If Left$(exeIMPORT_APINAME(j).ApiName, 8) = "!ordinal" Then
                        'via ordinal
                        TDs = VBFunction_Description(Val(Mid$(exeIMPORT_APINAME(j).ApiName, 12)), "", ouR)
                        If TDs = "undef" Then
                            tvProject.Nodes.Add exeIMPORT_APINAME(j).ApiName, 4, , "Name : " & ouR, 18
                        Else
                            tvProject.Nodes.Add exeIMPORT_APINAME(j).ApiName, 4, , "Name: " & ouR, 18
                          
                            tvProject.Nodes.Add exeIMPORT_APINAME(j).ApiName, 4, , TDs, 19
                        End If
                    Else
                        'via directname
                        TDs = VBFunction_Description(0, exeIMPORT_APINAME(j).ApiName, ouR)
                        If TDs = "undef" Then
                        Else
                            tvProject.Nodes.Add exeIMPORT_APINAME(j).ApiName, 4, , TDs, 19
                        End If
                    End If
                    End If
                    j = j + 1
                Loop
            


    CurrentNode = 0

    tvProject_NodeClick tvProject.Nodes(1)
End Sub
Public Function GetOpcode(FileNum As Variant) As Byte
'*****************************
'Purpose: Function just retrieves a byte used to get gui opcode
'*****************************
    Dim Opcode As Byte
    Get FileNum, , Opcode
    GetOpcode = Opcode
End Function
Private Function DumpObject(FileNum As Variant, ObjectName As String, length As Long, FileStart As Long, HeaderEnd As Long) As Long
'*****************************
'Purpose: Dumps a Gui Object
'*****************************
  On Error GoTo bad
    MakeDir (App.path & "\dump")
    MakeDir (App.path & "\dump\" & SFile)
    Dim bArray() As Byte
    ReDim bArray(length)
    'Get the ojbect information
    Seek FileNum, FileStart + 1
    
    Get FileNum, , bArray
    Dim fFileEnd As Long
    fFileEnd = Loc(FileNum)
    Seek FileNum, HeaderEnd
    'Save the information
    Open App.path & "\dump\" & SFile & "\" & ObjectName & ".txt" For Binary Access Write Lock Write As #12
        Put #12, , bArray
    Close #12
    
    DumpObject = (fFileEnd + 1)
    Exit Function
bad:
    DumpObject = -1
Exit Function
End Function

Sub GetStdPicture(FileNum As Variant, length As Variant, strName As String, ParentForm As String, fAddress As Long)
'*****************************
'Purpose: To save an STD Picture and detect what kind of picture file it is.
'*****************************
   On Error Resume Next
    Dim picHeader As typePictureHeader
    Dim bPicArray() As Byte
   
    'Get Picture Header
    Get FileNum, , picHeader

    
    Dim strExt As String
    strExt = ".ico"
    length = length - 8
    
    If length > 500000 Then Exit Sub
    If length < 0 Then Exit Sub
    
     ReDim bPicArray(length)
     Get FileNum, , bPicArray

    If bPicArray(0) = 66 And bPicArray(1) = 77 Then
        strExt = ".bmp"
    End If
    If bPicArray(0) = 71 And bPicArray(1) = 73 And bPicArray(2) = 70 Then
        strExt = ".gif"
    End If
    If bPicArray(0) = 0 And bPicArray(2) = 1 Then
        strExt = ".ico"
    End If
    If bPicArray(0) = 0 And bPicArray(2) = 2 Then
        strExt = ".cur"
    End If
    If bPicArray(0) = 255 And bPicArray(1) = 216 Then
        strExt = ".jpg"
    End If
    If bPicArray(0) = 215 And bPicArray(1) = 205 Then
        strExt = ".wmf"
    End If
    
    FrxPreview(UBound(FrxPreview)).strPath = strName & strExt
    FrxPreview(UBound(FrxPreview)).FRXAddress = fAddress
    FrxPreview(UBound(FrxPreview)).length = length
    FrxPreview(UBound(FrxPreview)).ParentForm = ParentForm
    Open App.path & "\dump\" & SFile & "\" & strName & strExt For Binary Access Write Lock Write As #23
        
        Put #23, , bPicArray

    Close #23
    ReDim Preserve FrxPreview(UBound(FrxPreview) + 1)
End Sub


Sub ProccessControls(f As Variant)
'*****************************
'Purpose: Process Forms And Control Properties
'*****************************
Dim bFormEndUsed As Boolean
Dim strCurrentForm As String
    'Erase existing data
    bFormEndUsed = False

    If gVBHeader.FormCount = 0 Then Exit Sub
    ReDim gObjectOffsetArray(0)

        Seek f, gVBHeader.aGUITable + 1 - OptHeader.ImageBase
        
        'Get Form table
        If gVBHeader.FormCount > 0 Then
            ReDim gGuiTable(gVBHeader.FormCount - 1)
            Get #f, , gGuiTable
          
        End If
        Dim fPos As Long 'Holds current location in the file used for controlheader
        Dim cListIndex As Integer ' Used for COM
        Dim cControlHeader As ControlHeader
        
        Dim lForm As Integer
        Dim FRXAddress As Long
        'Loop though each form...
        For lForm = 0 To UBound(gGuiTable)

        Seek f, gGuiTable(lForm).aFormPointer + 94 - OptHeader.ImageBase
        FRXAddress = 0
        Dim posNextControl As Long
       
'Loop from new child control
NewControl:
        fPos = Loc(f)
        
        Seek f, Loc(f) + 4
      
        Dim intArray As Integer
        intArray = GetInteger(f)
        Seek f, fPos + 1
   
        If intArray <> 384 Then
            Get #f, , cControlHeader
        Else
            Dim cArrayHeader As ControlArrayHeader
            Get #f, , cArrayHeader '
            'MsgBox cArrayHeader.cType
            cControlHeader.length = cArrayHeader.length
            cControlHeader.cName = cArrayHeader.cName
            cControlHeader.cType = cArrayHeader.cType
            cControlHeader.cId = cArrayHeader.cId
        End If
        posNextControl = fPos + cControlHeader.length + 2
        'MsgBox posNextControl
       If gDumpData = True Then
       
        Dim fHeaderEnd As Long
        
        fHeaderEnd = Loc(f)

        Dim fControlEnd As Long
        'Store each object's information in a file
        fControlEnd = DumpObject(f, cControlHeader.cName, CLng(cControlHeader.length), fPos, fHeaderEnd)
        If gSkipCom = True Then
        'Get all dumps of the controls even though COM is off
            If fControlEnd <> -1 Then
                Seek f, fControlEnd
                GoTo NewControl
            End If
        End If
       End If
       If gSkipCom = False Then
        Dim tliTypeInfo As TypeInfo 'Used for COM to find information about the properties of the control
        Dim FileLen As Long 'Used to caculate how much father to go in the control
        'Select what type of control it is
        If gShowOffsets = True Then AddText "'Object Offset: " & fPos
       ' MsgBox cControlHeader.cName
        Select Case cControlHeader.cType
            
            Case vbPictureBox '= 0
                cListIndex = 22
               
                Call AddText("Begin VB.PictureBox " & cControlHeader.cName)
                gIdentSpaces = gIdentSpaces + 1
            Case vbLabel '= 1
                cListIndex = 14
                Call AddText("Begin VB.Label " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
            Case vbTextBox ' = 2
                cListIndex = 27
                Call AddText("Begin VB.TextBox " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
            Case vbFrame '= 3
                cListIndex = 10
                Call AddText("Begin VB.Frame " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
            Case vbCommandbutton '= 4
                cListIndex = 4
                 
                Call AddText("Begin VB.CommandButton " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
            Case vbCheckbox '= 5
                cListIndex = 1
                Call AddText("Begin VB.Checkbox " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
            Case vbOptionbutton     ' = 6
                cListIndex = 21
                Call AddText("Begin VB.Optionbutton " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
            Case vbComboBox     ' = 7
                cListIndex = 3
                Call AddText("Begin VB.Combobox " & cControlHeader.cName)
                gIdentSpaces = gIdentSpaces + 1
            Case vbListbox     '= 8
                cListIndex = 17
                Call AddText("Begin VB.ListBox " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
            Case vbHscroll     '= 9
                cListIndex = 12
                Call AddText("Begin VB.HScrollBar " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
            Case vbVscroll     '= 10
                cListIndex = 32
                Call AddText("Begin VB.VScrollBar " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
            Case vbTimer     '= 11
                cListIndex = 28
                Call AddText("Begin VB.Timer " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
            Case vbForm     '= 13
                cListIndex = 9
                strCurrentForm = cControlHeader.cName
                
                Call modGlobals.LoadNewFormHolder(cControlHeader.cName)
                gIdentSpaces = 0
                Call AddText("Begin VB.Form " & cControlHeader.cName)
                 gObjectOffsetArray(UBound(gObjectOffsetArray)).ObjectName = cControlHeader.cName
                 gObjectOffsetArray(UBound(gObjectOffsetArray)).Address = fPos
                ReDim Preserve gObjectOffsetArray(UBound(gObjectOffsetArray) + 1)
                txtStatus.Text = txtStatus.Text & "Processing Form:" & strCurrentForm & vbCrLf
                FrameStatus.Refresh
                txtStatus.SelStart = Len(txtStatus)
                txtStatus.Refresh
                gIdentSpaces = 1
            Case vbDriveListbox     '= 16
                cListIndex = 7
                Call AddText("Begin VB.DriveListbox " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
            Case vbDirectoryListbox     '= 17
                cListIndex = 6
                Call AddText("Begin VB.DirectoryListbox " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
            Case vbFileListbox     '= 18
                cListIndex = 8
                Call AddText("Begin VB.FileListBox " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
            Case vbMenu     '= 19
                cListIndex = 19
                Call AddText("Begin VB.Menu " & cControlHeader.cName)
                gIdentSpaces = gIdentSpaces + 1
            Case vbMDIForm     '= 20
                cListIndex = 18
                gIdentSpaces = 0
                Call AddText("Begin VB.MDIForm " & cControlHeader.cName)
                gIdentSpaces = gIdentSpaces + 1
                strCurrentForm = cControlHeader.cName
                
                Call modGlobals.LoadNewFormHolder(cControlHeader.cName)
                 gObjectOffsetArray(UBound(gObjectOffsetArray)).ObjectName = cControlHeader.cName
                 gObjectOffsetArray(UBound(gObjectOffsetArray)).Address = fPos
                ReDim Preserve gObjectOffsetArray(UBound(gObjectOffsetArray) + 1)
                
                 
            Case vbShape     '= 22
                cListIndex = 26
                Call AddText("Begin VB.Shape " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
            Case vbLine     '= 23
                cListIndex = 16
                Call AddText("Begin VB.Line " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
            Case vbImage     '= 24
                cListIndex = 13
                Call AddText("Begin VB.Image " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
            Case vbData     '= 37
                cListIndex = 5
                Call AddText("Begin VB.Data " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
            Case vbOLE     '= 38
                cListIndex = 20
                Call AddText("Begin VB.OLE " & cControlHeader.cName)
            Case vbUserControl     '= 40
                cListIndex = 29
                gIdentSpaces = 0
                Call AddText("Begin VB.UserControl " & cControlHeader.cName)
                strCurrentForm = cControlHeader.cName
                Call modGlobals.LoadNewFormHolder(cControlHeader.cName)
                 gObjectOffsetArray(UBound(gObjectOffsetArray)).ObjectName = cControlHeader.cName
                 gObjectOffsetArray(UBound(gObjectOffsetArray)).Address = fPos
                ReDim Preserve gObjectOffsetArray(UBound(gObjectOffsetArray) + 1)
                
            
            Case vbPropertyPage     '= 41
                cListIndex = 24
                Call AddText("Begin VB.PropertyPage " & cControlHeader.cName)
            Case vbUserDocument     '= 42
                cListIndex = 30
                Call AddText("Begin VB.UserDocument " & cControlHeader.cName)
            Case 255 'external control
                Call AddText("Begin " & GetAllString(f) & " " & cControlHeader.cName & " 'Length:" & cControlHeader.length)
                'Load the control view COM if its on the computer
                  gIdentSpaces = gIdentSpaces + 1
                Seek f, fPos + cControlHeader.length ' - 2
                
        End Select
        Set tliTypeInfo = tliTypeLibInfo.GetTypeInfo(Replace(Replace(lstTypeInfos.List(cListIndex), "<", ""), ">", ""))
        'Use the ItemData in lstTypeInfos to set the SearchData for lstMembers
        tliTypeLibInfo.GetMembersDirect lstTypeInfos.ItemData(cListIndex), lstMembers.hwnd, , , True
        
        FileLen = Loc(f) - fPos
        FileLen = cControlHeader.length - FileLen
        
        Dim bCode As Byte 'holds gui opcode
        Dim varHold As Variant 'Holds the different data types
        Dim strHold As String 'holds the string
        Dim strReturnType As String 'holds the return type

        Do While Loc(f) < (fPos + cControlHeader.length - 2)
       
        'Do Until Loc(f) >= (fPos + cControlHeader.Length - 1)
         bCode = GetOpcode(f) 'Get the guiopcode
         
         FileLen = FileLen - 1
        
         Dim g As Integer
         For g = 0 To lstMembers.ListCount - 1
      
         
            'Control Postion opcode
            If bCode = 4 And cControlHeader.cType = vbDirectoryListbox Then
                Call GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 3 And cControlHeader.cType = vbImage Then
                Call GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 4 And cControlHeader.cType = vbListbox Then
                Call GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 4 And cControlHeader.cType = vbDriveListbox Then
                Call GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 4 And cControlHeader.cType = vbFileListbox Then
                Call GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 4 And cControlHeader.cType = vbTextBox Then
                Call GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 4 And cControlHeader.cType = vbCommandbutton Then
                Call GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 5 And cControlHeader.cType = vbPictureBox Then
                Call GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 5 And cControlHeader.cType = vbComboBox Then

                Call GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
   
            End If
            If cControlHeader.cType = vbShape And bCode = 4 Then

                Call GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
   
            End If
            If bCode = 5 And cControlHeader.cType = vbOptionbutton Then
                Call GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 5 And cControlHeader.cType = vbFrame Then
                Call GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 5 And cControlHeader.cType = vbCheckbox Then
                Call GetControlSize(f)
                FileLen = FileLen - 8
                Exit For
            End If
            If bCode = 5 And cControlHeader.cType = vbLabel Then
                Call GetControlSize(f)
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
                Call GetControlSize(f)
                FileLen = FileLen - 8
                Exit For

            End If
            If bCode = 2 And cControlHeader.cType = vbVscroll Then
                Call GetControlSize(f)
                FileLen = FileLen - 8
                Exit For

            End If
            If bCode = 37 And cControlHeader.cType = vbLabel Then
                'Font
                Call GetFontProperty(f)
                'MsgBox "Loc: " & Loc(f)
                'MsgBox fPos + cControlHeader.Length - 2
               ' MsgBox "Endtest " & (fPos + cControlHeader.Length - 2)
               ' MsgBox "Loc!:" & Loc(f)
                Exit For
            End If
            If bCode = 64 And cControlHeader.cType = vbForm Then
                'Font
                Call GetFontProperty(f)
                Exit For
            End If
            If cControlHeader.cType = vbCommandbutton And bCode = 29 Then
                Call GetFontProperty(f)
                Exit For
            End If
            If cControlHeader.cType = vbPictureBox And bCode = 57 Then
                Call GetFontProperty(f)
                Exit For
            End If
            If cControlHeader.cType = vbTextBox And bCode = 46 Then
                Call GetFontProperty(f)
                Exit For
            End If
            If cControlHeader.cType = vbCheckbox And bCode = 32 Then
                Call GetFontProperty(f)
                Exit For
            End If

            If cControlHeader.cType = vbFrame And bCode = 4 Then
                
                Call AddText("ForeColor=" & GetLong(f))
                Exit For
            End If
            If cControlHeader.cType = vbPictureBox And bCode = 66 Then
                GetByte2 (f)
                Exit For
            End If
            
            If cControlHeader.cType = vbLine And bCode = 3 Then
                Dim LineSize As LineSizeType
                Get f, , LineSize
                Call AddText("X1 = " & LineSize.X1)
                Call AddText("Y1 = " & LineSize.Y1)
                Call AddText("X2 = " & LineSize.X2)
                Call AddText("Y2 = " & LineSize.Y2)
                Exit For
            End If
            If ReturnGuiOpcode(lstTypeInfos.ItemData(cListIndex), tliInvokeKinds, lstMembers.List(g)) = bCode Then
              Dim strExtraInfo As String
              'MsgBox "Prop: " & lstMembers.List(g)
                strReturnType = Trim(ReturnDataType(lstTypeInfos.ItemData(cListIndex), tliInvokeKinds, lstMembers.List(g)))
              'MsgBox "Prop: " & lstMembers.List(g) & " " & strReturnType & " Gui: " & bCode & " FileLen: " & FileLen & " Loc" & Loc(F)
                If gShowOffsets = True Then
                strExtraInfo = "  ' GuiOpcode: " & bCode & " Offset Dec: " & Loc(f)
                End If
                'Com Hack Check
                For k = 0 To UBound(gComFix)
                    If lstTypeInfos.List(cListIndex) = gComFix(k).ObjectName And lstMembers.List(g) = gComFix(k).PropertyName Then
                   'If lstMembers.List(g) = gComFix(k).PropertyName Then
                        strReturnType = gComFix(k).NewType
                        
                        Exit For
                    End If
                Next

                If InStr(1, strReturnType, "Byte") Then
                    varHold = GetByte2(f)
                    Call AddText(lstMembers.List(g) & " = " & varHold & strExtraInfo)
                    FileLen = FileLen - 1
                    Exit For
                End If
                If InStr(1, strReturnType, "Boolean") Then
                    varHold = GetBoolean(f)
                    If varHold = True Then
                        Call AddText(lstMembers.List(g) & " = " & -1 & strExtraInfo)
                    Else
                        Call AddText(lstMembers.List(g) & " = " & 0 & strExtraInfo)
                    End If
                    Seek f, Loc(f)
                    FileLen = FileLen - 2
                    Exit For
                End If
                If InStr(1, strReturnType, "Integer") Then
                    varHold = gVBFile.GetInteger(Loc(f))
                    'varHold = GetInteger(f)
                    Call AddText(lstMembers.List(g) & " = " & varHold & strExtraInfo)
                    FileLen = FileLen - 2
                    Exit For
                End If
                If InStr(1, strReturnType, "Long") Then
                    varHold = GetLong(f)
                    Call AddText(lstMembers.List(g) & " = " & varHold & strExtraInfo)
                    FileLen = FileLen - 4
                    Exit For
                End If
                
                If InStr(1, strReturnType, "Single") Then
                    varHold = GetSingle(f)
                    Call AddText(lstMembers.List(g) & " = " & varHold & strExtraInfo)
                    FileLen = FileLen - 4
                    Exit For
                End If

                If InStr(1, strReturnType, "String") Then
                    ''Seek f, Loc(f) + 3
                    ''strHold = GetUntilNull(f)
                    strHold = GetAllString(f)
                    Call AddText(lstMembers.List(g) & " = " & Chr(34) & strHold & Chr(34) & strExtraInfo)
                    FileLen = FileLen - Len(strHold) - 3
                    Exit For
                End If
                If InStr(1, strReturnType, "Unicode") Then
                    ''Seek f, Loc(f) + 3
                    ''strHold = GetUntilNull(f)
                    strHold = gVBFile.GetString(Loc(f) + 2, , True)
                    
                    Call AddText(lstMembers.List(g) & " = " & Chr(34) & strHold & Chr(34) & strExtraInfo)
                    Seek f, Loc(f) - 5
                    
                    FileLen = FileLen - Len(strHold) - 3
                    Exit For
                End If
              
                If InStr(1, strReturnType, "stdole.Picture") Then
                    
                    varHold = GetLong(f)
                   
                    If varHold <> -1 Then
                    'MsgBox "Loc:" & Loc(f) & " " & varHold
                        If cControlHeader.cName <> strCurrentForm Then
                            Call GetStdPicture(f, varHold, strCurrentForm & "." & cControlHeader.cName, strCurrentForm, FRXAddress)
                        Else
                            Call GetStdPicture(f, varHold, cControlHeader.cName, strCurrentForm, FRXAddress)
                        End If
                        
                        
                        Call AddText(lstMembers.List(g) & "=" & Chr(34) & strCurrentForm & ".frx" & Chr(34) & ":" & PadHex(Hex(FRXAddress), 4) & strExtraInfo)
                        Seek f, Loc(f)
                     
                       FRXAddress = FRXAddress + varHold + 12
                   'Exit Sub
                        'FileLen = FileLen - varHold + 1 ' - 18
                        FileLen = FileLen - 12
                        'MsgBox varHold
                        'MsgBox "FileLen:" & FileLen

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
                    Call AddText("ClientLeft = " & objectSize.clientLeft)
                    Call AddText("ClientTop = " & objectSize.clientTop)
                    Call AddText("ClientWidth = " & objectSize.clientWidth)
                    Call AddText("ClientHeight = " & objectSize.clientHeight)
                End If
                
                Exit For
            End If
         Next
         
         'Exit the Process controls in case it hangs on a property
         If CancelDecompile = True Then Exit Sub
         DoEvents
        Loop
        
        
        'Get the seperator type for the end of the control
        Dim cControlEnd As Integer
        'Seek f, Loc(f) '+1
        '''Seek f, posNextControl - 2


        
        'Exit Sub
        
        cControlEnd = GetInteger(f)
        'MsgBox cControlEnd & " Loc:" & Loc(F)
        
        If cControlEnd = vbFormEnd Then
            bFormEndUsed = True
            gIdentSpaces = gIdentSpaces - 1
            Call AddText("End")
         
        End If
        If cControlEnd = vbFormNewChildControl Then
            Seek f, posNextControl
            GoTo NewControl
            
        End If
        If cControlEnd = vbFormChildControl Then
            gIdentSpaces = gIdentSpaces - 1
            Call AddText("End")
            'Seek f, posNextControl
            GoTo NewControl
            
        End If
        
        If cControlEnd = vbFormExistingChildControl Then
            gIdentSpaces = gIdentSpaces - 1
            Call AddText("End")
            Dim bCheckEnd As Byte
            'MsgBox Loc(f)
            
            Do
                Get f, , bCheckEnd
                If bCheckEnd = 2 Then
                    gIdentSpaces = gIdentSpaces - 1
                    Call AddText("End")
                End If
                If bCheckEnd = 3 Then
                    gIdentSpaces = gIdentSpaces - 1
                    Call AddText("End")
                End If
            Loop Until bCheckEnd = 3 Or bCheckEnd = 4 Or bCheckEnd >= 5
               ' MsgBox "End Loop: " & Loc(f)
                If bCheckEnd <> 4 Then
                    GoTo NewControl
                End If
            
            ''bCheckEnd = GetByte2(f)
           
            
            ''If bCheckEnd = 4 Then

            ''Else
               '' Seek f, Loc(f) - 1
               ' Seek f, posNextControl
                ''GoTo NewControl
            ''End If
            
        End If
        If cControlEnd = vbFormMenu Then
          'Seek f, posNextControl
            GoTo NewControl
        End If
        If bFormEndUsed = False Then
            gIdentSpaces = gIdentSpaces - 1
            Call AddText("End")
        End If
    End If 'For gSkipCom
    
    Next lForm 'Main Form Loop
'##########################################
'End of Form/Control Properties Loop
'##########################################
End Sub

Private Sub txtCode_Change()
'*****************************
'Purpose: Color Coding for the Syntax
'*****************************
If gUpdateText = False Then Exit Sub
If gShowColors = False Then Exit Sub

        Dim StringRef As Long
        Dim Texte As String
        Dim StartString As Long
        Dim SelStart As Long
        Dim Cursor As Long
        Dim CharCursor As FIRSTCHAR_INFO
        Dim InstrColor As Long
        Dim FuncColor As Long
        Dim StartFind As Long
        Dim LengthFind As Long
        txtBuffer.Text = txtCode.Text
        
        SelStart = txtBuffer.SelStart
        txtBuffer.MousePointer = rtfHourglass
        
        
        Dim CountLine As Long
        CountLine = SendMessage(txtBuffer.hwnd, EM_GETLINECOUNT, 0, 0)

        Texte = txtBuffer.Text
        
        '======================    SelectionChanged    ========================='

        Dim NewsLines() As String
        Dim i As Long, lStartComp As Long, lEndComp As Long, tStartComp As Long, tEndComp As Long
        Dim TrueRtfText As String, BuffRTFText As String, BuffLines() As String
        TrueRtfText = txtBuffer.TextRTF
        
        '#########################################################################'
      
        NewsLines = Split(TrueRtfText, vbCrLf)      '<<<<<   absolument   <<<<<<<#'
        '#########################################################################'
        BuffLines = NewsLines
        
        For i = 0 To UBound(NewsLines)
            If Not i > UBound(LinesCheck) Then
                If NewsLines(i) <> LinesCheck(i) Then
                    Exit For
                End If
            Else
                Exit For
            End If
        Next i
        If i - 1 >= 0 Then
            ReDim Preserve BuffLines(0 To i - 1)
        Else
            ReDim Preserve BuffLines(0 To 0)
        End If
        BuffRTFText = Join(BuffLines, vbCrLf)
        buffCodeAv.TextRTF = BuffRTFText & "}"
        tStartComp = Len(buffCodeAv.Text)
        
        BuffRTFText = ""
        For i = 0 To UBound(NewsLines)
            If UBound(NewsLines) - i >= 0 Then
                If NewsLines(UBound(NewsLines) - i) <> LinesCheck(UBound(LinesCheck) - i) Then
                    Exit For
                End If
                BuffRTFText = NewsLines(UBound(NewsLines) - i) & BuffRTFText
            Else
                Exit For
            End If
        Next i
        buffCodeAp.TextRTF = NewsLines(0) & BuffRTFText
        tEndComp = Len(buffCodeAp.Text)
        
        If Len(Texte) - tEndComp - tStartComp > 0 Then
            Text1 = Mid(txtBuffer.Text, tStartComp + 1, Len(Texte) - tEndComp - tStartComp)
            StartFind = tStartComp
            LengthFind = Len(Texte) - tEndComp - tStartComp
        Else
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        

            StartFind = InStrRev(Texte, vbCrLf, IIf(SelStart = 0, 1, SelStart)) + 1
            If StartFind = 1 Then
                StartFind = 0
            End If
                
            If InStr(SelStart + 1, Texte, vbCrLf) = 0 Then
                LengthFind = Len(Texte)
            Else
                LengthFind = InStr(SelStart + 1, Texte, vbCrLf) - 1
            End If
            
            LengthFind = LengthFind - StartFind
        End If
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        '======================================================================='
        
        '======================= Section of Colors ========================='
        If LengthFind > 0 Then
            InstrColor = vbBlue '&H400000
            FuncColor = vbBlue Xor vbRed
            
            txtBuffer.SelStart = StartFind
            txtBuffer.SelLength = LengthFind 'Len(Texte)
            txtBuffer.SelColor = vbBlack
            txtBuffer.SelBold = False
            txtBuffer.SelItalic = False
            txtBuffer.SelUnderline = False
            
            'KeyWords to highlight
            ColorWord "beginproperty", InstrColor, txtBuffer, , StartFind + 1, LengthFind
            ColorWord "endproperty", InstrColor, txtBuffer, , StartFind + 1, LengthFind
            ColorWord "begin", InstrColor, txtBuffer, , StartFind + 1, LengthFind
            ColorWord "end", InstrColor, txtBuffer, , StartFind + 1, LengthFind
            ColorWord "sub", InstrColor, txtBuffer, "/B/I/", StartFind + 1, LengthFind
            ColorWord "public sub", InstrColor, txtBuffer, "/B/I/", StartFind + 1, LengthFind
            ColorWord "private sub", InstrColor, txtBuffer, "/B/I/", StartFind + 1, LengthFind
            ColorWord "public function", InstrColor, txtBuffer, "/B/I/", StartFind + 1, LengthFind
            ColorWord "private function", InstrColor, txtBuffer, "/B/I/", StartFind + 1, LengthFind
            ColorWord "function", InstrColor, txtBuffer, "/B/I/", StartFind + 1, LengthFind
            ColorWord "end sub", InstrColor, txtBuffer, "/B/I/", StartFind + 1, LengthFind
            ColorWord "end function", InstrColor, txtBuffer, "/B/I/", StartFind + 1, LengthFind
            ColorWord "dim", InstrColor, txtBuffer, , StartFind + 1, LengthFind
            ColorWord "if", InstrColor, txtBuffer, "/I/", StartFind + 1, LengthFind
            ColorWord "else", InstrColor, txtBuffer, "/I/", StartFind + 1, LengthFind
            ColorWord "elseif", InstrColor, txtBuffer, "/I/", StartFind + 1, LengthFind
            ColorWord "then", InstrColor, txtBuffer, "/I/", StartFind + 1, LengthFind
            ColorWord "end if", InstrColor, txtBuffer, "/I/", StartFind + 1, LengthFind
            ColorWord "goto", InstrColor, txtBuffer, , StartFind + 1, LengthFind
            ColorWord "while", InstrColor, txtBuffer, , StartFind + 1, LengthFind
            ColorWord "wend", InstrColor, txtBuffer, , StartFind + 1, LengthFind
            ColorWord "for", InstrColor, txtBuffer, , StartFind + 1, LengthFind
            ColorWord "next", InstrColor, txtBuffer, , StartFind + 1, LengthFind
            ColorWord "not", InstrColor, txtBuffer, , StartFind + 1, LengthFind
            ColorWord "print", FuncColor, txtBuffer, , StartFind + 1, LengthFind
            
            txtBuffer.SelStart = StartFind
            CharCursor = GetFirstChar(txtBuffer.SelStart + 1, txtBuffer, """'")
   
            While CharCursor.lCursor < StartFind + LengthFind And CharCursor.lCursor > 0
                

                    Select Case CharCursor.sChar
                        Case """"
                            Dim InStrFind As Long
                            InStrFind = txtBuffer.Find("""", CharCursor.lCursor) + 1
                            
                            InStrFind = IIf(InStrFind < txtBuffer.Find(vbCrLf, CharCursor.lCursor) + 1, InStrFind, txtBuffer.Find(vbCrLf, CharCursor.lCursor) + 1)
                            txtBuffer.SelStart = CharCursor.lCursor - 1
                            If InStrFind > 0 Then
                                txtBuffer.SelLength = InStrFind - CharCursor.lCursor + 1
                                CharCursor.lCursor = InStrFind
                            Else
                                txtBuffer.SelLength = Len(Texte) - CharCursor.lCursor + 1
                                CharCursor.lCursor = Len(Texte) + 1
                            End If
                            txtBuffer.SelBold = False
                            txtBuffer.SelItalic = False
                            txtBuffer.SelUnderline = False
                            txtBuffer.SelColor = &H80&
                        Case "'"
                            txtBuffer.SelStart = CharCursor.lCursor - 1
                                If InStr(CharCursor.lCursor + 1, Texte, vbCrLf) > 0 Then
                                Dim Buff As String, counter As Long, TheLen As Long
                                TheLen = InStr(CharCursor.lCursor + 1, Texte, vbCrLf) - CharCursor.lCursor + 1
                                Buff = Mid(txtBuffer.Text, txtBuffer.SelStart + 1)
                                counter = 1
                                While Mid(Trim(GetPart(Buff, counter, vbCrLf)), 1, 1) = "'"
                                    TheLen = TheLen + Len(GetPart(Buff, counter, vbCrLf)) '+ 2
                                    counter = counter + 1
                                Wend
                                txtBuffer.SelLength = TheLen
                                CharCursor.lCursor = InStr(CharCursor.lCursor + 1, Texte, vbCrLf) + IIf(counter > 1, TheLen, 0) + 1
                            Else
                                txtBuffer.SelLength = Len(Texte) - CharCursor.lCursor + 1
                                CharCursor.lCursor = Len(Texte)
                            End If
                            txtBuffer.SelBold = False
                            txtBuffer.SelItalic = True
                            txtBuffer.SelUnderline = False
                            txtBuffer.SelColor = &H8000&
                    End Select
                    CharCursor = GetFirstChar(CharCursor.lCursor + 1, txtBuffer, """'")
              
            Wend
        End If
        '======================================================================='
        
        LinesCheck = Split(txtBuffer.TextRTF, vbCrLf)
        


        txtBuffer.SelStart = SelStart
        txtBuffer.Refresh
        txtBuffer.MousePointer = rtfArrow
        DoEvents
        txtCode_SelChange

        prevCountLine = CountLines(txtBuffer)
        txtCode.TextRTF = txtBuffer.TextRTF
      
End Sub

Private Sub txtCode_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button > 0 Then txtCode_SelChange
End Sub

Public Function CountLines(rText As RichTextBox) As Long
'*****************************
'Purpose: Count the number of lines in a RichTextBox
'*****************************
    CountLines = SendMessage(rText.hwnd, &HBA, 0, 0)
End Function

Private Sub txtCode_SelChange()
    Dim CurrLine As Long
    On Error Resume Next
    CurrLine = SendMessage(txtBuffer.hwnd, &HC9, txtBuffer.SelStart, 0)
End Sub
Public Function ColorWord(Word As String, Color As Long, txtBox As RichTextBox, Optional Style As String, Optional ByVal lCursor As Long, Optional ByVal length As Long)
'*****************************
'Purpose: To color a Keyword
'*****************************
        Dim Arguments As Boolean
        Dim Cursor As Long
        Cursor = lCursor
        Cursor = txtBox.Find(Word, Cursor - 1, , rtfWholeWord) '- 1
        While IIf(length > 0, (Cursor < lCursor + length) And (Cursor > -1), Cursor > -1)
             
            txtBox.SelColor = Color
            txtBox.SelBold = IIf(UCase(Style) Like "*/B/*", True, False)
            txtBox.SelItalic = IIf(UCase(Style) Like "*/I/*", True, False)
            txtBox.SelUnderline = IIf(UCase(Style) Like "*/U/*", True, False)
            Cursor = txtBox.Find(Word, Cursor + 1, , rtfWholeWord)
        Wend
End Function
Sub GetControlSize(f As Variant)
'*****************************
'Purpose: Get the control size type
'*****************************
    Dim cPosition As typeStandardControlSize
    Get f, , cPosition
    Call AddText("Left = " & cPosition.cLeft)
    Call AddText("Top = " & cPosition.cTop)
    Call AddText("Height = " & cPosition.cHeight)
    Call AddText("Width = " & cPosition.cWidth)
                
            
End Sub
Sub GetFontProperty(f As Variant)
'*****************************
'Purpose: Get the font property type.
'*****************************
    Dim cFont As FontType
    Dim bItalic As Boolean, bUnderLine As Boolean, bStrike As Boolean
    bItalic = False
    bUnderLine = False
    bStrike = False
    gIdentSpaces = gIdentSpaces + 1
    Call AddText("BeginProperty Font")
    gIdentSpaces = gIdentSpaces + 1
    Get f, , cFont
                'MsgBox cFont.Weight
    Call AddText("Name = " & Chr(34) & gVBFile.GetString(Loc(f), cFont.FontLen) & Chr(34))
               ' FileLen = FileLen - Len(cFont)
    Call AddText("Size = " & (cFont.Size / 10000))
    Call AddText("Charset = " & cFont.un2)
    Call AddText("Weight = " & cFont.Weight)
                'FileLen = FileLen - cFont.FontLen
            'Font Property Opcodes
            'action=2 italic
            'action4=underline
            'action6 underline+italic=6
            'action10=italic+strickough
            'action=8=strikethough
            '12=underline +strikethough
            '14 =italic+underline+strkethough
                If cFont.Action = 2 Then
                    bItalic = True
                End If
                If cFont.Action = 4 Then
                    bUnderLine = True
                End If
                If cFont.Action = 6 Then
                    bUnderLine = True
                    bItalic = True
                End If
                If cFont.Action = 8 Then
                    bStrike = True
                End If
                If cFont.Action = 10 Then
                    bItalic = True
                    bStrike = True
                End If
                If cFont.Action = 12 Then
                    bUnderLine = True
                    bStrike = True
                End If
                If cFont.Action = 14 Then
                    bItalic = True
                    bUnderLine = True
                    bStrike = True
                End If
                
                
                If bItalic = True Then
                    Call AddText("Italic = -1")
                Else
                    Call AddText("Italic = 0")
                End If
                If bUnderLine = True Then
                    Call AddText("Underline = -1")
                Else
                    Call AddText("Underline = 0")
                End If
                If bStrike = True Then
                    Call AddText("Strikethrough = -1")
                Else
                    Call AddText("Strikethrough = 0")
                End If
                gIdentSpaces = gIdentSpaces - 1
                Call AddText("EndProperty")
                gIdentSpaces = gIdentSpaces - 1
              '  Seek f, Loc(f) - 1 ' - 2
                'Seek f, Loc(f) + 3
               ' MsgBox Loc(f)
               If gShowOffsets = True Then
               ' Call AddText("'Offset Font End: " & Loc(f) & " " & GetByte2(f))
               End If
End Sub
Sub GetDataFormatProperty(f As Variant)
'*****************************
'Purpose: Get the DataFormat type
'*****************************
Dim cDataFormat As DataFormatType
    gIdentSpaces = gIdentSpaces + 1
    Call AddText("BeginProperty Font")
    gIdentSpaces = gIdentSpaces + 1
    Get f, , cDataFormat
                
    Call AddText("Type = ")
    Call AddText("Format =" & Chr(34) & Chr(34))
    Call AddText("HaveTrueFalseNull = ")
    Call AddText("FirstDayOfWeek = ")
    Call AddText("FirstWeekOfYear = ")
    Call AddText("LCID = ")
    Call AddText("SubFormatType = ")
    
    gIdentSpaces = gIdentSpaces - 1
    Call AddText("EndProperty")
    gIdentSpaces = gIdentSpaces - 1
End Sub


Private Sub txtEditArray_Change(index As Integer)
'*****************************
'Purpose: Used to detect changes to the exe from the form editor
'*****************************
    Dim i As Integer
    Dim bUsed As Boolean
    bUsed = False
    If frmMain.txtEditArray(index).Tag = "Single" Then
         If IsNumeric(txtEditArray(index).Text) = False Then
            txtEditArray(index).Text = 0
         End If
        For i = 0 To UBound(SingleChange)
            If lblArrayEdit(index).Tag = SingleChange(i).offset Then
                SingleChange(i).sSingle = txtEditArray(index).Text
                bUsed = True
            End If
        Next
        If bUsed = False Then
            ReDim Preserve SingleChange(UBound(SingleChange) + 1)
            SingleChange(UBound(SingleChange)).offset = lblArrayEdit(index).Tag
            SingleChange(UBound(SingleChange)).sSingle = txtEditArray(index).Text
        End If
         
    End If
    If frmMain.txtEditArray(index).Tag = "String" Then
        For i = 0 To UBound(StringChange)
            If lblArrayEdit(index).Tag = StringChange(i).offset Then
                StringChange(i).sString = txtEditArray(index).Text
                bUsed = True
            End If
        Next
        If bUsed = False Then
            ReDim Preserve StringChange(UBound(StringChange) + 1)
            StringChange(UBound(StringChange)).offset = lblArrayEdit(index).Tag
            StringChange(UBound(StringChange)).sString = txtEditArray(index).Text
        End If
        
    End If
    If frmMain.txtEditArray(index).Tag = "Boolean" Then
        For i = 0 To UBound(BooleanChange)
            If lblArrayEdit(index).Tag = BooleanChange(i).offset Then
            
                'BooleanChange(i).bBool = txtEditArray(index).Text
                If LCase(txtEditArray(index).Text) = "true" Then
                    BooleanChange(UBound(BooleanChange)).bBool = True
                End If
                If LCase(txtEditArray(index).Text) = "false" Then
                    BooleanChange(UBound(BooleanChange)).bBool = False
                End If
                bUsed = True
            End If
        Next
        If bUsed = False Then
            ReDim Preserve BooleanChange(UBound(BooleanChange) + 1)
            BooleanChange(UBound(BooleanChange)).offset = lblArrayEdit(index).Tag
            
            If LCase(txtEditArray(index).Text) = "true" Then
                BooleanChange(UBound(BooleanChange)).bBool = True
            End If
            If LCase(txtEditArray(index).Text) = "false" Then
                BooleanChange(UBound(BooleanChange)).bBool = False
            End If
            
        End If
    End If
    If frmMain.txtEditArray(index).Tag = "Long" Then
         If IsNumeric(txtEditArray(index).Text) = False Then
            txtEditArray(index).Text = 0
         End If
        For i = 0 To UBound(LongChange)
            If lblArrayEdit(index).Tag = LongChange(i).offset Then
                LongChange(i).lLong = txtEditArray(index).Text
                bUsed = True
            End If
        Next
        If bUsed = False Then
            ReDim Preserve LongChange(UBound(LongChange) + 1)
            LongChange(UBound(LongChange)).offset = lblArrayEdit(index).Tag
            LongChange(UBound(LongChange)).lLong = txtEditArray(index).Text
        End If
         
    End If
    If frmMain.txtEditArray(index).Tag = "Integer" Then
         If IsNumeric(txtEditArray(index).Text) = False Then
            txtEditArray(index).Text = 0
        Else
            If txtEditArray(index).Text < -32000 Or txtEditArray(index).Text > 32000 Then
                txtEditArray(index).Text = 0
            End If
        End If
        For i = 0 To UBound(IntegerChange)
            If lblArrayEdit(index).Tag = IntegerChange(i).offset Then
                IntegerChange(i).iInt = txtEditArray(index).Text
                bUsed = True
            End If
        Next
        If bUsed = False Then
            ReDim Preserve IntegerChange(UBound(IntegerChange) + 1)
            IntegerChange(UBound(IntegerChange)).offset = lblArrayEdit(index).Tag
            IntegerChange(UBound(IntegerChange)).iInt = txtEditArray(index).Text
        End If

    End If
    If frmMain.txtEditArray(index).Tag = "Byte" Then
        If IsNumeric(txtEditArray(index).Text) = False Then
            txtEditArray(index).Text = 0
        Else
            If txtEditArray(index).Text < 0 Or txtEditArray(index).Text > 255 Then
                txtEditArray(index).Text = 0
            End If
        End If

        For i = 0 To UBound(ByteChange)
            If lblArrayEdit(index).Tag = ByteChange(i).offset Then
                ByteChange(i).bByte = txtEditArray(index).Text
                bUsed = True
            End If
        Next
        If bUsed = False Then
            ReDim Preserve ByteChange(UBound(ByteChange) + 1)
            ByteChange(UBound(ByteChange)).offset = lblArrayEdit(index).Tag
            ByteChange(UBound(ByteChange)).bByte = txtEditArray(index).Text
        End If
    
    End If

    mnuFileSaveExe.Enabled = True
End Sub

