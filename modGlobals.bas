Attribute VB_Name = "modGlobals"
'#############################################
'modGlobals vbgamer45 2004
'##############################################
'Notes
'################################################
'"a" - means it is an Address
'"o" - means it is a relative Offset
'"Unknown" - self explanatory
'"Flag" - Variable Unknown Property
'"Const" - Constant Unknown Property
'"Address" - Unknown Address
'################################################
Const Signature = &H1F4 ' F4 01 00 00
Const MAX_PATH = 260
Public Const Version = "0.03"


Type VBHeader
    Signature               As String * 4  '00h 00d
    'VB5! identifier &quot;VB5!&quot;
    
   RuntimeBuild                 As Integer     '04h 04d
     'RuntimeBuild
  LanguageDLL             As String * 14 '06h 06d
    'Language DLL name. _
     0x2A meaning default or null terminated string.
  
  BackupLanguageDLL       As String * 14 '14h 20d
    'Backup Language DLL name. _
     0x7F meaning default or null terminated string. _
     Changing values do not effect working status of an exe.
  
  RuntimeDLLVersion       As Integer     '22h 34d
    'Run-time DLL version
  
  LanguageID              As Long        '24h 36d
  
  BackupLanguageID        As Long        '28h 40d
    'Backup Language ID &#40;only when Language DLL exists&#41;
  
  aSubMain                As Long        '2Ch 44d
    'Address to Sub Main&#40;&#41; code _
     &#40;If 0000 0000 then it's a load form call&#41;
  
  aProjectInfo            As Long        '30h 48d
    
  fMDLIntObjs                 As Long    '34h 52d

  fMDLIntObjs2               As Long        '38h 56d

  ThreadFlags           As Long        '3Ch 60d

  ThreadCount               As Long        '40h 64d
  
  
  FormCount               As Integer     '44h 68d
  
  ExternalComponentCount  As Integer     '46h 70d
    'Number of external components &#40;eg. winsock&#41; referenced
    
  ThunkCount As Long
  
  aGUITable               As Long        '4Eh 78d
  aExternalComponentTable As Long        '52h 82d
  'aProjectDescription     As Long        '56h 86d
  aComRegisterData         As Long
  
  oProjectExename         As Long        '5Ah 90d
  oProjectTitle           As Long        '5Eh 94d
  oHelpFile               As Long        '62h 98d
  oProjectName            As Long        '66h 102d
End Type

'* Thread Flags:
'+-------+----------------+--------------------------------------------------------+
'| Value | Name           | Description                                            |
'+-------+----------------+--------------------------------------------------------+
'|  0x01 | ApartmentModel | Specifies multi-threading using an apartment model     |
'|  0x02 | RequireLicense | Specifies to do license validation (OCX only)          |
'|  0x04 | Unattended     | Specifies that no GUI elements should be initialized   |
'|  0x08 | SingleThreaded | Specifies that the image is single-threaded            |
'|  0x10 | Retained       | Specifies to keep the file in memory (Unattended only) |
'+-------+----------------+--------------------------------------------------------+
'ex: A value of 0x15 specifies a multi-threaded, memory-resident ActiveX Object with no GUI


'* MDL Internal Object Flags
'+---------+------------+---------------+
'| Ctrl ID |      Value | Object Name   |
'+---------+------------+---------------+
'|                           First Flag |
'+---------+------------+---------------+
'|    0x00 | 0x00000001 | PictureBox    |
'|    0x01 | 0x00000002 | Label         |
'|    0x02 | 0x00000004 | TextBox       |
'|    0x03 | 0x00000008 | Frame         |
'|    0x04 | 0x00000010 | CommandButton |
'|    0x05 | 0x00000020 | CheckBox      |
'|    0x06 | 0x00000040 | OptionButton  |
'|    0x07 | 0x00000080 | ComboBox      |
'|    0x08 | 0x00000100 | ListBox       |
'|    0x09 | 0x00000200 | HScrollBar    |
'|    0x0A | 0x00000400 | VScrollBar    |
'|    0x0B | 0x00000800 | Timer         |
'|    0x0C | 0x00001000 | Print         |
'|    0x0D | 0x00002000 | Form          |
'|    0x0E | 0x00004000 | Screen        |
'|    0x0F | 0x00008000 | Clipboard     |
'|    0x10 | 0x00010000 | Drive         |
'|    0x11 | 0x00020000 | Dir           |
'|    0x12 | 0x00040000 | FileListBox   |
'|    0x13 | 0x00080000 | Menu          |
'|    0x14 | 0x00100000 | MDIForm       |
'|    0x15 | 0x00200000 | App           |
'|    0x16 | 0x00400000 | Shape         |
'|    0x17 | 0x00800000 | Line          |
'|    0x18 | 0x01000000 | Image         |
'|    0x19 | 0x02000000 | Unsupported   |
'|    0x1A | 0x04000000 | Unsupported   |
'|    0x1B | 0x08000000 | Unsupported   |
'|    0x1C | 0x10000000 | Unsupported   |
'|    0x1D | 0x20000000 | Unsupported   |
'|    0x1E | 0x40000000 | Unsupported   |
'|    0x1F | 0x80000000 | Unsupported   |
'+---------+------------+---------------+
'|                          Second Flag |
'+---------+------------+---------------+
'|    0x20 | 0x00000001 | Unsupported   |
'|    0x21 | 0x00000002 | Unsupported   |
'|    0x22 | 0x00000004 | Unsupported   |
'|    0x23 | 0x00000008 | Unsupported   |
'|    0x24 | 0x00000010 | Unsupported   |
'|    0x25 | 0x00000020 | DataQuery     |
'|    0x26 | 0x00000040 | OLE           |
'|    0x27 | 0x00000080 | Unsupported   |
'|    0x28 | 0x00000100 | UserControl   |
'|    0x29 | 0x00000200 | PropertyPage  |
'|    0x2A | 0x00000400 | Document      |
'|    0x2B | 0x00000800 | Unsupported   |
'+---------+------------+---------------+
'ex: A value of 0x30F000 (the so called "static binary constant on most sites") actually means to initialize the Print, Form, Screen, ClipBoard Objects (0xF000) as well as the Drive/Dir Objects (0x30000). This is the default on VB projects because those objects can always be accessed from a module (ie, they are not graphic, except Forms, wich can always be created)

'COM Data Types
'
Type tCOMRegData
  oRegInfo                As Long    ' 0x00 (00d) Offset to COM Interfaces Info
  oNTSProjectName         As Long    ' 0x04 (04d) Offset to Project/Typelib Name
  oNTSHelpDirectory       As Long    ' 0x08 (08d) Offset to Help Directory
  oNTSProjectDescription  As Long    ' 0x0C (12d) Offset to Project Description
  uuidProjectClsId(15)    As Byte    ' 0x10 (16d) CLSID of Project/Typelib
  lTlbLcid                As Long    ' 0x20 (32d) LCID of Type Library
  iPadding1               As Integer ' 0x24 (36d)
  iTlbVerMajor            As Integer ' 0x26 (38d) Typelib Major Version
  iTlbVerMinor            As Integer ' 0x28 (40d) Typelib Minor Version
  iPadding2               As Integer ' 0x2A (42d)
  lPadding3               As Long    ' 0x2C (44d)
                                     ' 0x30 (48d) <- Structure Size
End Type

Type tCOMRegInfo
  oNextObject          As Long    ' 0x00 (00d) Offset to COM Interfaces Info
  oObjectName          As Long    ' 0x04 (04d) Offset to Object Name
  oObjectDescription   As Long    ' 0x08 (08d) Offset to Object Description
  lInstancing          As Long    ' 0x0C (12d) Instancing Mode
  lObjectID            As Long    ' 0x10 (16d) Current Object ID in the Project
  uuidObjectClsID(15)  As Byte    ' 0x14 (20d) CLSID of Object
  fIsInterface         As Long    ' 0x24 (36d) Specifies if the next CLSID is valid
  oObjectClsID         As Long    ' 0x28 (40d) Offset to CLSID of Object Interface
  oControlClsID        As Long    ' 0x2C (44d) Offset to CLSID of Control Interface
  fIsControl           As Long    ' 0x30 (48d) Specifies if the CLSID above is valid
  lMiscStatus          As Long    ' 0x34 (52d) OLEMISC Flags (see MSDN docs)
  fClassType           As Byte    ' 0x38 (56d) Class Type
  fObjectType          As Byte    ' 0x39 (57d) Flag identifying the Object Type
  iToolboxBitmap32     As Integer ' 0x3A (58d) Control Bitmap ID in Toolbox
  iDefaultIcon         As Integer ' 0x3C (60d) Minimized Icon of Control Window
  fIsDesigner          As Integer ' 0x3E (62d) Specifies whether this is a Designer
  oDesignerData        As Long    ' 0x40 (64d) Offset to Designer Data
                                  ' 0x44 (68d) <-- Structure Size
End Type
'Object Type part of tCOMRegInfo
'+-------+---------------+-------------------------------------------+
'| Value | Name          | Description                               |
'+-------+---------------+-------------------------------------------+
'|  0x02 | Designer      | A Visual Basic Designer for an Add.in     |
'|  0x10 | Class Module  | A Visual Basic Class                      |
'|  0x20 | User Control  | A Visual Basic ActiveX User Control (OCX) |
'|  0x80 | User Document | A Visual Basic User Document              |
'+-------+---------------+-------------------------------------------+

Type tDesignerInfo
  uuidDesigner(15)       As Byte    '0x00 (00d)                           CLSID of the Addin/Designer
  lStructSize            As Long    '0x10 (16d)                           Total Size of the next fields
  
  iSizeAddinRegKey       As Integer '0x14 (20d)
  sAddinRegKey           As String  '0x16 (22d)                           Registry Key of the Addin
  
  iSizeAddinName         As Integer '0x16 (22d) + iSizeAddinRegKey
  sAddinName             As String  '0x18 (24d) + iSizeAddinRegKey        Friendly Name of the Addin
  
  iSizeAddinDescription  As Integer '0x18 (24d) + iSizeAddinRegKey _
                                                + iSizeAddinName
  iAddinDescription      As String  '0x1A (26d) + iSizeAddinRegKey _
                                                + iSizeAddinName          Description of Addin
  
  lLoadBehaviour         As Long    '0x1A (26d) + iSizeAddinRegKey _
                                                + iSizeAddinName _
                                                + iSizeAddinDescription   CLSID of Object
  
  iSizeSatelliteDLL      As Integer '0x1E (30d) + iSizeAddinRegKey _
                                                + iSizeAddinName _
                                                + iSizeAddinDescription
  sSatelliteDLL          As String  '0x20 (32d) + iSizeAddinRegKey _
                                                + iSizeAddinName _
                                                + iSizeAddinDescription   SatelliteDLL, if specified
  
  iSizeAdditionalRegKey  As Integer '0x20 (32d) + iSizeAddinRegKey _
                                                + iSizeAddinName _
                                                + iSizeAddinDescription _
                                                + iSizeSatteliteDLL
  sAdditionalRegKey      As String  '0x22 (34d) + iSizeAddinRegKey _
                                                + iSizeAddinName _
                                                + iSizeAddinDescription _
                                                + iSizeSatteliteDLL       Extra Registry Key, if specified
  
  lCommandLineSafe       As Long    '0x22 (34d) + iSizeAddinRegKey _
                                                + iSizeAddinName _
                                                + iSizeAddinDescription _
                                                + iSizeSatteliteDLL _
                                                + iSizeAdditionalRegKey   Specifies a GUI-less Addin if 1
                                    '0x14 + lStructSize  <-- Structure Size
End Type

Private Type tProjectInfo

  Signature As Long                            ' 0x00
  aObjectTable As Long                         ' 0x04
  Null1 As Long                                ' 0x08
  aStartOfCode As Long                         ' 0x0C
  aEndOfCode As Long                           ' 0x10
  Flag1 As Long                                ' 0x14
  ThreadSpace As Long                          ' 0x18
  aVBAExceptionhandler  As Long                ' 0x1C
  aNativeCode As Long                          ' 0x20
  oProjectLocation As Integer                  ' 0x24
  Flag2 As Integer                             ' 0x26
  Flag3 As Integer                             ' 0x28

  OriginalPathName(MAX_PATH * 2) As Byte       ' 0x2A
  NullSpacer As Byte                           ' 0x233
  aExternalTable As Long                       ' 0x234
  ExternalCount As Long                        ' 0x238

' Size 0x23C
End Type

Private Type tObject
    aObjectInfo As Long         ' 0x00
    Const1 As Long              ' 0x04
    aPublicBytes As Long        ' 0x08 (08d) Pointer to Public Variable Size integers
    aStaticBytes As Long        ' 0x0C (12d) Pointer to Static Variables Struct
    aModulePublic As Long       ' 0x10 (16d) Memory Pointer to Public Variables
    aModuleStatic As Long       ' 0x14 (20d) Pointer to Static Variables
    aObjectName As Long         ' 0x18  NTS
    ProcCount As Long           ' 0x1C events, funcs, subs
    aProcNamesArray As Long     ' 0x20 when non-zero
    oStaticVars As Long         ' 0x24 (36d) Offset to Static Vars from aModuleStatic
    ObjectType As Long          ' 0x28
    Null3 As Long               ' 0x2C
                                ' 0x30  <-- Structure Size
End Type

'tObject.ObjectTyper Properties...
'#########################################################
'form&#58;              0000 0001 1000 0000 1000 0011 --&gt; 18083
'                   0000 0001 1000 0000 1010 0011 --&gt; 180A3
'                   0000 0001 1000 0000 1100 0011 --&gt; 180C3
'module&#58;            0000 0001 1000 0000 0000 0001 --&gt; 18001
'                   0000 0001 1000 0000 0010 0001 --&gt; 18021
'class&#58;             0001 0001 1000 0000 0000 0011 --&gt; 118003
'                   0001 0011 1000 0000 0000 0011 --&gt; 138003
'                   0000 0001 1000 0000 0010 0011 --&gt; 18023
'                   0000 0001 1000 1000 0000 0011 --&gt; 18803
'                   0001 0001 1000 1000 0000 0011 --&gt; 118803
'usercontrol&#58;       0001 1101 1010 0000 0000 0011 --&gt; 1DA003
'                  0001 1101 1010 0000 0010 0011 --&gt; 1DA023
'                  0001 1101 1010 1000 0000 0011 --&gt; 1DA803
'propertypage&#58;      0001 0101 1000 0000 0000 0011 --&gt; 158003
'                      | ||     |  |    | |    |
'&#91;moog&#93;                | ||     |  |    | |    |
'HasPublicInterface ---+ ||     |  |    | |    |
'HasPublicEvents --------+|     |  |    | |    |
'IsCreatable/Visible? ----+     |  |    | |    |
'Same as &quot;HasPublicEvents&quot; -----+  |    | |    |
'&#91;aLfa&#93;                         |  |    | |    |
'usercontrol &#40;1&#41; ---------------+  |    | |    |
'ocx/dll &#40;1&#41; ----------------------+    | |    |
'form &#40;1&#41; ------------------------------+ |    |
'vb5 &#40;1&#41; ---------------------------------+    |
'HasOptInfo &#40;1&#41; -------------------------------+
'                                              |
'module&#40;0&#41; ------------------------------------+

Public Type tObjectInfo
    Flag1 As Integer       ' 0x00
    ObjectIndex As Integer ' 0x02
    aObjectTable As Long   ' 0x04
    Null1 As Long          ' 0x08
    aSmallRecord   As Long ' 0x0C  when it is a module this value is -1 [better name?]
    Const1 As Long         ' 0x10
    Null2 As Long          ' 0x14
    aObject As Long        ' 0x18
    RunTimeLoaded  As Long ' 0x1C [can someone verify this?]
    NumberOfProcs  As Long ' 0x20
    aProcTable As Long     ' 0x24
    iConstantsCount As Integer '0x28 (40d) Number of Constants
    iMaxConstants   As Integer '0x2A (42d) Maximum Constants to allocate.
    Flag5 As Long          ' 0x2C
    Flag6 As Integer       ' 0x30
    Flag7 As Integer       ' 0x32
    aConstantPool As Long  ' 0x34
                           ' 0x38 <-- Structure Size
                           'the rest is optional items[OptionalObjectInfo]
End Type
Private Type tObjectTable
    lNull1 As Long          ' 0x00 (00d)
    aExecProj As Long       ' 0x04 (04d) Pointer to a memory structure
    aProjectInfo2 As Long   ' 0x08 (08d) Pointer to Project Info 2
    Const1 As Long          ' 0x0C
    Null2 As Long           ' 0x10
    lpProjectObject As Long ' 0x14
    Flag1 As Long           ' 0x18
    Flag2 As Long           ' 0x1C
    Flag3 As Long           ' 0x20
    Flag4 As Long           ' 0x24
    fCompileType As Integer ' 0x28 (40d) Internal flag used during compilation
    ObjectCount1 As Integer ' 0x2A
    iCompiledObjects As Integer ' 0x2C (44d) Number of objects compiled.
    iObjectsInUse As Integer ' 0x2E (46d) Updated in the IDE to correspond the total number ' but will go up or down when initializing/unloading modules.
    aObject As Long         ' 0x30
    Null3 As Long           ' 0x34
    Null4 As Long           ' 0x38
    Null5 As Long           ' 0x3C
    aProjectName As Long    ' 0x40      NTS
    LangID1  As Long        ' 0x44
    LangID2  As Long        ' 0x48
    Null6  As Long          ' 0x4C
    Const3  As Long         ' 0x50
                            ' 0x54
End Type
Type ExternalTable
   flag As Long        '0x00
   aExternalLibrary As Long  '0x04
End Type

Type ExternalLibrary
   aLibraryName As Long     '0x00   points to NTS
   aLibraryFunction As Long '0x04   points to NTS
End Type

Public Type tEventLink

    Const1 As Integer        ' 0x00
    CompileType As Byte      ' 0x02
    aEvent As Long           ' 0x03
    PushCmd As Byte          ' 0x07
    pushAddress As Long      ' 0x08
    Const As Byte            ' 0x0C
                             ' 0x0D&lt;-- Structure Size
End Type
Private Type tEventTable
    Null1 As Long                                  ' 0x00
    aControl As Long                               ' 0x04
    aObjectInfo As Long                            ' 0x08
    aQueryInterface As Long                        ' 0x0C
    aAddRef As Long                                ' 0x10
    aRelease As Long                                ' 0x14
    'aEventPointer() As Long
    'aEventPointer(aControl.EventCount - 1) As Long ' 0x18
End Type
Global taEventPointer() As Long

Private Type tOptionalObjectInfo ' if &#40;&#40;tObject.ObjectType AND &amp;H80&#41;=&amp;H80&#41;

    fDesigner As Long              ' 0x00 (0d) If this value is 2 then this object is a designer
    aObjectCLSID As Long           ' 0x04
    Null1 As Long                  ' 0x08
    aGuidObjectGUI As Long         ' 0x0C
    lObjectDefaultIIDCount As Long ' 0x10  01 00 00 00
    aObjectEventsIIDTable As Long  ' 0x14
    lObjectEventsIIDCount As Long  ' 0x18
    aObjectDefaultIIDTable As Long ' 0x1C
    ControlCount As Long           ' 0x20
    aControlArray As Long          ' 0x24
    iEventCount As Integer         ' 0x28 (40d) Number of Events
    iPCodeCount As Integer         ' 0x2C
    oInitializeEvent As Integer    ' 0x2C (44d) Offset to Initialize Event from aMethodLinkTable
    oTerminateEvent As Integer     ' 0x2E (46d) Offset to Terminate Event from aMethodLinkTable
    aEventLinkArray As Long        ' 0x30  Pointer to pointers of MethodLink
    aBasicClassObject As Long      ' 0x34 Pointer to an in-memory
    Null3 As Long                  ' 0x38
    Flag2 As Long                  ' 0x3C usually null
                                   ' 0x40 &lt;-- Structure size
End Type
Public Type tEventPointer
    Const1 As Byte      ' 0x00
    Flag1 As Long       ' 0x01
    Const2 As Long      ' 0x05
    Const3 As Byte      ' 0x09
    aEvent As Long      ' 0x0A
                        ' 0x0E &lt;-- Structure Size
End Type

Public Type tCodeInfo
    aObjectInfo As Long     ' 0x00
    Flag1 As Integer        ' 0x04
    Flag2 As Integer        ' 0x06
    CodeLength As Integer   ' 0x08
    Flag3 As Long           ' 0x0A
    Flag4 As Integer        ' 0x0E
    Null1 As Integer        ' 0x10
    Flag5 As Long           ' 0x12
    Flag6 As Integer        ' 0x16
                            ' 0x18  &lt;-- Structure Size
End Type

Private Type tControl
    Flag1 As Integer        ' 0x00
    EventCount As Integer   ' 0x02
    Flag2 As Long           ' 0x04
    aGUID As Long           ' 0x08
    index As Integer        ' 0x0C
    Const1 As Integer       ' 0x0E
    Null1 As Long           ' 0x10
    Null2 As Long           ' 0x14
    aEventTable As Long     ' 0x18
    Flag3 As Byte           ' 0x1C
    Const2 As Byte          ' 0x1D
    Const3 As Integer       ' 0x1E
    aName As Long           ' 0x20
    Index2 As Integer       ' 0x24
    Const1Copy As Integer   ' 0x26
                            ' 0x28  &lt;-- Structure Size
End Type

Private Type oldtGuiTable
    SectionHeader As Long
    unknown(59) As Byte
    FormSize As Long
    un1 As Long
    aFormPointer As Long
    un2 As Long
    
End Type
Type tGuiTable
  lStructSize          As Long ' 0x00 (00d) Total size of this structure
  uuidObjectGUI(15)    As Byte ' 0x04 (04d) UUID of Object GUI
  Unknown1             As Long ' 0x14 (20d)
  Unknown2             As Long ' 0x18 (24d)
  Unknown3             As Long ' 0x1C (28d)
  Unknown4             As Long ' 0x20 (32d)
  lObjectID            As Long ' 0x24 (36d) Current Object ID in the Project
  Unknown5             As Long ' 0x28 (40d)
  fOLEMisc             As Long ' 0x2C (44d) OLEMisc Flags
  uuidObject(15)       As Byte ' 0x30 (48d) UUID of Object
  Unknown6             As Long ' 0x40 (64d)
  Unknown7             As Long ' 0x44 (68d)
  aFormPointer         As Long ' 0x48 (72d) Pointer to GUI Object Info
  Unknown8             As Long ' 0x4C (76d)
                               ' 0x50 (80d) <- Structure Size
End Type


Public Type tComponent
  StructLength   As Long
  l1             As Long
  l2             As Long
  l3             As Long
  l4             As Long
  l5             As Long
  l6             As Long
  GUIDoffset     As Long
  GUIDlength     As Long
  l7             As Long
  FileNameOffset As Long
  SourceOffset   As Long
  NameOffset     As Long
End Type
'If GUIDlength = -1 then there is no oUUID
'If GUIDlength = 72 then read a unicode UUID

Type tProjectInfo2
  lNull1                  As Long ' 0x00 (00d)
  aObjectTable            As Long ' 0x04 (04d) Pointer to Object Table
  lConst1                 As Long ' 0x08 (08d)
  lNull2                  As Long ' 0x0C (12d)
  aObjectDescriptorTable  As Long ' 0x10 (16d) Pointer to a table of ObjectDescriptors
  lNull3                  As Long ' 0x14 (20d)
  aNTSPrjDescription      As Long ' 0x18 (24d) Pointer to Project Description
  aNTSPrjHelpFile         As Long ' 0x1C (28d) Pointer to Project Help File
  lConst2                 As Long ' 0x20 (32d)
  lHelpContextID          As Long ' 0x24 (36d) Project Help Context ID
                                  ' 0x28 (40d) <- Structure size
End Type

Type ObjectDescriptor
  lNull1      As Long '0x00 (00d)
  aObjectInfo As Long '0x04 (04d) Pointer to Object Info
  lConst1     As Long '0x08 (08d)
  lNull2      As Long '0x0C (12d)
  lFlag1      As Long '0x10 (16d)
  lNull3      As Long '0x14 (20d)
  aUnknown1   As Long '0x18 (24d)
  lNull4      As Long '0x1C (28d)
  aUnknown2   As Long '0x20 (32d)
  aUnknown3   As Long '0x24 (36d)
  aUnknown4   As Long '0x28 (40d)
  lNull5      As Long '0x2C (44d)
  lNull6      As Long '0x30 (48d)
  lNull7      As Long '0x34 (52d)
  lFlag2      As Long '0x38 (56d)
  fObjectType As Long '0x3C (60d) Flags for this Object
                      '0x40 (64d) <- Structure Size
End Type

Type MethodLinkNative
  jmpOpCode As Byte '0x0 (0d)
  jmpOffset As Long '0x1 (1d) jmp <address>  ; <address> = <currentoffset> + <jmpOffset> + 5
                    '0x5 (5d) <-- Structure Size
End Type

Type MethodLinkPCode
  xorOpCode   As Integer '0x0 (00d) xor eax, eax
  movOpCode   As Byte    '0x2 (02d)
  movAddress  As Long    '0x3 (03d) mov edx, <movAddress>
  pushOpCode  As Byte    '0x7 (07d)
  pushAddress As Long    '0x8 (08d) push <pushAddress>
  retOpCode   As Byte    '0xC (12d) ret
                         '0xD (13d) <-- Structure Size
End Type

Type GUIObjectInfo
  lUnknown1            As Long ' 0x00 (00d)
  bUnknown2            As Byte ' 0x04 (04d)
  guidObjectGUI(15)    As Byte ' 0x05 (05d) GUID of this ObjectGUI
  uuidUnknown1(15)     As Byte ' 0x15 (21d)
  guidCOMEventsIID(15) As Byte ' 0x25 (37d) GUID of this object EventsIID
  lUnknown3            As Long ' 0x35 (53d)
  lUnknown4            As Long ' 0x39 (57d)
  lUnknown5            As Long ' 0x3D (61d)
  lUnknown6            As Long ' 0x41 (65d)
  lUnknown7            As Long ' 0x45 (69d)
  lUnknown8            As Long ' 0x49 (73d)
  lUnknown9            As Long ' 0x4D (77d)
  lUnknown10           As Long ' 0x51 (81d)
  lUnknown11           As Long ' 0x55 (85d)
  lPropertiesLength    As Long ' 0x59 (89d) Total Length of Properties
                               ' 0x5D (93d) <-- Structure Size
End Type

Private Type typeApiList
    strLibraryName As String
    strFunctionName As String
End Type

Private Type typeProcedureList
    strParent As String
    strProcedureName As String
End Type

'Globals begin
Global gProcedureList() As typeProcedureList
Global gApiList() As typeApiList
Global gVBHeader As VBHeader
'Com Stuff
Global gCOMRegInfo As tCOMRegInfo
Global gCOMRegData As tCOMRegData
Global gDesignerInfo As tDesignerInfo

Global gProjectInfo As tProjectInfo
Global gObjectTable As tObjectTable
Global gObject() As tObject
Global gObjectInfo As tObjectInfo
Global gExternalTable As ExternalTable
Global gExternalLibrary As ExternalLibrary
Global gOptionalObjectInfo As tOptionalObjectInfo
Global gEventLink As tEventLink
Global gEventPointer As tEventPointer
Global gControl() As tControl
Global gEventTable() As tEventTable
Global gCodeInfo As tCodeInfo
Global gProcedure()  As Long 'As tProcedure
Global gGuiTable() As tGuiTable
Global gObjectNameArray() As String
Global gObjectProcCountArray() As Integer
Global gObjectInfoHolder() As tObjectInfo
Private Type tObjectOffsetType
    Address As Long
    ObjectName As String
End Type
Global gObjectOffsetArray() As tObjectOffsetType

'Options
Global gSkipCom As Boolean
Global gDumpData As Boolean
Global gShowOffsets As Boolean
Global gShowColors As Boolean
Global gPcodeDecompile As Boolean

Private Type typeControlName
    strParentForm As String
    strControlName As String
    strGuid As String
    bControlImage As Byte
End Type
Global gControlNameArray() As typeControlName


'For Controls
Public Type typeStandardControlSize
    cLeft As Integer
    cTop As Integer
    cWidth As Integer
    cHeight As Integer
End Type

'Picture Header
Public Type typePictureHeader
    un1 As Integer
    un2 As Integer
    un3 As Integer
    un4 As Integer
End Type

'COM Fix Type to fix some COM problems
Private Type typeCOMFIX
    ObjectName As String
    PropertyName As String
    NewType As String
End Type
Global gComFix() As typeCOMFIX

Private Type ImportListtype
    strName As String
    strGuid As String
    strLib As String
End Type
Global ImportList() As ImportListtype

'Used for Memory Map
Public gVBFile As clsFile
Public gMemoryMap As clsMemoryMap

'Variables for .vbp file
Global ProjectExename As String                     ' Project exename. MaxLength: 0x104 (260d)
Global ProjectTitle As String                       ' Project title. MaxLength: 0x28 (40d)
Global HelpFile As String                           ' Helpfile. MaxLength: 0x28 (40d)
Global ProjectName As String                        ' Project name. MaxLength: 0x104 (260d)
Global ProjectDescription As String

'Get File Information File Version Properties
Public Type FILEPROPERTIE
    CompanyName As String
    FileDescription As String
    FileVersion As String
    InternalName As String
    LegalCopyright As String
    OrigionalFileName As String
    ProductName As String
    ProductVersion As String
    LanguageID As String
End Type
Global gFileInfo As FILEPROPERTIE
Declare Function GetFileVersionInfo Lib "Version.dll" Alias _
   "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal _
   dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias _
   "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, _
   lpdwHandle As Long) As Long
Declare Function VerQueryValue Lib "Version.dll" Alias _
   "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, _
   lplpBuffer As Any, puLen As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias _
   "GetSystemDirectoryA" (ByVal path As String, ByVal cbBytes As _
   Long) As Long
Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Dest As Any, ByVal Source As Long, ByVal length As Long)
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" ( _
    ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Const LANG_ENGLISH = &H9

Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Public Type FIRSTCHAR_INFO
    sChar As String
    lCursor As Long
End Type

Global gUpdateText As Boolean 'Update Syntax Coloring?
Global gIdentSpaces As Integer 'For identing code
Global gDllProject As Boolean 'Is it a dll project?
Global CancelDecompile As Boolean

Sub PrintReadMe()
    '*****************************
    'Prints the ReadMe of the program
    '*****************************
    On Error Resume Next
    Kill (App.path & "\readme.txt")
   
    Open App.path & "\ReadMe.txt" For Output As #1
        Print #1, "-------------------------------"
        Print #1, "Semi VB Decompiler by vbgamer45"
        Print #1, "Open Source"
        Print #1, "Version: " & Version
        Print #1, "-------------------------------"
        Print #1, "Contents"
        Print #1, "1. What's New?"
        Print #1, "2. Features"
        Print #1, "3. Questions?"
        Print #1, "4. Bugs"
        Print #1, "5. Contact"
        Print #1, "6. Credits"
        Print #1, ""
        Print #1, "1. What's New?"
        Print #1, ""
        Print #1, "   Version 0.03"
        Print #1, "     P-Code decoding started and image extraction."
        Print #1, "     Numerous bug fixes."
        Print #1, "     Event detection added."
        Print #1, "     Dll and OCX Support added."
        Print #1, "     External Components added to vbp file."
        Print #1, "     Begun work on a basic antidecompiler."
        Print #1, "     Form property editor, complete with a patch report generator."
        Print #1, "     Procedure names are recovered."
        Print #1, "     Api's used by the program are recovered."
        Print #1, "     Msvbvm60.dll imports are listed in the treeview."
        Print #1, "     Syntax coloring for Forms."
        Print #1, "     Fixed scrolling bug."
        Print #1, ""
        Print #1, "   Version 0.02"
        Print #1, "     Rebuilds the forms"
        Print #1, "     Gets most controls and their properties."
        Print #1, ""
        Print #1, "   Intial Release version 0.01"
        Print #1, ""
        Print #1, "2. Features"
        Print #1, "     Decompiling the pcode/native vb6/vb5 exe's"
        Print #1, "     Form Generation, P-Code code view"
        Print #1, "     Resource extraction wmf, ico, cur, gif, bmp, jpg, dib"
        Print #1, "     Form Editor"
        Print #1, "     P-Code Procedure Decompile View"
        Print #1, "     Shows offsets for controls"
        Print #1, "     SubMain Disassembly"
        Print #1, "     Memory Map of the exe file, so you can see what's going on."
        Print #1, "     Advanced decompiling using COM instead of hard coding property opcodes."
        Print #1, ""
        Print #1, "3. Questions?"
        Print #1, "   Q. What about Native Code Decompiling?"
        Print #1, "   A. It is in the works. I need to get a better understanding of how VBDE works before"
        Print #1, "      I begin to work on Native Code."
        Print #1, "      A site that is working on native vb decompiler is"
        Print #1, "      http://www.Decompiler.org"
        Print #1, "   Q. What the heck are the P-Code Tokens?"
        Print #1, "   A. P-Code tokens is the last step before turning the P-Code into readable VB Code."
        Print #1, "      All you have to do now is link the imports of the exe with the functions in P-Code."
        Print #1, "   Q. Why does it not show all the controls on my forms?"
        Print #1, "   A. Usally because its a property that is not detected by COM using vb6.olb."
        Print #1, "   Q. Why doesn't it get my procedure names for Modules?"
        Print #1, "   A. VB only saves procedures names for Form's and Classes."
        Print #1, "   Q. Why is there a ComFix file?"
        Print #1, "   A. Since Visual Basic does not support all the data types that IDL does it is needed."
        Print #1, "      Basicly it fixes when COM returns an integer when it should really be a VB byte."
        Print #1, "   Q. How does this decompiler work?"
        Print #1, "   A. First it gets all the main vb strutures from the exe."
        Print #1, "      Next it gets all the controls properties via COM using vb6.olb"
        Print #1, "      I am still looking for a static pointer for the table inside msvbvm60.dll to use instead."
        Print #1, "   Q. What files does this decompiler require?"
        Print #1, "   A. It requires the following files:"
        Print #1, "      TLBINF32.dll"
        Print #1, "      comdlg32.OCX"
        Print #1, "      RICHTX32.OCX"
        Print #1, "      MSCOMCTL.OCX"
        Print #1, "      TABCTL32.OCX"
        Print #1, "      MSFLXGRD.OCX"
        Print #1, "      Msvbvm60.dll"
        Print #1, "      And VB6.olb version 6.0.9"
        Print #1, "      All of the above files need to be registered!"
        Print #1, "   Q. Where can I learn more about Visual Basic 5/6 Decompiling?"
        Print #1, "   A. Head over to http://www.vb-decompiler.com/  tons of information on vb decompiling."
        Print #1, ""
        Print #1, "4. Bugs"
        Print #1, "     I know about most of them..."
        Print #1, "     MDI Forms and External Controls."
        Print #1, "     Some properties aren't handled yet dataformat, and some others"
        Print #1, "     P-Code decoding may hang use the disable P-Code option"
        Print #1, "     Overflow error is caused by a property that isn't detected yet..."
        Print #1, "     Currently it does not generate user control and property pages"
        Print #1, ""
        Print #1, "5. Contact"
        Print #1, "     Email=gmdecompiler@yahoo.com"
        Print #1, "     Aim=vbgamer45"
        Print #1, ""
        Print #1, "6. Credits"
        Print #1, "     I would like to thank the following people for helping me with this project."
        Print #1, "     Sarge, Mr. Unleaded, Moogman, _aLfa_, ionescu007, Warning and many others."
        
    Close #1
    
End Sub
Public Function sHexStringFromString(ByVal inp As String, Optional Spacing As Boolean = True) As String
Dim hc As String
Dim hs As String
Dim c As Long
While Len(inp)
    
    hc = Hex(Asc(Mid(inp, 1, 1)))
    inp = Mid(inp, 2)
    If Len(hc) = 1 Then hc = "0" & hc
    hs = hs & hc
    c = c + 1
    If Spacing Then
        If c Mod 4 = 0 Then
            hs = hs & "  "
        ElseIf c Mod 2 = 0 Then
            hs = hs & " "
        End If
        
    End If
Wend
sHexStringFromString = hs
End Function
Public Function PadHex(ByVal sHex As String, Optional Pad As Integer = 8) As String
'*****************************
'Purpose: To add extra zero's to a hexadecimal string
'*****************************
    If Len(sHex) > Pad Then
        PadHex = sHex
    Else
        PadHex = String(Pad - Len(sHex), 48) & sHex
    End If
End Function

Public Function AddChar(Val As String, TheLen As Long, Optional Char As String = "0") As String    'Permet d'ajouter un charactère à une chaine de charactère pour obtenir une certaine longueur.
    AddChar = Right(String(TheLen, Char) & Val, TheLen)
End Function
Public Function ExtString(DataStr As String) As String
    ExtString = Left(DataStr, lstrlen(DataStr))
End Function
Public Function GetUntilNull(FileNum As Variant) As String
    '*****************************
    'Purpose to get a null termintated string
    '*****************************
    Dim aList() As Byte
    Dim k As Byte
    k = 255
    ReDim aList(0)
    Do Until k = 0
        Get FileNum, , k
        ReDim Preserve aList(UBound(aList) + 1)
        aList(UBound(aList)) = k
        'MsgBox k
    Loop
    Dim i As Integer
    Dim Final As String
    For i = 1 To UBound(aList) - 1
        Final = Final & Chr(aList(i))
      
    Next i
    
    GetUntilNull = Final
End Function
Public Function GetUnicodeString(FileNum As Variant, length As Integer) As String
    '*****************************
    'Purpose to get a unicode string
    '*****************************
    Dim aList() As Byte

    ReDim aList((length * 2))
    Get FileNum, , aList

    Dim i As Integer
    Dim Final As String
    For i = 1 To UBound(aList) - 1
        If aList(i) <> 0 Then
            Final = Final & Chr(aList(i))
        End If
    Next i
    
    GetUnicodeString = Final
End Function
Public Function FileInfo(Optional ByVal PathWithFilename As String) As FILEPROPERTIE
'*****************************
'Purpose: To return file-properties of given file  (EXE , DLL , OCX)
'*****************************
 
Static BACKUP As FILEPROPERTIE   ' backup info for next call without filename
If Len(PathWithFilename) = 0 Then
    FileInfo = BACKUP
    Exit Function
End If

Dim lngBufferlen As Long
Dim lngDummy As Long
Dim lngRc As Long
Dim lngVerPointer As Long
Dim lngHexNumber As Long
Dim bytBuffer() As Byte
Dim bytBuff(255) As Byte
Dim strBuffer As String
Dim strLangCharset As String
Dim strVersionInfo(7) As String
Dim strTemp As String
Dim intTemp As Integer
       
' size
lngBufferlen = GetFileVersionInfoSize(PathWithFilename, lngDummy)
If lngBufferlen > 0 Then
   ReDim bytBuffer(lngBufferlen)
   lngRc = GetFileVersionInfo(PathWithFilename, 0&, lngBufferlen, bytBuffer(0))
   If lngRc <> 0 Then
      lngRc = VerQueryValue(bytBuffer(0), "\VarFileInfo\Translation", _
               lngVerPointer, lngBufferlen)
      If lngRc <> 0 Then
         'lngVerPointer is a pointer to four 4 bytes of Hex number,
         'first two bytes are language id, and last two bytes are code
         'page. However, strLangCharset needs a  string of
         '4 hex digits, the first two characters correspond to the
         'language id and last two the last two character correspond
         'to the code page id.
         MoveMemory bytBuff(0), lngVerPointer, lngBufferlen
         lngHexNumber = bytBuff(2) + bytBuff(3) * &H100 + _
                bytBuff(0) * &H10000 + bytBuff(1) * &H1000000
         strLangCharset = Hex(lngHexNumber)
         'now we change the order of the language id and code page
         'and convert it into a string representation.
         'For example, it may look like 040904E4
         'Or to pull it all apart:
         '04------        = SUBLANG_ENGLISH_USA
         '--09----        = LANG_ENGLISH
         ' ----04E4 = 1252 = Codepage for Windows:Multilingual
         'Do While Len(strLangCharset) < 8
         '    strLangCharset = "0" & strLangCharset
         'Loop
         If Mid(strLangCharset, 2, 2) = LANG_ENGLISH Then
         strLangCharset2 = "English (US)"

         
         End If

         Do While Len(strLangCharset) < 8
             strLangCharset = "0" & strLangCharset
         Loop
         
         ' assign propertienames
         strVersionInfo(0) = "CompanyName"
         strVersionInfo(1) = "FileDescription"
         strVersionInfo(2) = "FileVersion"
         strVersionInfo(3) = "InternalName"
         strVersionInfo(4) = "LegalCopyright"
         strVersionInfo(5) = "OriginalFileName"
         strVersionInfo(6) = "ProductName"
         strVersionInfo(7) = "ProductVersion"
         ' loop and get fileproperties
         For intTemp = 0 To 7
            strBuffer = String$(255, 0)
            strTemp = "\StringFileInfo\" & strLangCharset _
               & "\" & strVersionInfo(intTemp)
            lngRc = VerQueryValue(bytBuffer(0), strTemp, _
                  lngVerPointer, lngBufferlen)
            If lngRc <> 0 Then
               ' get and format data
               lstrcpy strBuffer, lngVerPointer
               strBuffer = Mid$(strBuffer, 1, InStr(strBuffer, Chr(0)) - 1)
               strVersionInfo(intTemp) = strBuffer
             Else
               ' property not found
               strVersionInfo(intTemp) = "?"
            End If
         Next intTemp
      End If
   End If
End If
' assign array to user-defined-type
FileInfo.CompanyName = strVersionInfo(0)
FileInfo.FileDescription = strVersionInfo(1)
FileInfo.FileVersion = strVersionInfo(2)
FileInfo.InternalName = strVersionInfo(3)
FileInfo.LegalCopyright = strVersionInfo(4)
FileInfo.OrigionalFileName = strVersionInfo(5)
FileInfo.ProductName = strVersionInfo(6)
FileInfo.ProductVersion = strVersionInfo(7)
FileInfo.LanguageID = strLangCharset2
BACKUP = FileInfo
End Function
'*****************************
'The following functions are used for COM
'*****************************
Public Function GetBoolean(FileNum As Variant) As Boolean
'*****************************
'Purpose: Get a boolean value from a file offset
'*****************************
        Dim k As Boolean
        Get FileNum, , k
        GetBoolean = k
End Function
Public Function GetByte2(FileNum As Variant) As Byte
'*****************************
'Purpose: Get a byte value from a file offset
'*****************************
        Dim k As Byte
        Get FileNum, , k
        GetByte2 = k
End Function
Public Function GetInteger(FileNum As Variant) As Integer
'*****************************
'Purpose: Get an integer value from a file offset
'*****************************
        Dim k As Integer
        Get FileNum, , k
        
        GetInteger = k
End Function
Public Function GetLong(FileNum As Variant) As Long
'*****************************
'Purpose: Get a long value from a file offset
'*****************************
        Dim k As Long
        Get FileNum, , k
        GetLong = k
End Function
Public Function GetSingle(FileNum As Variant) As Single
'*****************************
'Purpose: Get a single value from a file offset
'*****************************
        Dim k As Single
        Get FileNum, , k
        GetSingle = k
End Function
Public Function GetString(FileNum As Variant) As String
'*****************************
'Purpose: Get VB String(Not Used)
'*****************************
    'Not used...
        Dim k As String
        Seek FileNum, (Loc(FileNum) + 3)
        Get FileNum, , k
  
        GetString = k
End Function
Public Function GetAllString(FileNum As Variant) As String
'*****************************
'Purpose: Get any kind of string Unicode or Ascii
'*****************************
    Dim length As Integer
    Get FileNum, , length
    
    Dim strText As String
    strText = GetUntilNull(FileNum)
    'MsgBox strText
    If Len(strText) < length Then
    'get unicode string
   ' MsgBox "unicode"
        If length < 100 Then
            Seek FileNum, Loc(FileNum) - 2
            strText = GetUnicodeString(FileNum, length)
            Seek FileNum, Loc(FileNum) + 1
        End If
    End If
    GetAllString = strText
End Function

Sub AddText(strText As String)
'*****************************
'Purpose:Adds text to the current form's textbox. And idents it.
'*****************************
    If gIdentSpaces < 0 Then gIdentSpaces = 0

    frmMain.txtFinal(frmMain.txtFinal.UBound).Text = frmMain.txtFinal(frmMain.txtFinal.UBound).Text & Space(gIdentSpaces * 5) & strText & vbCrLf
End Sub
Sub LoadNewFormHolder(FormName As String)
'*****************************
'Purpose:To load a new textbox to hold each form's information
'*****************************
    Dim i As Integer
    For i = 0 To frmMain.txtFinal.UBound
        If frmMain.txtFinal(i).Tag = "" Then
           ' frmMain.txtFinal(i).Tag = FormName
            'frmMain.txtFinal(i).Text = ""
           ' Exit Sub
        End If
    Next
    
    i = frmMain.txtFinal.UBound + 1
    Load frmMain.txtFinal(i)
    With frmMain.txtFinal(i)
        .Tag = FormName
    
    End With
End Sub

Sub LoadCOMFIX()
'*****************************
'Load the COM Hacks
'Com Hack File Format
'Objectname,PropertyName,NewDataType
'Notes on NewDataType: Can be either Byte Boolean Integer Long Single String
'One more thing to remember all these Properties are case sensetive
'*****************************
    ReDim gComFix(0)
    Open App.path & "\ComFix.txt" For Input As #1
    Dim data As String
    Dim Temp
    
    Do While Not EOF(1)
        Line Input #1, data
        
        Temp = Split(data, ",")
        gComFix(UBound(gComFix)).ObjectName = Temp(0)
        gComFix(UBound(gComFix)).PropertyName = Temp(1)
        gComFix(UBound(gComFix)).NewType = Temp(2)
        ReDim Preserve gComFix(UBound(gComFix) + 1)
    Loop
    Close #1
    ReDim Preserve gComFix(UBound(gComFix) - 1)

    
End Sub

Function ReturnGuid(FileNum As Variant) As String
'*****************************
'Gets a guid from a file, then corrects it into a real guid
'*****************************
Dim bArray(15) As Byte
Dim strArray(15) As String
    Get FileNum, , bArray
    Dim i As Integer
    For i = 0 To 15
        If i = 0 Then
        strArray(0) = Hex(bArray(0) - 2)
        Else
            strArray(i) = Hex(bArray(i))
        End If
        If Len(strArray(i)) = 1 Then
            strArray(i) = ("0" & strArray(i))
        End If
        
    Next
    
   
    Dim strFinal As String
   ' strFinal = "{" & Hex(bArray(3)) & Hex(bArray(2)) & Hex(bArray(1)) & Hex(bArray(0) - 2)
   ' strFinal = strFinal & "-" & Hex(bArray(5)) & Hex(bArray(4))
   ' strFinal = strFinal & "-" & Hex(bArray(7)) & Hex(bArray(6))
   ' strFinal = strFinal & "-" & Hex(bArray(8)) & Hex(bArray(9))
   ' strFinal = strFinal & "-" & Hex(bArray(10)) & Hex(bArray(11)) & Hex(bArray(12)) & Hex(bArray(13)) & Hex(bArray(14)) & Hex(bArray(15)) & "}"
   strFinal = "{" & strArray(3) & strArray(2) & strArray(1) & strArray(0)
   strFinal = strFinal & "-" & strArray(5) & strArray(4)
   strFinal = strFinal & "-" & strArray(7) & strArray(6)
   strFinal = strFinal & "-" & strArray(8) & strArray(9)
   strFinal = strFinal & "-" & strArray(10) & strArray(11) & strArray(12) & strArray(13) & strArray(14) & strArray(15) & "}"
    ReturnGuid = strFinal
End Function

Function ReturnGuidByString(strGuid As String) As String
'*****************************
'Purpose: Experimantal
'*****************************
'Gets and generates the Guid
Dim bArray(15) As Byte
Dim strArray(15) As String
  Dim i As Integer
 For i = 1 To Len(strGuid)
    bArray(i - 1) = Asc(Mid(strGuid, i, 1))
 Next
 
   
    For i = 0 To 15
        If i = 0 Then
        strArray(0) = Hex(bArray(0) - 2)
        Else
            strArray(i) = Hex(bArray(i))
        End If
        If Len(strArray(i)) = 1 Then
            strArray(i) = ("0" & strArray(i))
        End If
        
    Next
    
   
    Dim strFinal As String
   strFinal = "{" & strArray(3) & strArray(2) & strArray(1) & strArray(0)
   strFinal = strFinal & "-" & strArray(5) & strArray(4)
   strFinal = strFinal & "-" & strArray(7) & strArray(6)
   strFinal = strFinal & "-" & strArray(8) & strArray(9)
   strFinal = strFinal & "-" & strArray(10) & strArray(11) & strArray(12) & strArray(13) & strArray(14) & strArray(15) & "}"
    ReturnGuidByString = strFinal
End Function


Sub WriteApiList()
'*****************************
'Purpose: To write the Api's
'*****************************
    Dim i As Integer
    frmMain.txtCode.Text = ""
    For i = 0 To UBound(gApiList) - 1
        frmMain.txtCode.Text = frmMain.txtCode.Text & "Declare " & gApiList(i).strFunctionName & " Lib " & Chr(34) & gApiList(i).strLibraryName & Chr(34) & vbCrLf
        
    Next
End Sub
Public Function GetFirstChar(Start As Long, TextToFind As RichTextBox, ListToLike As String) As FIRSTCHAR_INFO
    Dim i As Long, Cursor As Long, TheChar As String, theCursor As Long, SStart As Long, SLength As Long
    SStart = TextToFind.SelStart
    SLength = TextToFind.SelLength
    Cursor = Len(TextToFind.Text)
    For i = 1 To Len(ListToLike)
        theCursor = TextToFind.Find(Mid(ListToLike, i, 1), Start - 1) + 1
        If theCursor < Cursor And theCursor > 0 Then
            Cursor = theCursor
            TheChar = Mid(ListToLike, i, 1)
        End If
    Next i
    TextToFind.SelStart = SStart
    TextToFind.SelLength = SLength
    If Cursor < Start Then
        Cursor = Start
    Else
        GetFirstChar.lCursor = Cursor
    End If
    GetFirstChar.sChar = TheChar
End Function

Public Function GetPart(DataStr As String, DataId As Long, Separator As String) As String
    Dim Pointer As Long
    On Error Resume Next
    For i = 1 To DataId
        Pointer = InStr(Pointer + 1, DataStr, Separator)
    Next i
    GetPart = Mid(DataStr, Pointer + 1, InStr(Pointer + 1, DataStr, Separator) - Pointer - 1)
End Function

Public Function CountParts(DataStr As String, Separator As String) As Long
    Dim Pointer As Long
    Pointer = 1
    While Pointer <> 0
        Pointer = InStr(Pointer + 1, DataStr, Separator)
        CountParts = CountParts + 1
    Wend
    CountParts = CountParts - 1
End Function
Sub AddPropertyToTheList(strPropertyName As String, Value As Variant, VarType As String, offset As Long, HelpString As String)
'*****************************
'Purpose: Used for Form Editor. To add a textbox and label to hold property name and value
'*****************************
    Dim i As Integer
    i = frmMain.txtEditArray.UBound + 1
    
    Load frmMain.txtEditArray(i)
    
    With frmMain.txtEditArray(i)
        .Text = Value
        .Tag = VarType
        .Top = frmMain.txtEditArray(i - 1).Top + 300
        .Left = frmMain.txtEditArray(0).Left
        If VarType = "String" Then
            .MaxLength = Len(Value)
        End If
        .ToolTipText = HelpString
       ' .BackColor = vbRed
        .Visible = True
    End With
    
    Load frmMain.lblArrayEdit(i)
    
    With frmMain.lblArrayEdit(i)
        .Caption = strPropertyName
        .Top = frmMain.lblArrayEdit(i - 1).Top + 300 ' frmMain.lblArrayEdit(i - 1).Top + frmMain.lblArrayEdit(i - 1).Height
        .Left = frmMain.lblArrayEdit(0).Left
        .Tag = offset
        If VarType = "String" Then
        .Tag = (offset - Len(Value))
        End If
        
        .Visible = True
    End With
    'MsgBox strPropertyName
End Sub


Public Function FileExists(path) As Boolean
'*****************************
'Purpose: Checks wether a FileExists or not
'*****************************
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function
