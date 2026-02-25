Attribute VB_Name = "var"
' Public variables
Public Echo As Boolean        ' Echo On/Off flag.
'Public CancelSend As Integer  ' Flag to stop sending a text file.
Public endTimeOut As Long

Public RecordCount As Long
Public record As Variant
Public TransErrorCode As Integer
Public RecordAddress As Variant
Public RecordLength As Variant
Public RecordType As Variant
Public ok_SFile As Boolean
Public MSBAddr As Variant
Public LSBAddr As Variant
Public HexValue As Variant
Public talkBaseAddress As Variant

'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------
Public dwnLoadFlag As Boolean
Public ConfigFileFlag As Boolean
Public NewCfgFileFlag As Boolean
Public ConfigFileName As Variant    ' Configuration File

'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------
Public RstRAM As Integer
Public ROMStart As Integer
Public Settings As String

' [COMDEF]
Public CommPort As String
Public BaudRate As Long

' [STRINGSREF]
Public StrVersion As String        ' HC11 Application Version

' [SRECFILE]
Public progFileName As String      ' S-RECORD FILE To DOWNLOAD


'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------
Public Const Bit0 = 1
Public Const Bit1 = 2
Public Const Bit2 = 4
Public Const Bit3 = 8
Public Const Bit4 = 16
Public Const Bit5 = 32
Public Const Bit6 = 64
Public Const Bit7 = 128

'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------
Public Const READ_MEM = &H1
Public Const WRITE_MEM = &H41
Public Const READ_MCU = &H81
Public Const WRITE_MCU = &HC1
Public Const END_STOP = &HE1

'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------
Public MCU_mem As String
Public Target_Rec As String
Public File_Rec As String

'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------
Public Const DEFAUT_WIN = -1
Public Const NB_MAX_WINDOW = 7      ' Nombre maximum de fenêtres info + nb. tag

Public ReaderInitOk As Boolean
Public firstTime As Boolean
Public flagAutoDetect As Boolean
Public TypeOfTag As Integer
Public Mag_OnOff As Integer
Public dataStream As Variant        ' String from tag
Public strTag_Type As String
Public winTagAct As Integer         ' Active window tag
Public LastWinTagAct As Integer
Public Rd_Auto As Boolean
Public stopAutoRead As Boolean      ' Use in frmAutRd ...
Public win_TT(NB_MAX_WINDOW) As Boolean
Public writeStartAdd As Integer
Public writeEndAdd As Integer
Public LockBit As Integer
Public TagTransactionError As Variant      ' Error while read tag
Public tTag(10) As Integer

Public Const WIN_Normal = 0
Public Const WIN_Minimized = 1
Public Const WIN_Maximized = 2

Public Const win_DEFAULT = 0
Public Const win_INFO = 1
Public Const win_TT_V4050 = 2
Public Const win_TT_H400X = 3
Public Const win_TT_TK5530 = 4
Public Const win_TT_TK5550 = 5
Public Const win_TT_TK5552 = 6
Public Const win_TT_HITAG1 = 7
Public Const win_TT_HITAG2 = 8

'Public Const win_TT_TK5560 = 6


Public Const TT_TK5530 = &H0
Public Const TT_TK5550 = &H1
Public Const TT_TK5560 = &H2
Public Const TT_TK5552 = &H3
Public Const TT_H400X = &H10
Public Const TT_V4050 = &H12
Public Const TT_HITAG1 = &H20
Public Const TT_HITAG2 = &H21

Public Const TT_AUTO_DETECT = &HFF

Public Const MAX_ADD_TK5550 = 7
Public Const MAX_ADD_TK5552 = 7
Public Const MAX_ADD_V4050 = 31

Public Const LenTagDATA = 8         ' nombre de caractères pour un mot de donnée.
Public Const PLeftLogin = 48
Public Const PLeftRead = 940
Public Const PLeftAutoRd = 1832
Public Const PLeftWrite = 2724
Public Const PLeftFormat = 3616
Public Const PLeftChPwd = 4508
Public Const PLeftReset = 5400

Public Const PTop = 36
Public Const PHeight = 372
Public Const PWidth = 876

'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------
Public CommPortSettings As String
Public CommPortSettings_1 As String
Public CommPortSettings_2 As String
Public Const TXBUFFERSIZE = 1024
Public Const RXBUFFERSIZE = 1024

Public flagComDetect As Boolean

Type SciDef
    comID As MSComm
    ptrT As Timer
    rxIndex As Integer
    rxUserIndex As Integer
    rxBufLength As Integer
    rxbuffer(RXBUFFERSIZE) As Integer
End Type
Public MSC As SciDef



'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------
Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'just added 180897
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

