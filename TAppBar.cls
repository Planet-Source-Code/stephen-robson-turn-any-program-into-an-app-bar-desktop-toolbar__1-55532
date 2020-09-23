VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TAppBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************'
'                                                                             '
' APPBAR.CLS                                                                  '
'                                                                             '
' TAppBar Class v1.4 (Class Module)                                           '
' implements Application Desktop Toolbar                                      '
' (based on J.Richter's CAppBar MFC Class)                                    '
'                                                                             '
' Copyright (c) 1997/98 Paolo Giacomuzzi                                      '
' e-mail: paolo.giacomuzzi@usa.net                                            '
' http://www.geocities.com/SiliconValley/9486                                 '
'                                                                             '
'*****************************************************************************'

Option Explicit

Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public Event Moved()

' CONST Section ---------------------------------------------------------------
 
  ' Debugging Mode On/Off
  #Const DEBUG_MODE = False

  ' AppBar's user notification message
  Const WM_USER = &H400
  Const WM_APPBARNOTIFY = WM_USER + 100
  
  ' We need a timer to determine when the AppBar should be re-hidden
  Const AUTO_HIDE_TIMER_ID = 100
  Const SLIDE_DEF_TIMER_INTERVAL = 400 ' milliseconds
  
  ' Registry Root Keys
  Const HKEY_CURRENT_USER = &H80000001
  Const HKEY_LOCAL_MACHINE = &H80000002

  ' Defaults
  Const AB_DEF_SIZE_INC = 1
  Const AB_DEF_DOCK_SIZE = 32
  Const AB_DEF_ROOT_KEY = HKEY_CURRENT_USER
  Const AB_DEF_KEY_NAME = "Software\AppBar\1.4\VB"

  ' You can send to the Windows shell one of the following messages:
  ' Message             Description
  ' --------------      --------------------------------------------------
  ' ABM_NEW             Register a new AppBar to the system
  ' ABM_REMOVE          Remove a previously created AppBar from the system
  ' ABM_QUERYPOS        Query the AppBar position
  ' ABM_SETPOS          Set the AppBar position
  ' ABM_GETSTATE        Get the edge the Appbar is docked to
  ' ABM_GETTASKBARPOS   Get the Explorer Taskbar position
  ' ABM_ACTIVATE        Activate the AppBar
  ' ABM_GETAUTOHIDEBAR  Query if AppBar has Auto-hide behavior
  ' ABM_SETAUTOHIDEBAR  Set the AppBar's Auto-hide behavior

  ' The ABM_message constants are defined in SHELLAPI.H as follows:
  Const ABM_NEW = &H0
  Const ABM_REMOVE = &H1
  Const ABM_QUERYPOS = &H2
  Const ABM_SETPOS = &H3
  Const ABM_GETSTATE = &H4
  Const ABM_GETTASKBARPOS = &H5
  Const ABM_ACTIVATE = &H6
  Const ABM_GETAUTOHIDEBAR = &H7
  Const ABM_SETAUTOHIDEBAR = &H8
  Const ABM_WINDOWPOSCHANGED = &H9

  ' An AppBar can be in one of 6 states shown in the table below:
  ' State          Description
  ' -----------    -----------------------------------------------------
  ' ABE_UNKNOWN    The Appbar is in an unknown state
  '                (usually during construction/destruction)
  ' ABE_FLOAT      The AppBar is floating on the screen
  ' ABE_LEFT       The Appbar is docked on the left   edge of the screen
  ' ABE_TOP        The Appbar is docked on the top    edge of the screen
  ' ABE_RIGHT      The Appbar is docked on the right  edge of the screen
  ' ABE_BOTTOM     The Appbar is docked on the bottom edge of the screen

  ' The ABE_edge state constants are defined in SHELLAPI.H as follows:
  Const ABE_LEFT = 0
  Const ABE_TOP = 1
  Const ABE_RIGHT = 2
  Const ABE_BOTTOM = 3

  ' The ABE_UNKNOWN and ABE_FLOAT constants are defined here as follows:
  Const ABE_UNKNOWN = 4
  Const ABE_FLOAT = 5
  
  ' An AppBar can have several behavior flags as shown below:
  ' Flag                        Description
  ' --------------------------- -----------------------------------
  ' ABF_ALLOWLEFT               Allow dock on left   of screen
  ' ABF_ALLOWRIGHT              Allow dock on right  of screen
  ' ABF_ALLOWTOP                Allow dock on top    of screen
  ' ABF_ALLOWBOTTOM             Allow dock on bottom of screen
  ' ABF_ALLOWFLOAT              Allow float in the middle of screen
  
  ' The ABF_* constants are defined here as follows:
  Const ABF_ALLOWLEFT = 1
  Const ABF_ALLOWRIGHT = 2
  Const ABF_ALLOWTOP = 4
  Const ABF_ALLOWBOTTOM = 8
  Const ABF_ALLOWFLOAT = 16

  ' The ABN_* constants are defined here as follows:
  Const ABN_FULLSCREENAPP = &H2
  Const ABN_POSCHANGED = &H1
  Const ABN_WINDOWARRANGE = &H3

  ' DeleteMenu Selectors
  Const SC_RESTORE = &HF120
  Const SC_MINIMIZE = &HF020
  Const SC_MAXIMIZE = &HF030
  Const MF_BYCOMMAND = &H0&
  
  ' GetKeyState and GetAsyncKeyState Selectors
  Const VK_LBUTTON = &H1
  Const VK_RBUTTON = &H2
  Const VK_CONTROL = &H11

  ' GetSystemMetrics Selectors
  Const SM_CXSCREEN = 0
  Const SM_CYSCREEN = 1
  Const SM_CXVSCROLL = 2
  Const SM_CYHSCROLL = 3
  Const SM_CXBORDER = 5
  Const SM_CYBORDER = 6
  Const SM_SWAPBUTTON = 23
  Const SM_CXDOUBLECLK = 36
  Const SM_CYDOUBLECLK = 37

  ' MessageBox Selectors
  Const MB_OK = &H0&
  Const MB_ICONINFORMATION = &H40&
  
  ' ModifyStyle Selectors
  Const GWL_STYLE = (-16)
  Const GWL_EXSTYLE = (-20)
  Const WS_CAPTION = &HC00000
  Const WS_SYSMENU = &H80000
  Const WS_EX_APPWINDOW = &H40000
  
  ' Registry Selectors
  Const REG_OPTION_NON_VOLATILE = 0
  Const REG_BINARY = 3
  Const STANDARD_RIGHTS_ALL = &H1F0000
  Const SYNCHRONIZE = &H100000
  Const KEY_QUERY_VALUE = &H1
  Const KEY_SET_VALUE = &H2
  Const KEY_CREATE_SUB_KEY = &H4
  Const KEY_ENUMERATE_SUB_KEYS = &H8
  Const KEY_NOTIFY = &H10
  Const KEY_CREATE_LINK = &H20
  Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or _
                           KEY_QUERY_VALUE Or _
                           KEY_SET_VALUE Or _
                           KEY_CREATE_SUB_KEY Or _
                           KEY_ENUMERATE_SUB_KEYS Or _
                           KEY_NOTIFY Or _
                           KEY_CREATE_LINK) And (Not SYNCHRONIZE))
  
  ' SetWindowPos Selectors
  Const SWP_NOSIZE = &H1
  Const SWP_NOMOVE = &H2
  Const SWP_NOZORDER = &H4
  Const SWP_NOACTIVATE = &H10
  Const SWP_DRAWFRAME = &H20
  
  Const HWND_NOTOPMOST = -2
  Const HWND_TOPMOST = -1
  Const HWND_BOTTOM = 1

  ' ShowWindow Selectors
  Const SW_HIDE = 0
  Const SW_SHOW = 5
  
  ' SystemParametersInfo Selectors
  Const SPI_GETDRAGFULLWINDOWS = 38
  Const SPI_GETWORKAREA = 48

  ' WM_ACTIVATE Selectors
  Const WA_INACTIVE = 0

  ' WM_NCHITTEST Selectors
  Const HTCLIENT = 1
  Const HTCAPTION = 2
  Const HTLEFT = 10
  Const HTRIGHT = 11
  Const HTTOP = 12
  Const HTTOPLEFT = 13
  Const HTTOPRIGHT = 14
  Const HTBOTTOM = 15
  Const HTBOTTOMLEFT = 16
  Const HTBOTTOMRIGHT = 17
  Const HTBORDER = 18
  Const HTSIZEFIRST = HTLEFT
  Const HTSIZELAST = HTBOTTOMRIGHT
  
  ' WM_SIZING Selectors
  Const WMSZ_LEFT = 1
  Const WMSZ_RIGHT = 2
  Const WMSZ_TOP = 3
  Const WMSZ_TOPLEFT = 4
  Const WMSZ_TOPRIGHT = 5
  Const WMSZ_BOTTOM = 6
  Const WMSZ_BOTTOMLEFT = 7
  Const WMSZ_BOTTOMRIGHT = 8


' TYPE Section ----------------------------------------------------------------

  ' The following enumerated type defines the constants in the ABM_* table
  Enum TAppBarMessage
    abmNew = ABM_NEW
    abmRemove = ABM_REMOVE
    abmQueryPos = ABM_QUERYPOS
    abmSetPos = ABM_SETPOS
    abmGetState = ABM_GETSTATE
    abmGetTaskBarPos = ABM_GETTASKBARPOS
    abmActivate = ABM_ACTIVATE
    abmGetAutoHideBar = ABM_GETAUTOHIDEBAR
    abmSetAutoHideBar = ABM_SETAUTOHIDEBAR
    abmWindowPosChanged = ABM_WINDOWPOSCHANGED
  End Enum

  ' The following enumerated type defines the constants in the ABE_* table
  ' (Values are mutually exclusive)
  Enum TAppBarEdge
    abeLeft = ABE_LEFT
    abeTop = ABE_TOP
    abeRight = ABE_RIGHT
    abeBottom = ABE_BOTTOM
    abeUnknown = ABE_UNKNOWN
    abeFloat = ABE_FLOAT
  End Enum

  ' The following enumerated type defines the constants in the ABF_* table
  ' (Values can be OR'ed)
  Enum TAppBarFlags
    abfAllowLeft = ABF_ALLOWLEFT
    abfAllowTop = ABF_ALLOWTOP
    abfAllowRight = ABF_ALLOWRIGHT
    abfAllowBottom = ABF_ALLOWBOTTOM
    abfAllowFloat = ABF_ALLOWFLOAT
  End Enum
  
  ' The following enumerated type defines the AppBar behavior in the Taskbar
  Enum TAppBarTaskEntry
    abtShow
    abtHide
    abtFloatDependent
  End Enum

  ' Rectangle
  Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type
  
  ' Size
  Private Type Size
    cx As Long
    cy As Long
  End Type

  ' The record below contains all of the AppBar settings that
  ' can be saved/loaded in/from the Registry
  Private Type TAppBarSettings
    cbSize         As Long         ' Size of this structure
    abEdge         As TAppBarEdge  ' ABE_UNKNOWN, ABE_FLOAT, or ABE_edge
    abFlags        As TAppBarFlags ' ABF_* flags
    bAutoHide      As Boolean      ' Should AppBar be auto-hidden when docked?
    bAlwaysOnTop   As Boolean      ' Should AppBar always be on top?
    bSlideEffect   As Boolean      ' Should AppBar slide?
    nTimerInterval As Long         ' Slide Timer Interval (determines speed)
    szSizeInc      As Size         ' Discrete width/height size increments
    szDockSize     As Size         ' Width/Height for docked bar
    rcFloat        As Rect         ' Floating rectangle in screen coordinates
    nMinWidth      As Long         ' Min allowed width
    nMinHeight     As Long         ' Min allowed height
    nMaxWidth      As Long         ' Max allowed width
    nMaxHeight     As Long         ' Max allowed height
    szMinDockSize  As Size         ' Min Width/Height when docked
    szMaxDockSize  As Size         ' Max Width/Height when docked
    abTaskEntry    As TAppBarTaskEntry ' AppBar behavior in the Taskbar
  End Type
    Private c_KeepBorderVisibleWhenHidden As Boolean 'if TRUE then when the AppBar is hidden then border will still be shown
    Private c_AllowAutoShowHide As Boolean 'if True then the DLL can auto hide/show an autohidden AppBar
  ' The record below contains the settings location in the registry
  Private Type TAppBarSettingsLocation
    nRootKey As Long   ' HKEY_CURRENT_USER or HKEY_LOCAL_MACHINE
    KeyName  As String ' Key Name starting from root
  End Type
  
  ' Point
  Private Type POINTAPI
    x As Long
    y As Long
  End Type

  ' SmallPoint
  Private Type POINTS
    x  As Integer
    y  As Integer
  End Type
  
  ' MinMaxInfo
  Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
  End Type

  ' CreateStruct
  Private Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    y As Long
    x As Long
    style As Long
    lpszName As String
    lpszClass As String
    ExStyle As Long
  End Type

  ' AppBarData
  Private Type APPBARDATA
    cbSize As Long
    hwnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As Rect
    lParam As Long
  End Type


' DECL Section ----------------------------------------------------------------

' ClientToScreen
Private Declare Function ClientToScreen _
                Lib "user32" _
                (ByVal hwnd As Long, _
                 ByRef lpPoint As POINTAPI) As Long
                 
' CopyMemory
Private Declare Function CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" _
                (ByVal pDest As Any, _
                 ByVal pSource As Any, _
                 ByVal ByteLen As Long) As Long

' DeleteMenu
Private Declare Function DeleteMenu _
                Lib "user32" _
                (ByVal hMenu As Long, _
                 ByVal nPosition As Long, _
                 ByVal wFlags As Long) As Long

' GetActiveWindow
Private Declare Function GetActiveWindow _
                Lib "user32" () As Long

' GetAsyncKeyState
Private Declare Function GetAsyncKeyState _
                Lib "user32" _
                (ByVal vKey As Long) As Integer
                
' GetClientRect
Private Declare Function GetClientRect _
                Lib "user32" _
                (ByVal hwnd As Long, _
                 ByRef lpRect As Rect) As Long

' GetKeyState
Private Declare Function GetKeyState _
                Lib "user32" _
                (ByVal nVirtKey As Long) As Integer

' GetMessagePos
Private Declare Function GetMessagePos _
                Lib "user32" () As Long

' GetSystemMenu
Private Declare Function GetSystemMenu _
                Lib "user32" _
                (ByVal hwnd As Long, _
                 ByVal bRevert As Long) As Long

' GetSystemMetrics
Private Declare Function GetSystemMetrics _
                Lib "user32" _
                (ByVal nIndex As Long) As Long

' GetTickCount
Private Declare Function GetTickCount _
                Lib "kernel32" () As Long

' GetWindowLong
Private Declare Function GetWindowLong _
                Lib "user32" Alias "GetWindowLongA" _
                (ByVal hwnd As Long, _
                 ByVal nIndex As Long) As Long

' GetWindowRect
Private Declare Function GetWindowRect _
                Lib "user32" _
                (ByVal hwnd As Long, _
                 ByRef lpRect As Rect) As Long

' InflateRect
Private Declare Function InflateRect _
                Lib "user32" _
                (ByRef lpRect As Rect, _
                 ByVal x As Long, _
                 ByVal y As Long) As Long

' KillTimer
Private Declare Function KillTimer _
                Lib "user32" _
                (ByVal hwnd As Long, _
                 ByVal nIDEvent As Long) As Long

' MessageBox
Private Declare Function MessageBox _
                Lib "user32" _
                Alias "MessageBoxA" _
                (ByVal hwnd As Long, _
                 ByVal lpText As String, _
                 ByVal lpCaption As String, _
                 ByVal wType As Long) As Long

' RegCloseKey
Private Declare Function RegCloseKey _
                Lib "advapi32.dll" _
                (ByVal hKey As Long) As Long

' RegCreateKeyEx
Private Declare Function RegCreateKeyEx _
                Lib "advapi32.dll" _
                Alias "RegCreateKeyExA" _
                (ByVal hKey As Long, _
                 ByVal lpSubKey As String, _
                 ByVal Reserved As Long, _
                 ByVal lpClass As String, _
                 ByVal dwOptions As Long, _
                 ByVal samDesired As Long, _
                 ByVal lpSecurityAttributes As Long, _
                 ByRef phkResult As Long, _
                 ByRef lpdwDisposition As Long) As Long

' RegOpenKeyEx
Private Declare Function RegOpenKeyEx _
                Lib "advapi32.dll" _
                Alias "RegOpenKeyExA" _
                (ByVal hKey As Long, _
                 ByVal lpSubKey As String, _
                 ByVal ulOptions As Long, _
                 ByVal samDesired As Long, _
                 ByRef phkResult As Long) As Long

' RegQueryValueEx
Private Declare Function RegQueryValueEx _
                Lib "advapi32.dll" _
                Alias "RegQueryValueExA" _
                (ByVal hKey As Long, _
                 ByVal lpValueName As String, _
                 ByVal lpReserved As Long, _
                 ByRef lpType As Long, _
                 ByRef lpData As Any, _
                 ByRef lpcbData As Long) As Long
                 
' RegSetValueEx
Private Declare Function RegSetValueEx _
                Lib "advapi32.dll" _
                Alias "RegSetValueExA" _
                (ByVal hKey As Long, _
                 ByVal lpValueName As String, _
                 ByVal Reserved As Long, _
                 ByVal dwType As Long, _
                 ByRef lpData As Any, _
                 ByVal cbData As Long) As Long
                 
' ScreenToClient
Private Declare Function ScreenToClient _
                Lib "user32" _
                (ByVal hwnd As Long, _
                 ByRef lpPoint As POINTAPI) As Long

' SetTimer
Private Declare Function SetTimer _
                Lib "user32" _
                (ByVal hwnd As Long, _
                 ByVal nIDEvent As Long, _
                 ByVal uElapse As Long, _
                 ByVal lpTimerFunc As Long) As Long

' SetWindowLong
Private Declare Function SetWindowLong _
                Lib "user32" Alias "SetWindowLongA" _
                (ByVal hwnd As Long, _
                 ByVal nIndex As Long, _
                 ByVal dwNewLong As Long) As Long

' SetWindowPos
Private Declare Function SetWindowPos _
                Lib "user32" _
                (ByVal hwnd As Long, _
                 ByVal hWndInsertAfter As Long, _
                 ByVal x As Long, _
                 ByVal y As Long, _
                 ByVal cx As Long, _
                 ByVal cy As Long, _
                 ByVal wFlags As Long) As Long

' SHAppBarMessage
Private Declare Function SHAppBarMessage _
                Lib "shell32.dll" _
                (ByVal dwMessage As Long, pData As APPBARDATA) As Long

' ShowWindow
Private Declare Function ShowWindow _
                Lib "user32" _
                (ByVal hwnd As Long, _
                 ByVal nCmdShow As Long) As Long

' SystemParametersInfo
Private Declare Function SystemParametersInfo _
                Lib "user32" _
                Alias "SystemParametersInfoA" _
                (ByVal uAction As Long, _
                 ByVal uParam As Long, _
                 ByVal lpvParam As Any, _
                 ByVal fuWinIni As Long) As Long

' UpdateWindow
Private Declare Function UpdateWindow _
                Lib "user32" _
                (ByVal hwnd As Long) As Long

Private Type WINDOWPLACEMENT
        Length As Long
        flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As Rect
End Type


' CLASS DATA MEMBERS Section --------------------------------------------------


' PRIVATE
  
  ' Internal implementation state variables
  
    ' The AppBar Form
    Private FormHwnd As Long
    
    ' This AppBar's settings info
    Private FABS As TAppBarSettings

    ' We need a member variable which tracks the proposed edge of the
    ' AppBar while the user is moving it, deciding where to position it.
    ' While not moving, this member must contain ABE_UNKNOWN so that
    ' GetEdge returns the current edge contained in FABS.abEdge.
    ' While moving the AppBar, FabEdgeProposedPrev contains the
    ' proposed edge based on the position of the AppBar.  The proposed
    ' edge becomes the new edge when the user stops moving the AppBar.
    Private FabEdgeProposedPrev As TAppBarEdge

    ' We need a member variable which tracks whether a full screen
    ' application window is open
    Private FbFullScreenAppOpen As Boolean

    ' We need a member variable which tracks whether our autohide window
    ' is visible or not
    Private FbAutoHideIsVisible As Boolean
    
    ' We need a member variable to store the settings location in the registry
    Private FabSettingsLocation As TAppBarSettingsLocation


' CLASS METHODS Section -------------------------------------------------------


' PRIVATE


' Class Initialization and Termination

' Constructs an AppBar
Private Sub Class_Initialize()
  
  ' Force the shell to update its list of AppBars and the workarea.
  ' This is a precaution and is very useful when debugging.  If you create
  ' an AppBar and then just terminate the application, the shell still
  ' thinks that the AppBar exists and the user's workarea is smaller than
  ' it should be.  When a new AppBar is created, calling this function
  ' fixes the user's workarea.
  ResetSystemKnowledge

  ' Set default state of AppBar to float with no width & height
  FABS.cbSize = Len(FABS)
  FABS.abEdge = abeFloat
  FABS.abFlags = abfAllowLeft Or _
                 abfAllowTop Or _
                 abfAllowRight Or _
                 abfAllowBottom Or _
                 abfAllowFloat
  FABS.bAutoHide = False
  FABS.bAlwaysOnTop = True
  FABS.bSlideEffect = True
  FABS.nTimerInterval = SLIDE_DEF_TIMER_INTERVAL
  FABS.szSizeInc.cx = AB_DEF_SIZE_INC
  FABS.szSizeInc.cy = AB_DEF_SIZE_INC
  FABS.szDockSize.cx = AB_DEF_DOCK_SIZE
  FABS.szDockSize.cy = AB_DEF_DOCK_SIZE
  FABS.rcFloat.Left = 0
  FABS.rcFloat.Top = 0
  FABS.rcFloat.Right = 0
  FABS.rcFloat.Bottom = 0
  FABS.nMinWidth = 0
  FABS.nMinHeight = 0
  FABS.nMaxWidth = GetSystemMetrics(SM_CXSCREEN)
  FABS.nMaxHeight = GetSystemMetrics(SM_CYSCREEN)
  FABS.szMinDockSize.cx = 0
  FABS.szMinDockSize.cy = 0
  FABS.szMaxDockSize.cx = GetSystemMetrics(SM_CXSCREEN) \ 2
  FABS.szMaxDockSize.cy = GetSystemMetrics(SM_CYSCREEN) \ 2
  FABS.abTaskEntry = abtFloatDependent
  FabEdgeProposedPrev = abeUnknown
  FbFullScreenAppOpen = False
  FbAutoHideIsVisible = False
  Me.KeepBorderVisibleWhenHidden = True
  Me.AllowAutoShowHide = True
  
  ' Set default location of the settings in the registry
  With FabSettingsLocation
    RootKey = AB_DEF_ROOT_KEY
    KeyName = AB_DEF_KEY_NAME
  End With

End Sub

' Destroys a previously created AppBar
Private Sub Class_Terminate()
  
  ResetSystemKnowledge

End Sub


' Internal implementation functions

' AppBarMessage(1,2,3,4) encapsulate the shell's SHAppBarMessage function
Private Function AppBarMessage(ByVal abMessage As TAppBarMessage, _
                               ByVal abEdge As TAppBarEdge, _
                               ByVal lParam As Long, _
                               ByVal bRect As Boolean, _
                               ByRef rc As Rect) As Long
  Dim abd As APPBARDATA
  ' Initialize an APPBARDATA structure
  abd.cbSize = Len(abd)
  abd.hwnd = FormHwnd
  abd.uCallbackMessage = WM_APPBARNOTIFY
  abd.uEdge = abEdge
  If bRect Then
    abd.rc = rc
  Else
    abd.rc.Left = 0
    abd.rc.Top = 0
    abd.rc.Right = 0
    abd.rc.Bottom = 0
  End If
  abd.lParam = lParam
  AppBarMessage = SHAppBarMessage(abMessage, abd)

  ' If the caller passed a rectangle, return the updated rectangle
  If bRect Then
    rc = abd.rc
  End If

End Function

Private Function AppBarMessage1(ByVal abMessage As TAppBarMessage) As Long
  
  Dim rc As Rect
  AppBarMessage1 = AppBarMessage(abMessage, abeFloat, 0, False, rc)

End Function

Private Function AppBarMessage2(ByVal abMessage As TAppBarMessage, _
                                ByVal abEdge As TAppBarEdge) As Long
  Dim rc As Rect
  AppBarMessage2 = AppBarMessage(abMessage, abEdge, 0, False, rc)

End Function

Private Function AppBarMessage3(ByVal abMessage As TAppBarMessage, _
                                ByVal abEdge As TAppBarEdge, _
                                ByVal lParam As Long) As Long
  Dim rc As Rect
  AppBarMessage3 = AppBarMessage(abMessage, abEdge, lParam, False, rc)

End Function

Private Function AppBarMessage4(ByVal abMessage As TAppBarMessage, _
                                ByVal abEdge As TAppBarEdge, _
                                ByVal lParam As Long, _
                                ByRef rc As Rect) As Long
  
  AppBarMessage4 = AppBarMessage(abMessage, abEdge, lParam, True, rc)

End Function

' Gets a edge (ABE_FLOAT or ABE_edge) from a point (screen coordinates)
Private Function CalcProposedState(ByRef pt As POINTS) As TAppBarEdge
  
  Dim bForceFloat As Boolean
  
  ' Force the AppBar to float if the user is holding down the Ctrl key
  ' and the AppBar's style allows floating
  bForceFloat = CBool(GetKeyState(VK_CONTROL) And &H8000) And _
                CBool(abfAllowFloat And FABS.abFlags)
  If bForceFloat Then
    CalcProposedState = abeFloat
  Else
    CalcProposedState = GetEdgeFromPoint(FABS.abFlags, pt)
  End If

End Function

Property Let SelfHwnd(val As Long)
    FormHwnd = val
End Property

' Gets a retangle position (screen coordinates) from a proposed state
Private Function GetRect(ByVal abEdgeProposed As TAppBarEdge, _
                         ByRef rcProposed As Rect)
  
  ' This function finds the x, y, cx, cy of the AppBar window
  If abEdgeProposed = abeFloat Then
    ' The AppBar is floating, the proposed rectangle is correct
  Else
    ' The AppBar is docked or auto-hide
    ' Set dimensions to full screen
    With rcProposed
      .Left = 0
      .Top = 0
      .Right = GetSystemMetrics(SM_CXSCREEN)
      .Bottom = GetSystemMetrics(SM_CYSCREEN)
    End With

    ' Subtract off what we want from the full screen dimensions
    If Not AutoHide Then
      ' Ask the shell where we can dock
      AppBarMessage4 abmQueryPos, abEdgeProposed, False, rcProposed
    End If

    Select Case abEdgeProposed
      Case abeLeft
        rcProposed.Right = rcProposed.Left + FABS.szDockSize.cx
      Case abeTop
        rcProposed.Bottom = rcProposed.Top + FABS.szDockSize.cy
      Case abeRight
        rcProposed.Left = rcProposed.Right - FABS.szDockSize.cx
      Case abeBottom
        rcProposed.Top = rcProposed.Bottom - FABS.szDockSize.cy
    End Select

  End If

End Function
                          
' Adjusts the AppBar's location to account for autohide
' Returns TRUE if rectangle was adjusted
Private Function AdjustLocationForAutohide(ByVal bShow As Boolean, _
                                           ByRef rc As Rect) As Boolean
  
  Dim x As Long
  Dim y As Long
  Dim cxVisibleBorder As Long
  Dim cyVisibleBorder As Long
  
  If (Edge = abeUnknown) Or (Edge = abeFloat) Or (Not AutoHide) Then
    ' If we are not docked on an edge OR we are not auto-hidden, there is
    ' nothing for us to do; just return
    AdjustLocationForAutohide = False
    Exit Function
  End If

  ' Showing/hiding doesn't change our size; only our position
  x = 0
  y = 0 ' Assume a position of (0, 0)

  If bShow Then
    ' If we are on the right or bottom, calculate our visible position
    Select Case Edge
      Case abeRight
        x = GetSystemMetrics(SM_CXSCREEN) - (rc.Right - rc.Left)
      Case abeBottom
        y = GetSystemMetrics(SM_CYSCREEN) - (rc.Bottom - rc.Top)
    End Select
  Else
    If Me.KeepBorderVisibleWhenHidden = True Then
        ' Keep a part of the AppBar visible at all times
        cxVisibleBorder = 2 * GetSystemMetrics(SM_CXBORDER)
        cyVisibleBorder = 2 * GetSystemMetrics(SM_CYBORDER)
    End If
    ' Calculate our x or y coordinate so that only the border is visible
    Select Case Edge
      Case abeLeft
        x = -((rc.Right - rc.Left) - cxVisibleBorder)
      Case abeRight
        x = GetSystemMetrics(SM_CXSCREEN) - cxVisibleBorder
      Case abeTop
        y = -((rc.Bottom - rc.Top) - cyVisibleBorder)
      Case abeBottom
        y = GetSystemMetrics(SM_CYSCREEN) - cyVisibleBorder
    End Select
  End If

  With rc
    .Right = x + (.Right - .Left)
    .Bottom = y + (.Bottom - .Top)
    .Left = x
    .Top = y
  End With

  AdjustLocationForAutohide = True

End Function

' If AppBar is Autohide and docked, shows/hides the AppBar
Public Function ShowHiddenAppBar(ByVal bShow As Boolean)
  
  Dim rc As Rect
  
  ' Get our window location in screen coordinates
  GetWindowRect FormHwnd, rc

  ' Assume  that we are visible
  FbAutoHideIsVisible = True

  If AdjustLocationForAutohide(bShow, rc) Then
    ' The rectangle was adjusted, we are an autohide bar
    ' Remember whether we are visible or not
    FbAutoHideIsVisible = bShow

    ' Slide window in from or out to the edge
    SlideWindow rc
  End If

End Function

' When Autohide AppBar is shown/hidden, slides in/out of view
Private Function SlideWindow(ByRef rcEnd As Rect)

  Dim bFullDragOn As Long
  Dim rcStart As Rect
  Dim dwTimeStart As Long
  Dim dwTimeEnd As Long
  Dim dwTime As Long
  Dim x As Long
  Dim y As Long
  Dim w As Long
  Dim h As Long
  
  ' Only slide the window if the user has FullDrag turned on
  SystemParametersInfo SPI_GETDRAGFULLWINDOWS, 0, VarPtr(bFullDragOn), 0

  ' Get the current window position
  GetWindowRect FormHwnd, rcStart
  If (FABS.bSlideEffect And bFullDragOn And _
      ((rcStart.Left <> rcEnd.Left) Or _
       (rcStart.Top <> rcEnd.Top) Or _
       (rcStart.Right <> rcEnd.Right) Or _
       (rcStart.Bottom <> rcEnd.Bottom))) Then

    ' Get our starting and ending time
    dwTimeStart = GetTickCount()
    dwTimeEnd = dwTimeStart + FABS.nTimerInterval
    dwTime = dwTimeStart
    While (dwTime < dwTimeEnd)
      ' While we are still sliding, calculate our new position
      x = rcStart.Left - (rcStart.Left - rcEnd.Left) _
          * (dwTime - dwTimeStart) \ FABS.nTimerInterval

      y = rcStart.Top - (rcStart.Top - rcEnd.Top) _
          * (dwTime - dwTimeStart) \ FABS.nTimerInterval

      w = (rcStart.Right - rcStart.Left) _
          - ((rcStart.Right - rcStart.Left) - (rcEnd.Right - rcEnd.Left)) _
          * (dwTime - dwTimeStart) \ FABS.nTimerInterval

      h = (rcStart.Bottom - rcStart.Top) _
          - ((rcStart.Bottom - rcStart.Top) - (rcEnd.Bottom - rcEnd.Top)) _
          * (dwTime - dwTimeStart) \ FABS.nTimerInterval

      ' Show the window at its changed position
      SetWindowPos FormHwnd, 0, x, y, w, h, _
                   SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_DRAWFRAME
      UpdateWindow FormHwnd
      dwTime = GetTickCount()
    Wend
  End If

  ' Make sure that the window is at its final position
    
    MoveWindow FormHwnd, rcEnd.Left, rcEnd.Top, (rcEnd.Right - rcEnd.Left), (rcEnd.Bottom - rcEnd.Top), True

End Function

' Returns which edge we're autohidden on or ABE_UNKNOWN
Private Function GetAutohideEdge() As TAppBarEdge
  
  If FormHwnd = AppBarMessage2(abmGetAutoHideBar, abeLeft) Then
    GetAutohideEdge = abeLeft
  ElseIf FormHwnd = AppBarMessage2(abmGetAutoHideBar, abeTop) Then
    GetAutohideEdge = abeTop
  ElseIf FormHwnd = AppBarMessage2(abmGetAutoHideBar, abeRight) Then
    GetAutohideEdge = abeRight
  ElseIf FormHwnd = AppBarMessage2(abmGetAutoHideBar, abeBottom) Then
    GetAutohideEdge = abeBottom
  Else
    ' NOTE: If AppBar is docked but not auto-hidden, we return ABE_UNKNOWN
    GetAutohideEdge = abeUnknown
  End If

End Function

' Returns a TSmallPoint that gives the cursor position in screen coords
Private Function GetMessagePosition() As POINTS
  
  Dim pt As POINTS
  Dim dw As Long
  
  dw = GetMessagePos()
  pt.x = CInt(dw And &H7FFF)
  pt.y = CInt((dw And &H7FFF0000) \ &H10000)
  GetMessagePosition = pt

End Function

' Changes the style of a window (translated from AfxModifyStyle)
Private Function ModifyStyle(ByVal hwnd As Long, _
                             ByVal nStyleOffset As Long, _
                             ByVal dwRemove As Long, _
                             ByVal dwAdd As Long, _
                             ByVal nFlags As Long, _
                             ByVal bRefresh As Boolean) As Boolean
  
  Dim dwStyle As Long
  Dim dwNewStyle As Long
  
  dwStyle = GetWindowLong(hwnd, nStyleOffset)
  dwNewStyle = (dwStyle And (Not dwRemove)) Or dwAdd

  If dwStyle = dwNewStyle Then
    ModifyStyle = False
    Exit Function
  End If

  If bRefresh Then
    ShowWindow hwnd, SW_HIDE
  End If
  
  SetWindowLong hwnd, nStyleOffset, dwNewStyle

  If bRefresh Then
    ShowWindow hwnd, SW_SHOW
  End If
  
  If nFlags <> 0 Then
    SetWindowPos hwnd, 0, 0, 0, 0, 0, _
      SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or nFlags
  End If
  
  ModifyStyle = True

End Function

Private Function PtInRect(ByRef rc As Rect, ByRef pt As POINTS) As Boolean
  
  PtInRect = (pt.x >= rc.Left) And (pt.x <= rc.Right) And _
             (pt.y >= rc.Top) And (pt.y <= rc.Bottom)

End Function

Private Function ControlAtPos(ByRef Pos As POINTAPI, _
                              ByVal AllowDisabled As Boolean) As Boolean
'  Dim Control As Control
'  Dim pt As POINTS
'  Dim rc As Rect
'
'  ControlAtPos = False
'  For Each Control In Self.Controls
'    With Control
'      pt.x = Pos.x - .Left \ Screen.TwipsPerPixelX
'      pt.y = Pos.y - .Top \ Screen.TwipsPerPixelY
'      rc.Left = 0
'      rc.Top = 0
'      rc.Right = .Width \ Screen.TwipsPerPixelX
'      rc.Bottom = .Height \ Screen.TwipsPerPixelY
'      If PtInRect(rc, pt) And .Visible And (.Enabled Or AllowDisabled) Then
'        ControlAtPos = True
'        Exit Function
'      End If
'    End With
'  Next Control
  
End Function


' FRIEND


' Property selector functions

' Property Flags : Allowed dockable edges
' Get
Property Get flags() As TAppBarFlags
  flags = FABS.abFlags
End Property

' Let
Property Let flags(ByVal abFlags As TAppBarFlags)
  FABS.abFlags = abFlags
End Property

' Property HorzSizeInc : Horizontal size increment
' Get
Property Get HorzSizeInc() As Long
  HorzSizeInc = FABS.szSizeInc.cx
End Property

' Let
Property Let HorzSizeInc(ByVal nIncrement As Long)
  FABS.szSizeInc.cx = nIncrement
End Property

' Property VertSizeInc : Vertical size increment
' Get
Property Get VertSizeInc() As Long
  VertSizeInc = FABS.szSizeInc.cy
End Property

' Let
Property Let VertSizeInc(ByVal nIncrement As Long)
  FABS.szSizeInc.cy = nIncrement
End Property

' Property Edge : Edge to dock on
' Get : Retrieves the AppBar's edge.  If the AppBar is being positioned, its
' proposed state is returned instead
Property Get Edge() As TAppBarEdge
  If FabEdgeProposedPrev <> abeUnknown Then
    Edge = FabEdgeProposedPrev
  Else
    Edge = FABS.abEdge
  End If
End Property

' Let : Changes the AppBar's edge to ABE_UNKNOWN, ABE_FLOAT or an ABE_edge
Property Let Edge(ByVal abEdge As TAppBarEdge)

  Dim abCurrentEdge As TAppBarEdge
  Dim currentRect As Rect
  Dim rc As Rect
  Dim hwnd As Long
  
  ' If the AppBar is registered as auto-hide, unregister it
  abCurrentEdge = GetAutohideEdge()

  If abCurrentEdge <> abeUnknown Then
    ' Our AppBar is auto-hidden, unregister it
    AppBarMessage3 abmSetAutoHideBar, abCurrentEdge, False
  End If
  
  ' Save the new requested state
  FABS.abEdge = abEdge

  Select Case abEdge

    Case abeUnknown
      ' We are being completely unregistered.
      ' Probably, the AppBar window is being destroyed.
      ' If the AppBar is registered as NOT auto-hide, unregister it
      AppBarMessage1 abmRemove

    Case abeFloat
      ' We are floating and therefore are just a regular window.
      ' Tell the shell that the docked AppBar should be of 0x0 dimensions
      ' so that the workspace is not affected by the AppBar
      currentRect.Left = 0
      currentRect.Top = 0
      currentRect.Right = 0
      currentRect.Bottom = 0
      AppBarMessage4 abmSetPos, abEdge, False, currentRect
      MoveWindow FormHwnd, FABS.rcFloat.Left, FABS.rcFloat.Top, _
                    (FABS.rcFloat.Right - FABS.rcFloat.Left), _
                    (FABS.rcFloat.Bottom - FABS.rcFloat.Top), True

    Case Else
      If AutoHide And (AppBarMessage3(abmSetAutoHideBar, Edge, True) = 0) Then
        ' We couldn't set the AppBar on a new edge, let's dock it instead
        FABS.bAutoHide = False
        ' Call a virtual function to let derived classes know that the AppBar
        ' changed from auto-hide to docked
        OnAppBarForcedToDocked
      End If

      GetRect Edge, rc
      If AutoHide Then
        currentRect.Left = 0
        currentRect.Top = 0
        currentRect.Right = 0
        currentRect.Bottom = 0
        AppBarMessage4 abmSetPos, abeLeft, False, currentRect
      Else
        ' Tell the shell where the AppBar is
        AppBarMessage4 abmSetPos, abEdge, False, rc
      End If

      AdjustLocationForAutohide FbAutoHideIsVisible, rc

      ' Slide window in from or out to the edge
      SlideWindow rc

  End Select

  ' Set the AppBar's z-order appropriately
  hwnd = HWND_NOTOPMOST ' Assume normal Z-Order
  If FABS.bAlwaysOnTop Then
    ' If we are supposed to be always-on-top, put us there
    hwnd = HWND_TOPMOST
    If FbFullScreenAppOpen Then
      ' But, if a full-screen window is opened, put ourself at the bottom
      ' of the z-order so that we don't cover the full-screen window
      hwnd = HWND_BOTTOM
    End If
  End If
  SetWindowPos FormHwnd, _
               hwnd, _
               0, 0, 0, 0, _
               SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE

  ' Make sure that any auto-hide appbars stay on top of us after we move
  ' even though our activation state has not changed
  AppBarMessage1 abmActivate

  ' Tell our derived class that there is a state change
  OnAppBarStateChange False, abEdge

  ' Show or hide the taskbar entry depending on AppBar position
  
  Select Case FABS.abTaskEntry
    Case abtShow
      ModifyStyle FormHwnd, _
                  GWL_EXSTYLE, _
                  0, _
                  WS_EX_APPWINDOW, _
                  0, _
                  True
    Case abtHide
      ModifyStyle FormHwnd, _
                  GWL_EXSTYLE, _
                  WS_EX_APPWINDOW, _
                  0, _
                  0, _
                  True
    Case abtFloatDependent
      Select Case abEdge
        Case abeFloat
          ModifyStyle FormHwnd, _
                      GWL_EXSTYLE, _
                      0, _
                      WS_EX_APPWINDOW, _
                      0, _
                      True
        Case abeLeft, abeTop, abeRight, abeBottom
          ModifyStyle FormHwnd, _
                      GWL_EXSTYLE, _
                      WS_EX_APPWINDOW, _
                      0, _
                      0, _
                      True
      End Select
  End Select

End Property

' Property AutoHide : Auto-hide On/Off
' Get : Returns TRUE if Auto-hide is on, FALSE if Auto-hide is off
Property Get AutoHide() As Boolean
  AutoHide = FABS.bAutoHide
End Property

' Let : Sets the Auto-hide behavior
Property Let AutoHide(ByVal bAutoHide As Boolean)
  FABS.bAutoHide = bAutoHide
End Property
    
' Property AlwaysOnTop : Always On Top On/Off
' Get : Returns TRUE if AppBar is always on topAuto-hide, FALSE otherwise
Property Get AlwaysOnTop() As Boolean
  AlwaysOnTop = FABS.bAlwaysOnTop
End Property

' Let : Sets the AlwaysOnTop behavior
Property Let AlwaysOnTop(ByVal bAlwaysOnTop As Boolean)
  FABS.bAlwaysOnTop = bAlwaysOnTop
End Property

'when True the autohide border will be visible at the edge of the screen
Property Get KeepBorderVisibleWhenHidden() As Boolean
    KeepBorderVisibleWhenHidden = c_KeepBorderVisibleWhenHidden
    End Property
    Property Let KeepBorderVisibleWhenHidden(val As Boolean)
        c_KeepBorderVisibleWhenHidden = val
End Property

'when True the DLL can autohide/show an autohide AppBar
Property Get AllowAutoShowHide() As Boolean
    AllowAutoShowHide = c_AllowAutoShowHide
    End Property
    Property Let AllowAutoShowHide(val As Boolean)
        c_AllowAutoShowHide = val
End Property

' Property SlideEffect : Slide Effect On/Off
' Get
Property Get SlideEffect() As Boolean
  SlideEffect = FABS.bSlideEffect
End Property

' Let
Property Let SlideEffect(ByVal bSlideEffect As Boolean)
  FABS.bSlideEffect = bSlideEffect
End Property

' Property SlideTime
' Get
Property Get SlideTime() As Long
  SlideTime = FABS.nTimerInterval
End Property

' Let
Property Let SlideTime(ByVal nInterval As Long)
  ' Kill the current timer
  KillTimer FormHwnd, AUTO_HIDE_TIMER_ID
  ' Set the new timer interval
  FABS.nTimerInterval = nInterval
  ' Recreate the timer
  SetTimer FormHwnd, AUTO_HIDE_TIMER_ID, FABS.nTimerInterval, 0
End Property

' Property HorzDockSize : Height when docked horizontally
' Get
Property Get HorzDockSize() As Long
  HorzDockSize = FABS.szDockSize.cy
End Property

' Let
Property Let HorzDockSize(ByRef lSize As Long)
  FABS.szDockSize.cy = lSize
End Property

' Property VertDockSize : Width when docked vertically
' Get
Property Get VertDockSize() As Long
  VertDockSize = FABS.szDockSize.cx
End Property

' Let
Property Let VertDockSize(ByRef lSize As Long)
  FABS.szDockSize.cx = lSize
End Property

' AppBar rectangle when floating

' Property FloatLeft
' Get
Property Get FloatLeft() As Long
  FloatLeft = FABS.rcFloat.Left
End Property

' Let
Property Let FloatLeft(ByRef lPos As Long)
  FABS.rcFloat.Left = lPos
End Property

' Property FloatTop
' Get
Property Get FloatTop() As Long
  FloatTop = FABS.rcFloat.Top
End Property

' Let
Property Let FloatTop(ByRef lPos As Long)
  FABS.rcFloat.Top = lPos
End Property

' Property FloatRight
' Get
Property Get FloatRight() As Long
  FloatRight = FABS.rcFloat.Right
End Property

' Let
Property Let FloatRight(ByRef lPos As Long)
  FABS.rcFloat.Right = lPos
End Property

' Property FloatBottom
' Get
Property Get FloatBottom() As Long
  FloatBottom = FABS.rcFloat.Bottom
End Property

' Let
Property Let FloatBottom(ByRef lPos As Long)
  FABS.rcFloat.Bottom = lPos
End Property

' AppBar MinMax dimensions when floating

' Property MinWidth
' Get
Property Get MinWidth() As Long
  MinWidth = FABS.nMinWidth
End Property

' Let
Property Let MinWidth(ByVal nWidth As Long)
  FABS.nMinWidth = nWidth
End Property

' Property MinHeight
' Get
Property Get MinHeight() As Long
  MinHeight = FABS.nMinHeight
End Property

' Let
Property Let MinHeight(ByVal nHeight As Long)
  FABS.nMinHeight = nHeight
End Property

' Property MaxWidth
' Get
Property Get MaxWidth() As Long
  MaxWidth = FABS.nMaxWidth
End Property

' Let
Property Let MaxWidth(ByVal nWidth As Long)
  FABS.nMaxWidth = nWidth
End Property

' Property MinHeight
' Get
Property Get MaxHeight() As Long
  MaxHeight = FABS.nMaxHeight
End Property

' Let
Property Let MaxHeight(ByVal nHeight As Long)
  FABS.nMaxHeight = nHeight
End Property

' Property MinHorzDockSize : Min Height when docked horizontally
' Get
Property Get MinHorzDockSize() As Long
  MinHorzDockSize = FABS.szMinDockSize.cy
End Property

' Let
Property Let MinHorzDockSize(ByVal lSize As Long)
  FABS.szMinDockSize.cy = lSize
End Property

' Property MaxHorzDockSize : Max Height when docked horizontally
' Get
Property Get MaxHorzDockSize() As Long
  MaxHorzDockSize = FABS.szMaxDockSize.cy
End Property

' Let
Property Let MaxHorzDockSize(ByVal lSize As Long)
  FABS.szMaxDockSize.cy = lSize
End Property

' Property MinVertDockSize : Min Width when docked vertically
' Get
Property Get MinVertDockSize() As Long
  MinVertDockSize = FABS.szMinDockSize.cx
End Property

' Let
Property Let MinVertDockSize(ByVal lSize As Long)
  FABS.szMinDockSize.cx = lSize
End Property

' Property MaxVertDockSize : Max Width when docked vertically
' Get
Property Get MaxVertDockSize() As Long
  MaxVertDockSize = FABS.szMaxDockSize.cx
End Property

' Let
Property Let MaxVertDockSize(ByVal lSize As Long)
  FABS.szMaxDockSize.cx = lSize
End Property

' Property TaskEntry : AppBar behavior in the Window Taskbar
' Get
Property Get TaskEntry() As TAppBarTaskEntry
  TaskEntry = FABS.abTaskEntry
End Property

' Let
Property Let TaskEntry(abTaskEntry As TAppBarTaskEntry)
  FABS.abTaskEntry = abTaskEntry
End Property

' Property RootKey : where settings should be loaded/saved in the registry
' Get
Property Get RootKey() As Long
  RootKey = FabSettingsLocation.nRootKey
End Property

' Let
Property Let RootKey(lKey As Long)
  FabSettingsLocation.nRootKey = lKey
End Property

' Property KeyName : where settings should be loaded/saved in the registry
' Get
Property Get KeyName() As String
  KeyName = FabSettingsLocation.KeyName
End Property

' Let
Property Let KeyName(strKey As String)
  FabSettingsLocation.KeyName = strKey
End Property

' Overridable functions

' Called when the AppBar's proposed state changes
Friend Function OnAppBarStateChange(ByVal bProposed As Boolean, _
                                    ByVal abEdgeProposed As TAppBarEdge)

  Dim bFullDragOn As Long
  
  ' Find out if the user has FullDrag turned on
  SystemParametersInfo SPI_GETDRAGFULLWINDOWS, 0, VarPtr(bFullDragOn), 0

  ' If FullDrag is turned on OR the appbar has changed position
  If bFullDragOn Or Not bProposed Then
    If abEdgeProposed = abeFloat Then
      ' Show the window adornments
      ModifyStyle FormHwnd, _
                  GWL_STYLE, _
                  0, _
                  WS_CAPTION Or WS_SYSMENU, _
                  SWP_DRAWFRAME, _
                  False
    Else
      ' Hide the window adornments
      ModifyStyle FormHwnd, _
                  GWL_STYLE, _
                  WS_CAPTION Or WS_SYSMENU, _
                  0, _
                  SWP_DRAWFRAME, _
                  False
    End If
  End If

End Function

' Called if user attempts to dock an Autohide AppBar on
' an edge that already contains an Autohide AppBar
Friend Function OnAppBarForcedToDocked()
  
  ' Display the application name as the message box caption text.
  MessageBox FormHwnd, _
             "There is already an auto hidden window on this edge." + _
             Chr(10) + Chr(13) + _
             "Only one auto hidden window is allowed on each edge.", _
             "", _
             MB_OK + MB_ICONINFORMATION

End Function

' Called when AppBar gets an ABN_FULLSCREENAPP notification
Friend Function OnABNFullScreenApp(ByVal bOpen As Boolean)

  ' This function is called when a FullScreen window is openning or
  ' closing. A FullScreen window is a top-level window that has its caption
  ' above the top of the screen allowing the entire screen to be occupied
  ' by the window's client area.

  ' If the AppBar is a topmost window when a FullScreen window is activated,
  ' we need to change our window to a non-topmost window so that the AppBar
  ' doesn't cover the FullScreen window's client area.

  ' If the FullScreen window is closing, we need to set the AppBar's
  ' Z-Order back to when the user wants it to be.
  FbFullScreenAppOpen = bOpen
  UpdateBar

End Function

' Called when AppBar gets an ABN_POSCHANGED notification
Friend Function OnABNPosChanged()

  ' The TaskBar or another AppBar has changed its size or position
  If (Edge <> abeFloat) And (Not AutoHide) Then
    ' If we're not floating and we're not auto-hidden, we have to
    ' reposition our window
    UpdateBar
      RaiseEvent Moved
  End If

End Function

' Called when AppBar gets an ABN_WINDOWARRANGE notification
Friend Function OnABNWindowArrange(ByVal bBeginning As Boolean)
  
  ' This function intentionally left blank

End Function


' Message handlers

' Called when the AppBar receives a WM_APPBARNOTIFY window message
Friend Function OnAppBarCallbackMsg(ByVal wParam, ByVal lParam) As Long

  Select Case wParam
    Case ABN_FULLSCREENAPP
      OnABNFullScreenApp CBool(lParam)
    Case ABN_POSCHANGED
      OnABNPosChanged
    Case ABN_WINDOWARRANGE
      OnABNWindowArrange CBool(lParam)
  End Select
  
  OnAppBarCallbackMsg = 0

End Function

' Called when the AppBar form is first created
Friend Function OnCreate()

  Dim hMenu As Long
  
  ' Associate a timer with the AppBar.  The timer is used to determine
  ' when a visible, inactive, auto-hide AppBar should be re-hidden
  SetTimer FormHwnd, AUTO_HIDE_TIMER_ID, FABS.nTimerInterval, 0

  ' Save the initial size and position of the floating AppBar
  With FABS.rcFloat
    
    Dim WP As WINDOWPLACEMENT
    WP.Length = Len(WP)
    GetWindowPlacement FormHwnd, WP
    
    .Left = WP.rcNormalPosition.Left
    .Top = WP.rcNormalPosition.Top
    .Right = WP.rcNormalPosition.Right
    .Bottom = WP.rcNormalPosition.Bottom
  End With

  ' Register our AppBar window with the Shell
  AppBarMessage1 abmNew

  ' Update AppBar internal state
  UpdateBar

  ' Remove system menu
  hMenu = GetSystemMenu(FormHwnd, False)
  DeleteMenu hMenu, SC_RESTORE, MF_BYCOMMAND
  DeleteMenu hMenu, SC_MINIMIZE, MF_BYCOMMAND
  DeleteMenu hMenu, SC_MAXIMIZE, MF_BYCOMMAND

End Function

' Called when the AppBar form is about to be destroyed
Friend Function OnDestroy()

  ' Kill the Autohide timer
  KillTimer FormHwnd, AUTO_HIDE_TIMER_ID
  
  ' Unregister our AppBar window with the Shell
  Edge = abeUnknown

End Function

' Called when the AppBar receives a WM_WINDOWPOSCHANGED message
Friend Function OnWindowPosChanged()
  
  ' When our window changes position, tell the Shell so that any
  ' auto-hidden AppBar on our edge stays on top of our window making it
  ' always accessible to the user
  AppBarMessage1 abmWindowPosChanged
End Function

' Called when the AppBar receives a WM_ACTIVATE message
Friend Function OnActivate(ByVal wParam As Long)

  If wParam = WA_INACTIVE Then
    ' Hide the AppBar if we are docked and auto-hidden
    ShowHiddenAppBar False
  End If
  
  ' When our window changes position, tell the Shell so that any
  ' auto-hidden AppBar on our edge stays on top of our window making it
  ' always accessible to the user.
  AppBarMessage1 abmActivate

End Function

' Called every timer tick
Friend Function OnAppBarTimer()

  Dim pt As POINTS
  Dim rc As Rect
  
  If GetActiveWindow <> FormHwnd Then
    ' Possibly hide the AppBar if we are not the active window
    ' Get the position of the mouse and the AppBar's position
    ' Everything must be in screen coordinates
    pt = GetMessagePosition
    GetWindowRect FormHwnd, rc
    ' Add a little margin around the AppBar
    InflateRect rc, _
                2 * GetSystemMetrics(SM_CXDOUBLECLK), _
                2 * GetSystemMetrics(SM_CYDOUBLECLK)
    If Not PtInRect(rc, pt) Then
      ' If the mouse is NOT over the AppBar, hide the AppBar
      If AllowAutoShowHide = True Then
          ShowHiddenAppBar False
      End If
    End If
  End If

End Function

' Called when the AppBar receives a WM_NCMOUSEMOVE message
Friend Function OnNcMouseMove()
  
  ' If we are a docked, auto-hidden AppBar, shown us
  ' when the user moves over our non-client area
    If AllowAutoShowHide = True Then
        ShowHiddenAppBar True
    End If
  
End Function

' Called when the AppBar receives a WM_NCHITTEST message
Friend Function OnNcHitTest(ByVal lParam As Long, ByRef Result As Long) As Long

  Dim u As Long
  Dim bPrimaryMouseBtnDown As Boolean
  Dim rcClient As Rect
  Dim pt As POINTAPI
  Dim vKey As Long
  Dim XPos As Integer
  Dim YPos As Integer
  
  ' Find out what the system thinks is the hit test code
  u = Result
  
  ' Get cursor position in screen coordinates
  XPos = lParam And &H7FFF
  YPos = (lParam And &H7FFF0000) \ &H10000

  ' NOTE: If the user presses the secondary mouse button, pretend that the
  ' user clicked on the client area so that we get WM_CONTEXTMENU messages
  If GetSystemMetrics(SM_SWAPBUTTON) <> 0 Then
    vKey = VK_RBUTTON
  Else
    vKey = VK_LBUTTON
  End If
  bPrimaryMouseBtnDown = CBool(GetAsyncKeyState(vKey) And &H8000)

  pt.x = XPos
  pt.y = YPos
  ScreenToClient FormHwnd, pt
  If (u = HTCLIENT) And bPrimaryMouseBtnDown _
     And Not ControlAtPos(pt, False) Then
    ' User clicked in client area, allow AppBar to move.  We get this
    ' behavior by pretending that the user clicked on the caption area
    u = HTCAPTION
  End If

  ' If the AppBar is floating and the hittest code is a resize code...
  If ((Edge = abeFloat) And _
      (HTSIZEFIRST <= u) And (u <= HTSIZELAST)) Then
    Select Case u
      Case HTLEFT, HTRIGHT
        If FABS.szSizeInc.cx = 0 Then
          u = HTBORDER
        End If
      Case HTTOP, HTBOTTOM
        If FABS.szSizeInc.cy = 0 Then
          u = HTBORDER
        End If
      Case HTTOPLEFT
        If (FABS.szSizeInc.cx = 0) And (FABS.szSizeInc.cy = 0) Then
          u = HTBORDER
        ElseIf (FABS.szSizeInc.cx = 0) And (FABS.szSizeInc.cy <> 0) Then
          u = HTTOP
        ElseIf (FABS.szSizeInc.cx <> 0) And (FABS.szSizeInc.cy = 0) Then
          u = HTLEFT
        End If
      Case HTTOPRIGHT
        If (FABS.szSizeInc.cx = 0) And (FABS.szSizeInc.cy = 0) Then
          u = HTBORDER
        ElseIf (FABS.szSizeInc.cx = 0) And (FABS.szSizeInc.cy <> 0) Then
          u = HTTOP
        ElseIf (FABS.szSizeInc.cx <> 0) And (FABS.szSizeInc.cy = 0) Then
          u = HTRIGHT
        End If
      Case HTBOTTOMLEFT
        If (FABS.szSizeInc.cx = 0) And (FABS.szSizeInc.cy = 0) Then
          u = HTBORDER
        ElseIf (FABS.szSizeInc.cx = 0) And (FABS.szSizeInc.cy <> 0) Then
          u = HTBOTTOM
        ElseIf (FABS.szSizeInc.cx <> 0) And (FABS.szSizeInc.cy = 0) Then
          u = HTLEFT
        End If
      Case HTBOTTOMRIGHT
        If (FABS.szSizeInc.cx = 0) And (FABS.szSizeInc.cy = 0) Then
          u = HTBORDER
        ElseIf (FABS.szSizeInc.cx = 0) And (FABS.szSizeInc.cy <> 0) Then
          u = HTBOTTOM
        ElseIf (FABS.szSizeInc.cx <> 0) And (FABS.szSizeInc.cy = 0) Then
          u = HTRIGHT
        End If
    End Select
  End If

  ' When the AppBar is docked, the user can resize only one edge.
  ' This next section determines which edge the user can resize.
  ' To allow resizing, the AppBar window must have the WS_THICKFRAME style

  ' If the AppBar is docked and the hittest code is a resize code...
  If ((Edge <> abeFloat) And (Edge <> abeUnknown) And _
      (HTSIZEFIRST <= u) And (u <= HTSIZELAST)) Then

    If (IsEdgeLeftOrRight(Edge) And (FABS.szSizeInc.cx = 0)) Or _
       (Not IsEdgeLeftOrRight(Edge) And (FABS.szSizeInc.cy = 0)) Then
      ' If the width/height size increment is zero, then resizing is NOT
      ' allowed for the edge that the AppBar is docked on
      u = HTBORDER ' Pretend that the mouse is not on a resize border
    Else
      ' Resizing IS allowed for the edge that the AppBar is docked on
      ' Get the location of the appbar's client area in screen coordinates
      GetClientRect FormHwnd, rcClient
      pt.x = rcClient.Left
      pt.y = rcClient.Top
      ClientToScreen FormHwnd, pt
      rcClient.Left = pt.x
      rcClient.Top = pt.y
      pt.x = rcClient.Right
      pt.y = rcClient.Bottom
      ClientToScreen FormHwnd, pt
      rcClient.Right = pt.x
      rcClient.Bottom = pt.y

      u = HTBORDER ' Assume that we can't resize
      Select Case Edge
        Case abeLeft
          If XPos > rcClient.Right Then
            u = HTRIGHT
          End If
        Case abeTop
          If YPos > rcClient.Bottom Then
            u = HTBOTTOM
          End If
        Case abeRight
          If XPos < rcClient.Left Then
            u = HTLEFT
          End If
        Case abeBottom
          If YPos < rcClient.Top Then
            u = HTTOP
          End If
      End Select
    End If
  End If

  ' Return the hittest code
  Result = u

End Function

' Called when the AppBar receives a WM_ENTERSIZEMOVE message
Friend Function OnEnterSizeMove() As Long

  ' The user started moving/resizing the AppBar, save its current state
  FabEdgeProposedPrev = Edge
  
  ' Trap default processing
  OnEnterSizeMove = 0

End Function

' Called when the AppBar receives a WM_EXITSIZEMOVE message
Friend Function OnExitSizeMove() As Long
  
  Dim abEdgeProposedPrev As TAppBarEdge
  Dim rc As Rect
  Dim rcWorkArea As Rect
  Dim w As Long
  Dim h As Long

  ' The user stopped moving/resizing the AppBar, set the new state
  ' Save the new proposed state of the AppBar
  abEdgeProposedPrev = FabEdgeProposedPrev

  ' Set the proposed state back to unknown.  This causes GetState
  ' to return the current state rather than the proposed state
  FabEdgeProposedPrev = abeUnknown

  ' Get the location of the window in screen coordinates
  GetWindowRect FormHwnd, rc

  ' If the AppBar's state has changed...
  If Edge = abEdgeProposedPrev Then
    Select Case Edge
      Case abeLeft, abeRight
        ' Save the new width of the docked AppBar
        FABS.szDockSize.cx = rc.Right - rc.Left
      Case abeTop, abeBottom
        ' Save the new height of the docked AppBar
        FABS.szDockSize.cy = rc.Bottom - rc.Top
    End Select
  End If

  ' Always save the new position of the floating AppBar
  If abEdgeProposedPrev = abeFloat Then
    ' If AppBar was floating and keeps floating...
    If Edge = abeFloat Then
      FABS.rcFloat = rc
    ' If AppBar was docked and is going to float...
    Else
      ' Propose width and height depending on the current window position
      w = rc.Right - rc.Left
      h = rc.Bottom - rc.Top
      ' Adjust width and height
      SystemParametersInfo SPI_GETWORKAREA, 0, VarPtr(rcWorkArea), 0
      If (w >= (rcWorkArea.Right - rcWorkArea.Left)) Or _
         (h >= (rcWorkArea.Bottom - rcWorkArea.Top)) Then
        w = FABS.rcFloat.Right - FABS.rcFloat.Left
        h = FABS.rcFloat.Bottom - FABS.rcFloat.Top
      End If
      ' Save new floating position
      FABS.rcFloat.Left = rc.Left
      FABS.rcFloat.Top = rc.Top
      FABS.rcFloat.Right = rc.Left + w
      FABS.rcFloat.Bottom = rc.Top + h
    End If
  End If

  ' After setting the dimensions, set the AppBar to the proposed state
  Edge = abEdgeProposedPrev
  
  ' Trap default processing
  OnExitSizeMove = 0

End Function

' Called when the AppBar receives a WM_MOVING message
Friend Function OnMoving(ByVal lParam As Long) As Long
  
  Dim rc As Rect
  Dim pt As POINTS
  Dim abEdgeProposed As TAppBarEdge
  Dim w As Long
  Dim h As Long
  
  ' We control the moving of the AppBar.  For example, if the mouse moves
  ' close to an edge, we want to dock the AppBar
  ' The lParam contains the window's position proposed by the system
  CopyMemory VarPtr(rc), lParam, Len(rc)

  ' Get the location of the mouse cursor
  pt = GetMessagePosition

  ' Where should the AppBar be based on the mouse position?
  abEdgeProposed = CalcProposedState(pt)

  If (FabEdgeProposedPrev <> abeFloat) And (abEdgeProposed = abeFloat) Then
    ' While moving, the user took us from a docked/autohidden state to
    ' the float state.  We have to calculate a rectangle location so that
    ' the mouse cursor stays inside the window.
    rc = FABS.rcFloat
    w = rc.Right - rc.Left
    h = rc.Bottom - rc.Top
    With rc
      .Left = pt.x - w \ 2
      .Top = pt.y
      .Right = pt.x - w \ 2 + w
      .Bottom = pt.y + h
    End With
  End If

  ' Remember the most-recently proposed state
  FabEdgeProposedPrev = abEdgeProposed

  ' Tell the system where to move the window based on the proposed state
  GetRect abEdgeProposed, rc

  ' Tell our derived class that there is a proposed state change
  OnAppBarStateChange True, abEdgeProposed
    
  ' Tell the system the new rectangle
  CopyMemory lParam, VarPtr(rc), Len(rc)
  
  ' Trap default processing
  OnMoving = 0
End Function

' Called when the AppBar receives a WM_SIZING message
Friend Function OnSizing(ByVal wParam As Long, ByVal lParam As Long)
  
  Dim rc As Rect
  Dim rcBorder As Rect
  Dim nWidthNew As Long
  Dim nHeightNew As Long
  
  ' We control the sizing of the AppBar.  For example, if the user re-sizes
  ' an edge, we want to change the size in discrete increments.
  ' The lParam contains the window's position proposed by the system
  CopyMemory VarPtr(rc), lParam, Len(rc)

  ' Get the minimum allowed size of the window depending on current edge.
  ' This is the width/height of the window that must always be present
  rcBorder.Left = 0
  rcBorder.Top = 0
  With FABS
    Select Case .abEdge
      Case abeFloat
        rcBorder.Right = .nMinWidth
        rcBorder.Bottom = .nMinHeight
      Case Else
        rcBorder.Right = .szMinDockSize.cx
        rcBorder.Bottom = .szMinDockSize.cy
    End Select
  End With
  
  ' We force the window to resize in discrete units set by the FABS.szSizeInc
  ' member.  From the new, proposed window dimensions passed to us, round
  ' the width/height to the nearest discrete unit
  If FABS.szSizeInc.cx <> 0 Then
    nWidthNew = (((rc.Right - rc.Left) - (rcBorder.Right - rcBorder.Left) _
                 + (FABS.szSizeInc.cx \ 2)) \ FABS.szSizeInc.cx) _
                 * FABS.szSizeInc.cx + (rcBorder.Right - rcBorder.Left)
  Else
    nWidthNew = rc.Right - rc.Left
  End If

  If FABS.szSizeInc.cy <> 0 Then
    nHeightNew = (((rc.Bottom - rc.Top) - (rcBorder.Bottom - rcBorder.Top) _
                  + (FABS.szSizeInc.cy \ 2)) \ FABS.szSizeInc.cy) _
                  * FABS.szSizeInc.cy + (rcBorder.Bottom - rcBorder.Top)
  Else
    nHeightNew = rc.Bottom - rc.Top
  End If

  ' Adjust the rectangle's dimensions
  Select Case wParam
    
    Case WMSZ_LEFT
      rc.Left = rc.Right - nWidthNew

    Case WMSZ_TOP
      rc.Top = rc.Bottom - nHeightNew

    Case WMSZ_RIGHT
      rc.Right = rc.Left + nWidthNew

    Case WMSZ_BOTTOM
      rc.Bottom = rc.Top + nHeightNew

    Case WMSZ_BOTTOMLEFT
      rc.Bottom = rc.Top + nHeightNew
      rc.Left = rc.Right - nWidthNew

    Case WMSZ_BOTTOMRIGHT
      rc.Bottom = rc.Top + nHeightNew
      rc.Right = rc.Left + nWidthNew

    Case WMSZ_TOPLEFT
      rc.Left = rc.Right - nWidthNew
      rc.Top = rc.Bottom - nHeightNew

    Case WMSZ_TOPRIGHT
      rc.Top = rc.Bottom - nHeightNew
      rc.Right = rc.Left + nWidthNew
  
  End Select
  
  ' Tell the system the new rectangle
  CopyMemory lParam, VarPtr(rc), Len(rc)
  
  ' Trap default processing
  OnSizing = 0

End Function

' Called when the AppBar receives a WM_GETMINMAXINFO message
Friend Function OnGetMinMaxInfo(ByVal lpMinMaxInfo As Long) As Long
  
  Dim mmi As MINMAXINFO
  
  CopyMemory VarPtr(mmi), lpMinMaxInfo, Len(mmi)

  If Edge = abeFloat Then
    With mmi
      .ptMinTrackSize.x = FABS.nMinWidth
      .ptMinTrackSize.y = FABS.nMinHeight
      .ptMaxTrackSize.x = FABS.nMaxWidth
      .ptMaxTrackSize.y = FABS.nMaxHeight
    End With
  Else
    With mmi
      .ptMinTrackSize.x = FABS.szMinDockSize.cx
      .ptMinTrackSize.y = FABS.szMinDockSize.cy
      .ptMaxTrackSize.x = GetSystemMetrics(SM_CXSCREEN)
      .ptMaxTrackSize.y = GetSystemMetrics(SM_CYSCREEN)
      If Not IsEdgeTopOrBottom(Edge) Then
        .ptMaxTrackSize.x = FABS.szMaxDockSize.cx
      End If
      If Not IsEdgeLeftOrRight(Edge) Then
        .ptMaxTrackSize.y = FABS.szMaxDockSize.cy
      End If
    End With
  End If

  CopyMemory lpMinMaxInfo, VarPtr(mmi), Len(mmi)
  
  ' Trap default processing
  OnGetMinMaxInfo = 0

End Function


' AppBar-specific helper functions

' Returns TRUE if abEdge is ABE_LEFT or ABE_RIGHT, else FALSE is returned
Friend Function IsEdgeLeftOrRight(ByVal abEdge As TAppBarEdge) As Boolean
  
  If (abEdge = abeLeft) Or (abEdge = abeRight) Then
    IsEdgeLeftOrRight = True
  Else
    IsEdgeLeftOrRight = False
  End If

End Function
    
' Returns TRUE if abEdge is ABE_TOP or ABE_BOTTOM, else FALSE is returned
Friend Function IsEdgeTopOrBottom(ByVal abEdge As TAppBarEdge) As Boolean

  If (abEdge = abeTop) Or (abEdge = abeBottom) Then
    IsEdgeTopOrBottom = True
  Else
    IsEdgeTopOrBottom = False
  End If

End Function

' Returns TRUE if abEdge is ABE_FLOAT, else FALSE is returned
Friend Function IsFloating(ByVal abEdge As TAppBarEdge) As Boolean

  If abEdge = abeFloat Then
    IsFloating = True
  Else
    IsFloating = False
  End If
  
End Function

' Returns TRUE if abFlags contain an at least allowed edge to dock on
Friend Function IsDockable(ByVal abFlags As TAppBarFlags) As Boolean
  
  IsDockable = abFlags And _
               (abfAllowLeft Or abfAllowTop Or abfAllowRight Or abfAllowBottom)

End Function

' Returns TRUE if abFlags contain abfAllowLeft and abfAllowRight
Friend Function IsDockableVertically(ByVal abFlags As TAppBarFlags) As Boolean

  IsDockableVertically = abFlags And (abfAllowLeft Or abfAllowRight)

End Function

' Returns TRUE if abFlags contain abfAllowTop and abfAllowBottom
Friend Function IsDockableHorizontally(ByVal abFlags As TAppBarFlags) _
                                       As Boolean
  
  IsDockableHorizontally = abFlags And (abfAllowTop Or abfAllowBottom)

End Function

' Forces the shell to update its AppBar list and the workspace area
Friend Function ResetSystemKnowledge()

#If DEBUG_MODE Then
  Dim abd As APPBARDATA
  abd.cbSize = Len(abd)
  abd.hwnd = 0
  SHAppBarMessage ABM_REMOVE, abd
#Else
  ' nothing to do when not in debug mode
#End If

End Function

' Returns a proposed edge or ABE_FLOAT based on ABF_* flags and a
' point specified in screen coordinates
Friend Function GetEdgeFromPoint(ByVal abFlags As TAppBarFlags, _
                                 ByRef pt As POINTS) As TAppBarEdge

  Dim rc As Rect
  Dim cxScreen As Long
  Dim cyScreen As Long
  Dim ptCenter As POINTS
  Dim ptOffset As POINTS
  Dim bIsLeftOrRight As Boolean
  Dim abSubstEdge As TAppBarEdge
  
  ' Let's get floating out of the way first
  If CBool(abfAllowFloat And abFlags) Then

    ' Get the rectangle that bounds the size of the screen
    ' minus any docked (but not-autohidden) AppBars
    SystemParametersInfo SPI_GETWORKAREA, 0, VarPtr(rc), 0

    ' Leave a 1/2 width/height-of-a-scrollbar gutter around the workarea
    InflateRect rc, _
                -GetSystemMetrics(SM_CXVSCROLL), _
                -GetSystemMetrics(SM_CYHSCROLL)

    ' If the point is in the adjusted workarea OR no edges are allowed
    If PtInRect(rc, pt) Or Not IsDockable(abFlags) Then
      ' The AppBar should float
      GetEdgeFromPoint = abeFloat
      Exit Function
    End If
  
  End If

  ' If we get here, the AppBar should be docked; determine the proper edge
  ' Get the dimensions of the screen
  cxScreen = GetSystemMetrics(SM_CXSCREEN)
  cyScreen = GetSystemMetrics(SM_CYSCREEN)

  ' Find the center of the screen
  ptCenter.x = cxScreen \ 2
  ptCenter.y = cyScreen \ 2

  ' Find the distance from the point to the center
  ptOffset.x = pt.x - ptCenter.x
  ptOffset.y = pt.y - ptCenter.y

  ' Determine if the point is farther from the left/right or top/bottom
  bIsLeftOrRight = _
    CBool((Abs(ptOffset.y) * cxScreen) <= (Abs(ptOffset.x) * cyScreen))

  ' Propose an edge
  If bIsLeftOrRight Then
    If 0 <= ptOffset.x Then
      GetEdgeFromPoint = abeRight
    Else
      GetEdgeFromPoint = abeLeft
    End If
  Else
    If 0 <= ptOffset.y Then
      GetEdgeFromPoint = abeBottom
    Else
      GetEdgeFromPoint = abeTop
    End If
  End If

  ' Calculate an edge substitute
  If CBool(abfAllowFloat And abFlags) Then
    abSubstEdge = abeFloat
  Else
    abSubstEdge = FABS.abEdge
  End If

  ' Check if the proposed edge is allowed. If not, return the edge substitute
  Select Case GetEdgeFromPoint
    Case abeLeft
      If Not CBool(abfAllowLeft And abFlags) Then
        GetEdgeFromPoint = abSubstEdge
      End If
    Case abeTop
      If Not CBool(abfAllowTop And abFlags) Then
        GetEdgeFromPoint = abSubstEdge
      End If
    Case abeRight
      If Not CBool(abfAllowRight And abFlags) Then
        GetEdgeFromPoint = abSubstEdge
      End If
    Case abeBottom
      If Not CBool(abfAllowBottom And abFlags) Then
        GetEdgeFromPoint = abSubstEdge
      End If
  End Select

End Function


' PUBLIC


' Public member functions

' Extends the calling form with the AppBar properties and methods
Public Function Extends()
  LinkCallback FormHwnd, Me
  OnCreate

End Function

' Detaches the AppBar behavior from the calling form
Public Function Detach()

  OnDestroy
  DetachCallback

End Function

' Forces the AppBar's visual appearance to match its internal state
Public Function UpdateBar()
  
  Edge = Edge

End Function

'''' Loads settings from the registry at RootKey and KeyName location.
'''' Returns TRUE if the settings are available, else FALSE
'''Public Function LoadSettings() As Boolean
'''
'''  Dim hkSettings As Long
'''  Dim dwType As Long
'''  Dim dwSize As Long
'''  Dim abSettings As TAppBarSettings
'''
'''  ' Set the default return value
'''  LoadSettings = False
'''
'''  ' Open the key where settings are stored
'''  If RegOpenKeyEx(FabSettingsLocation.nRootKey, _
'''                  FabSettingsLocation.KeyName, _
'''                  0, _
'''                  KEY_ALL_ACCESS, _
'''                  hkSettings) = 0 Then
'''
'''    ' Query the settings
'''    dwType = REG_BINARY
'''    dwSize = Len(abSettings)
'''    If RegQueryValueEx(hkSettings, "", 0, dwType, abSettings, dwSize) = 0 Then
'''
'''      ' If the settings have been read, copy them in the FABS record
'''      If dwSize = Len(abSettings) Then
'''        FABS = abSettings
'''        LoadSettings = True
'''      End If
'''
'''    End If
'''
'''  End If
'''
'''  ' Close the key
'''  RegCloseKey hkSettings
'''
'''End Function
'''
'''' Saves settings into the registry at RootKey and KeyName location.
'''' Returns TRUE if succeeded, else FALSE
'''Public Function SaveSettings() As Boolean
'''
'''  Dim hkSettings As Long
'''  Dim dwDisposition As Long
'''
'''  ' Set the default return value
'''  SaveSettings = False
'''
'''  ' Create or open the key where settings must be stored
'''  If RegCreateKeyEx(FabSettingsLocation.nRootKey, _
'''                    FabSettingsLocation.KeyName, _
'''                    0, _
'''                    "", _
'''                    REG_OPTION_NON_VOLATILE, _
'''                    KEY_ALL_ACCESS, _
'''                    0, _
'''                    hkSettings, _
'''                    dwDisposition) = 0 Then
'''
'''    ' Store the settings
'''    If RegSetValueEx(hkSettings, "", 0, REG_BINARY, FABS, Len(FABS)) = 0 Then
'''      ' Return success
'''      SaveSettings = True
'''
'''    End If
'''
'''  End If
'''
'''  ' Close the key
'''  RegCloseKey hkSettings
'''
'''End Function




