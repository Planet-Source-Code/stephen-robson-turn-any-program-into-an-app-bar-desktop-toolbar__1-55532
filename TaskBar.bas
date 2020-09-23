Attribute VB_Name = "TaskBar"
Option Explicit

Public Const TCM_FIRST = &H1300&                ' Tab control messages
Public Const TCM_GETITEMCOUNT = (TCM_FIRST + 4)
Public Const TCM_GETITEMA = (TCM_FIRST + 5)

'  Tab control item structure
Private Type TCITEM
   mask As Long
   dwState As Long
   dwStateMask As Long
   pszText As Long
   cchTextMax As Long
   iImage As Long
   lParam As Long  ' TaskBar items have window handle here!!
   End Type

Type TBBUTTON      ' ToolBar items (XP TaskBar) don't have window
   iBitmap As Long '   handle, but we can get text and use FindWindow
   idCommand As Long
   fsState As Byte
   fsStyle As Byte
   bReserved1 As Byte
   bReserved2 As Byte
   dwData As Long
   iString As Long
   End Type

Public Const WM_USER = &H400
Public Const TB_GETBUTTON = (WM_USER + 23)
Public Const TB_BUTTONCOUNT = (WM_USER + 24)
Public Const TB_GETBUTTONTEXTA = (WM_USER + 45)


'========================================================================================

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Declare Function SendMessageS Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As String) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
   (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long

Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long


Function GetTaskBarHandle() As Long
   '
   ' Find the handle of the TaskBar's Tab Control window
   '   (it's a Toolbar on Win XP).
   '
   Dim w As Long
   GetWindowsVersion
   
   w = FindWindowEx(0, 0, "Shell_TrayWnd", vbNullString)
   w = FindWindowEx(w, 0, "ReBarWindow32", vbNullString)
   w = FindWindowEx(w, 0, "MSTaskSwWClass", vbNullString)
   If XPflag Then
      w = FindWindowEx(w, 0, "ToolBarWindow32", vbNullString)
   Else
      w = FindWindowEx(w, 0, "SysTabControl32", vbNullString)
      End If
   GetTaskBarHandle = w
   End Function

Sub GetTaskBarList(ByVal hTab As Long, wList() As Long)

   '
   ' Retrieve list of TaskBar "application" handles from the tabs
   '   in the TaskBar "Tab" control on all systems other than XP
   '
   ' The window handles are stored in each tabs TCITEM.lParam field
   '      MathImagics (May 2004)
   '-----------------------------------------------------------------
   Dim xpBuffer As Long    ' address of cross-process buffer
   
   Dim bInfo As TCITEM     ' individual Tab/buton info (incl ID)
   Dim nItems As Integer, tbIndex As Long, lret As Long
   '
   ' Get button count
   '
   nItems = SendMessage(hTab, TCM_GETITEMCOUNT, 0&, 0&)
   If nItems <= 0 Then Exit Sub
   ReDim wList(nItems)
    
   xpBuffer = drMemoryAlloc(hTab, 32)
   Const TCIF_PARAM = &H8&
   bInfo.mask = TCIF_PARAM
   drMemoryWrite xpBuffer, VarPtr(bInfo), Len(bInfo)
   
   For tbIndex = 1 To nItems
      SendMessage hTab, TCM_GETITEMA, tbIndex - 1, xpBuffer
      drMemoryRead xpBuffer, VarPtr(bInfo), Len(bInfo)
      wList(tbIndex) = bInfo.lParam
   Next tbIndex
   drMemoryFree xpBuffer
   End Sub
   

Sub GetTaskBarListXP(ByVal hToolbar As Long, wList() As Long)

   '
   ' Retrieve list of TaskBar "application" handles from the buttons
   '   in the TaskBar "ToolBar" control on Windows XP
   '
   Dim xpBuffer As Long    ' address of cross-process buffer
   
   Dim bInfo As TBBUTTON   ' individual button info (incl ID)
   Dim bText As String     '                   caption
   
   Dim nItems As Integer, tbIndex As Long, lret As Long
   Dim w As Long, nw As Long, n As Long
 
   '
   ' Get button count
   '
   nItems = SendMessage(hToolbar, TB_BUTTONCOUNT, 0&, 0&)
   If nItems <= 0 Then Exit Sub
   
   ReDim wList(nItems)
   xpBuffer = drMemoryAlloc(hToolbar, 1024)

   For tbIndex = 0 To nItems - 1
      '
      ' TB_GETBUTTON
      '
      SendMessage hToolbar, TB_GETBUTTON, tbIndex, xpBuffer
      drMemoryRead xpBuffer, VarPtr(bInfo), Len(bInfo)
      
      If bInfo.iBitmap >= 0 And bInfo.iBitmap < 128 Then
         '
         ' we can only get window handles via the button text
         '    so we have to look for possible multiple copies
         '    of the same window, and avoid duplicates
         '
         lret = SendMessage(hToolbar, TB_GETBUTTONTEXTA, bInfo.idCommand, xpBuffer)
         bText = ""
         If lret > 0 Then
            bText = String$(lret + 1, 0)
            drMemoryRead xpBuffer, StrPtr(bText), lret
            bText = Left(StrConv(bText, vbUnicode), lret)
            End If
         w = FindWindowEx(0, 0, vbNullString, bText)
         Do While w
            For n = 1 To nw
               If wList(n) = w Then n = 0: Exit For
               Next
            If n Then  ' add new window to list
               nw = nw + 1
               wList(nw) = w
               End If
            w = FindWindowEx(0, w, vbNullString, bText)
            Loop
         End If
      Next tbIndex
   ReDim Preserve wList(nw)
   drMemoryFree xpBuffer
   End Sub

