VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Program Appbar Creator"
   ClientHeight    =   5160
   ClientLeft      =   165
   ClientTop       =   435
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   344
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   204
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pct1 
      Height          =   5055
      Left            =   0
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   205
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdNewWindow 
         Caption         =   "Choose New Window"
         Height          =   495
         Left            =   720
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
      Begin VB.Frame fraGroup 
         Caption         =   "Choose WIndow to Capture"
         Height          =   4215
         Left            =   0
         TabIndex        =   2
         Top             =   600
         Width           =   3015
         Begin VB.CommandButton cmdCaptureWindowWithTitle 
            Caption         =   "Capture Window with Title..."
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   3240
            Width           =   2775
         End
         Begin VB.CommandButton cmdRefreshList 
            Caption         =   "Refresh List"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   3720
            Width           =   2775
         End
         Begin VB.CommandButton cmdCapture 
            Caption         =   "Capture Selected Window"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   2760
            Width           =   2775
         End
         Begin VB.ListBox lst1 
            Height          =   2400
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Options"
         Height          =   495
         Left            =   1920
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const HWND_TOP = 0
Const SW_SHOWNORMAL = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Dim destHwnd As Long
Private PrevOwner As Long

Dim Taskbar As New SteWare_TaskBar.clsTaskBarPrograms

Private Sub cmdCapture_Click()
    'the user clicked to capture a window
    If Me.lst1.ListIndex <> -1 Then 'if something was selected
        'find the handle to the window we want
        destHwnd = Taskbar.TaskbarData(Me.lst1.ListIndex + 1).WindowHandle
'        destHwnd = FindWindow(vbNullString, Title)
        If destHwnd <> 0 Then Capture (destHwnd) 'now capture that window
    End If
End Sub
Function Capture(dhwnd As Long)
    'This function captures a window and places it onto our window.
    'This works because the desktop is just a Window (remember Windows
    '3.1?) so we just change who owns this window.
    PrevOwner = GetParent(dhwnd) 'store the previous owner of the window so that we can restore it later
    StoreWindowSizeData dhwnd 'store the window size data
    SetParent dhwnd, Me.hwnd 'this changes the owner of the window to our form
    'position the window
    SetWindowPos dhwnd, HWND_TOP, 2, Me.cmdQuit.Height + Me.cmdQuit.Top + 2, Me.ScaleWidth - 4, Me.ScaleHeight - (Me.cmdQuit.Height + Me.cmdQuit.Top + 2), SWP_NOACTIVATE
    Me.fraGroup.Visible = False 'hide the window selection controls
End Function
Function Release()
    'This function puts the window back to where it belongs
    Dim Title As String: Title = Me.lst1.List(Me.lst1.ListIndex)
    RestoreWindowSizeData destHwnd 'restore the Window's size
    SetParent destHwnd, PrevOwner 'set the parent of the window
    destHwnd = -1
End Function

Private Sub cmdCaptureWindowWithTitle_Click()
    'This allows the user to enter the caption of a window which doesn't
    'appear in the taskbar.  The program will then attempt to capture that
    'window.
    Dim Title As String
    Title = InputBox("Enter the window's title", "Choose Title") 'get the title
    If Title <> "" Then 'if the title isn't blank
        destHwnd = FindWindow(vbNullString, Title) 'find the handle to the window with that title
        If destHwnd <> 0 Then 'if it exists
            Capture destHwnd 'capture it
        Else 'if it doesn't exist then alert the user
            MsgBox "The window with the title '" & Title & "' could not be found.", vbInformation, "Window Not Found"
        End If
    End If
End Sub

Private Sub cmdNewWindow_Click()
    'User wants to capture a different window
    Release
    Me.fraGroup.Visible = True 'show the selection controls
End Sub

Private Sub cmdOptions_Click()
    frmOptions.Show
End Sub

Private Sub cmdQuit_Click()
    Unload Me 'quit the application
End Sub

Private Sub cmdRefreshList_Click()
    'This function refreshes the list of available windows.
    Taskbar.RefreshArray 'refresh the array
    Me.lst1.Clear 'clear the list box
    Dim i As Integer, Info As TaskbarInfo
    For i = Taskbar.ArrayLBound To Taskbar.ArrayUBound 'for all windows
        Info = Taskbar.TaskbarData(i)
        Me.lst1.AddItem Info.WindowCaption 'add the caption to the listbox
    Next
End Sub

Private Sub Form_Load()
    'start the program
    Call cmdRefreshList_Click 'display the windows available
    
    'appbar stuff
    AppBar.SelfHwnd = Me.hwnd
    AppBar.SlideEffect = False
    AppBar.VertDockSize = Me.ScaleWidth
    AppBar.HorzDockSize = Me.ScaleHeight
    AppBar.Edge = abeRight
    AppBar.Extends 'this applies the AppBar
End Sub

Private Sub Form_Resize()
    'when our form is resized, so must the captured form be resized
    Me.pct1.Width = Me.ScaleWidth
    Me.pct1.Height = Me.ScaleHeight
    SetWindowPos destHwnd, HWND_TOP, 2, Me.cmdQuit.Height + Me.cmdQuit.Top + 2, Me.ScaleWidth - 4, Me.ScaleHeight - (Me.cmdQuit.Height + Me.cmdQuit.Top + 2), SWP_NOACTIVATE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Release 'return the captured form to its owner
    AppBar.Detach 'remove the appbar and hence restore the desktop
End Sub
