VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTaskBarPrograms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'This class acts as a wrapper for the TaskBar and drMemory modules written by Jim White.

Public Type TaskbarInfo
    WindowCaption As String
    WindowHandle As Long
    WindowClass As String
End Type

Private TaskbarArray() As TaskbarInfo

Property Get TaskbarData(Index As Integer) As TaskbarInfo
    'Gets the information for a taskbar program
    If Index >= LBound(TaskbarArray) And Index <= UBound(TaskbarArray) Then
        TaskbarData = TaskbarArray(Index)
    End If
End Property
'These two properties get the upper and lower bounds of the TaskbarArray
Property Get ArrayUBound() As Integer
    ArrayUBound = UBound(TaskbarArray)
    End Property
    Property Get ArrayLBound() As Integer
    ArrayLBound = LBound(TaskbarArray)
End Property

Public Function RefreshArray()
'This function refreshes the array with the current data.
'This is based upon a similar function written by Jim White and much of this
'code must go to him.

    Dim hTaskBar As Long:   hTaskBar = GetTaskBarHandle ' find the TaskBar Toolbar handle
    Dim tbWindow() As Long
    If XPflag Then
        GetTaskBarListXP hTaskBar, tbWindow
    Else
        GetTaskBarList hTaskBar, tbWindow
    End If

    Dim w As Long
    Dim wClass  As String
    Dim wText   As String
    Dim j       As Integer
    Const WM_GETTEXT = &HD
    ReDim TaskbarArray(1 To UBound(tbWindow))
    For w = 1 To UBound(tbWindow)
        wClass = Space(64)
        j = GetClassName(tbWindow(w), wClass, 63)
        wClass = Left(wClass, j)
           
        wText = Space(256)
        j = SendMessageS(tbWindow(w), WM_GETTEXT, 255, wText)
        wText = Left(wText, j)
        'Fill the array
        TaskbarArray(w).WindowCaption = wText
        TaskbarArray(w).WindowClass = wClass
        TaskbarArray(w).WindowHandle = tbWindow(w)
        'tbWindow(w) is the window handle
        'wClass is the class name running the process
        'wText is the Window caption
    Next
End Function

Private Sub Class_Initialize()
    'populate the array when the object is loaded
    RefreshArray
End Sub
