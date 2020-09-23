Attribute VB_Name = "modWindowSize"
Option Explicit

'This module stores and restores the captured window's size

Private Const SW_MINIMIZE = 6
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type WINDOWPLACEMENT
        Length As Long
        flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Dim WinEst As WINDOWPLACEMENT

Function StoreWindowSizeData(hwnd As Long)
    'Tip submitted by pyp99 (pyp99@hotmail.com)
    Dim rtn As Long
    WinEst.Length = Len(WinEst)
    'get the current window placement
    rtn = GetWindowPlacement(hwnd, WinEst)
End Function
Function RestoreWindowSizeData(hwnd As Long)
    SetWindowPlacement hwnd, WinEst
End Function
