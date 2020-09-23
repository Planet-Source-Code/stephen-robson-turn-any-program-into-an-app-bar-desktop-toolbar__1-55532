Attribute VB_Name = "drMemory"
Option Explicit
'=======================================================================
'
' drMemory   - Cross-Process Memory Buffer support
'
'  (c) 2003  "Dr Memory" ==> Jim White
'            MathImagics
'            Uki, NSW, Australia
'            Puttenham, Surrey, UK
'
'    This module contains functions that provide both WinNT and Win9x
'    style cross-process memory buffer allocation, read and write
'    functions.
'
'    These functions are typically required when trying to use
'    SendMessage to exchange data with windows in another process.
'
'
'=======================================================================
' Usage guide:
'
'   1.  Allocate buffer(s) in the target process
'
'    xpWindow& = <<target control window handle>>
'    xpBuffer& = drMemoryAlloc(xpWindow, nBytes)
'
'   2.  Prepare the data to be passed to the control
'       and copy it into the buffer
'
'    drMemoryWrite xpBuffer, myBuffer, nBytes
'
'   3.  SendMessage xpWindow, MSG_CODE, wParam, ByVal xpBuffer
'                                               ==============
'
'   4.  Extract return data
'
'    drMemoryRead  xpBuffer, myBuffer, nBytes
'
'    (repeat 3/4 as necessary)
'
'   5.  Release the buffer
'
'    drMemoryFree  xpBuffer
'
'=======================================================================
'
   Private PlatformKnown As Boolean  ' have we identified the platform?
   Public NTflag         As Boolean  ' if so, are we NT family (NT, 2K,XP) or non-NT (9x)?
   Public XPflag         As Boolean
   
   Private fpHandle      As Long     ' the foreign-process instance handle. When we want
                                     ' memory on NT platforms, this is returned to us by
                                     ' OpenProcess, and we pass it in to VirtualAllocEx.
                                     
                                     ' We must preserve it, as we need it for read/write
                                     ' operations, and to release the memory when we've
                                     ' finished with it.
                                     
      ' For this reason, on NT/2K/XP platforms this module should only be used to
      '    interface with ONE TARGET PROCESS at a time. In the future I'll rewrite
      '    this as a class, which can handle multiple-targets, automatic allocation
      '    de-allocation, etc
      '
         
'
'================== Platform Identification is necessary!
'
   Public Type OSVERSIONINFO
      dwOSVersionInfoSize As Long
      dwMajorVersion As Long
      dwMinorVersion As Long
      dwBuildNumber As Long
      dwPlatformId As Long
      szCSDVersion As String * 128
      End Type
   Public WIN As OSVERSIONINFO

   Private Declare Function GetVersionEx Lib "KERNEL32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
'
'================== Win95/98   Process Memory functions
   Private Declare Function CreateFileMapping Lib "KERNEL32" Alias "CreateFileMappingA" (ByVal hFile As Long, ByVal lpFileMappigAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
   Private Declare Function MapViewOfFile Lib "KERNEL32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
   Private Declare Function UnmapViewOfFile Lib "KERNEL32" (lpBaseAddress As Any) As Long
   Private Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
'
'================== WinNT/2000 Process Memory functions
   Private Declare Function OpenProcess Lib "KERNEL32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
   Private Declare Function VirtualAllocEx Lib "KERNEL32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
   Private Declare Function VirtualFreeEx Lib "KERNEL32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
   Private Declare Function WriteProcessMemory Lib "KERNEL32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As Long, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
   Private Declare Function ReadProcessMemory Lib "KERNEL32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As Long, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
'
'
'================== Common Platform
'
   Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

   Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSource As Long, ByVal cBytes As Long)
   Private Declare Function lstrlenA Lib "KERNEL32" (ByVal lpsz As Long) As Long
   Private Declare Function lstrlenW Lib "KERNEL32" (ByVal lpString As Long) As Long
   Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
   Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

' ----------
   Const PAGE_READWRITE = &H4
   Const MEM_RESERVE = &H2000&
   Const MEM_RELEASE = &H8000&
   Const MEM_COMMIT = &H1000&
   Const PROCESS_VM_OPERATION = &H8
   Const PROCESS_VM_READ = &H10
   Const PROCESS_VM_WRITE = &H20
   Const STANDARD_RIGHTS_REQUIRED = &HF0000
   Const SECTION_QUERY = &H1
   Const SECTION_MAP_WRITE = &H2
   Const SECTION_MAP_READ = &H4
   Const SECTION_MAP_EXECUTE = &H8
   Const SECTION_EXTEND_SIZE = &H10
   Const SECTION_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE
   Const FILE_MAP_ALL_ACCESS = SECTION_ALL_ACCESS


Public Function drMemoryAlloc(ByVal xpWindow As Long, ByVal nBytes As Long) As Long
   '
   ' Returns pointer to a share-able buffer (size nBytes) in target process
   '   that owns xpWindow
   '
   Dim xpThread As Long    ' target control's thread id
   Dim xpID As Long        '                  process id
   If WindowsNT Then
      xpThread = GetWindowThreadProcessId(xpWindow, xpID)
      drMemoryAlloc = VirtualAllocNT(xpID, nBytes)
   Else
      drMemoryAlloc = VirtualAlloc9X(nBytes)
      End If
   End Function

Public Sub drMemoryFree(ByVal mPointer As Long)
   If WindowsNT Then
      VirtualFreeNT mPointer
   Else
      VirtualFree9X mPointer
      End If
   End Sub
   
Public Sub drMemoryRead(ByVal xpBuffer As Long, ByVal myBuffer As Long, ByVal nBytes As Long)
   If WindowsNT Then
      ReadProcessMemory fpHandle, xpBuffer, myBuffer, nBytes, 0
   Else
      CopyMemory myBuffer, xpBuffer, nBytes
      End If
   End Sub

Public Sub drMemoryWrite(ByVal xpBuffer As Long, ByVal myBuffer As Long, ByVal nBytes As Long)
   If WindowsNT Then
      WriteProcessMemory fpHandle, xpBuffer, myBuffer, nBytes, 0
   Else
      CopyMemory xpBuffer, myBuffer, nBytes
      End If
   End Sub

Public Function WindowsNT() As Boolean
   ' return TRUE if NT-like platform (NT, 2000, XP, etc)
   If Not PlatformKnown Then GetWindowsVersion
   WindowsNT = NTflag
   End Function

Public Function WindowsXP() As Boolean
   ' return TRUE only if XP
   If Not PlatformKnown Then GetWindowsVersion
   WindowsXP = NTflag And (WIN.dwMinorVersion <> 0)
   End Function

Public Sub GetWindowsVersion()
   WIN.dwOSVersionInfoSize = Len(WIN)
   If (GetVersionEx(WIN)) = 0 Then Exit Sub  ' in deep doo if this fails
   NTflag = (WIN.dwPlatformId = 2)
   XPflag = NTflag And (WIN.dwMajorVersion = 5) And (WIN.dwMinorVersion > 0)
   PlatformKnown = True
   End Sub

'============================================
'  The NT/2000 Allocate and Release functions
'============================================

Private Function VirtualAllocNT(ByVal fpID As Long, ByVal memSize As Long) As Long
   fpHandle = OpenProcess(PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE, False, fpID)
   VirtualAllocNT = VirtualAllocEx(fpHandle, ByVal 0&, ByVal memSize, MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
   End Function

Private Sub VirtualFreeNT(ByVal MemAddress As Long)
   Call VirtualFreeEx(fpHandle, ByVal MemAddress, 0&, MEM_RELEASE)
   CloseHandle fpHandle
   End Sub

'============================================
'  The 95/98 Allocate and Release functions
'============================================

Private Function VirtualAlloc9X(ByVal memSize As Long) As Long
   fpHandle = CreateFileMapping(&HFFFFFFFF, 0, PAGE_READWRITE, 0, memSize, vbNullString)
   VirtualAlloc9X = MapViewOfFile(fpHandle, FILE_MAP_ALL_ACCESS, 0, 0, 0)
   End Function

Private Sub VirtualFree9X(ByVal lpMem As Long)
   UnmapViewOfFile lpMem
   CloseHandle fpHandle
   End Sub

Public Function dmWindowClass(ByVal hWindow As Long) As String
   Dim className As String, cLen As Long
   className = String(64, 0)
   cLen = GetClassName(hWindow, className, 63)
   If cLen > 0 Then className = Left(className, cLen)
   dmWindowClass = className
   End Function



