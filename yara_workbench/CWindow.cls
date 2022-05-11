VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public hwnd As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, ByRef hINst As Long) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 
Private Declare Function AttachConsole Lib "kernel32" (ByVal dwProcessId As Long) As Boolean
Private Declare Function FreeConsole Lib "kernel32" () As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function GetConsoleScreenBufferInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
Private Declare Function ReadConsoleOutputCharacter Lib "kernel32" Alias "ReadConsoleOutputCharacterW" (ByVal hConsoleOutput As Long, ByVal lpCharacter As Long, ByVal nLength As Long, ByVal dwReadCoord As Long, lpNumberOfCharsRead As Long) As Long



Private Const WM_SYSCOMMAND  As Long = &H112
Private Const SC_MINIMIZE As Long = &HF020
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const WPF_ASYNCWINDOWPLACEMENT = 4

Private Type POINTAPI
        x As Long
        Y As Long
End Type
 
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
 
Private Const STD_OUTPUT_HANDLE         As Long = -11

Private Type COORD
    x As Integer
    Y As Integer
End Type

Private Type SMALL_RECT
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type

Private Type CONSOLE_SCREEN_BUFFER_INFO
    dwSize As COORD
    dwCursorPosition As COORD
    wAttributes As Integer
    srWindow As SMALL_RECT
    dwMaximumWindowSize As COORD
End Type

 
Private Type WINDOWPLACEMENT
        Length As Long
        Flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type

Private Enum ProcessAccessTypes
      PROCESS_TERMINATE = (&H1)
      PROCESS_CREATE_THREAD = (&H2)
      PROCESS_SET_SESSIONID = (&H4)
      PROCESS_VM_OPERATION = (&H8)
      PROCESS_VM_READ = (&H10)
      PROCESS_VM_WRITE = (&H20)
      PROCESS_DUP_HANDLE = (&H40)
      PROCESS_CREATE_PROCESS = (&H80)
      PROCESS_SET_QUOTA = (&H100)
      PROCESS_SET_INFORMATION = (&H200)
      PROCESS_QUERY_INFORMATION = (&H400)
      STANDARD_RIGHTS_REQUIRED = &HF0000
      SYNCHRONIZE = &H100000
      PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
End Enum

Enum eWindowStates
    SW_HIDE = 0  'Hides the window and activates another window.
    SW_MAXIMIZE = 3 'Maximizes the specified window.
    SW_MINIMIZE = 6 'Minimizes the specified window and activates the next top-level window in the z-order.
    SW_RESTORE = 9  'Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window.
    SW_SHOW = 5     'Activates the window and displays it in its current size and position.
    SW_SHOWMAXIMIZED = 3 'Activates the window and displays it as a maximized window.
    SW_SHOWMINIMIZED = 2 'Activates the window and displays it as a minimized window.
    SW_SHOWMINNOACTIVE = 7 'Displays the window as a minimized window. This value is similar to SW_SHOWMINIMIZED, except the window is not activated.
    SW_SHOWNA = 8  'Displays the window in its current size and position. This value is similar to SW_SHOW, except the window is not activated.
    SW_SHOWNOACTIVATE = 4 'Displays a window in its most recent size and position. This value is similar to SW_SHOWNORMAL, except the window is not activated.
    SW_SHOWNORMAL = 1
End Enum

Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, wParam As Any, lParam As Any) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClassName As String, ByVal lpszWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Const LB_GETCOUNT = &H18B
Private Const LB_GETTEXT = &H189
Private Const LB_GETTEXTLEN = &H18A
Private Const CB_GETCOUNT = &H146
Private Const CB_GETLBTEXT = &H148
Private Const CB_GETLBTEXTLEN = &H149
Private Const WM_CLOSE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const GWL_EXSTYLE    As Long = -20
Private Const WS_EX_TOPMOST As Long = 8

Property Get Visible() As Boolean
    Visible = IIf(IsWindowVisible(Me.hwnd) = 0, False, True)
End Property

Property Let TopMost(v As Boolean)
    Dim Flags As Long, zflags As Long
    Flags = SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    zflags = IIf(v, HWND_TOPMOST, HWND_NOTOPMOST)
    SetWindowPos hwnd, zflags, 0, 0, 0, 0, Flags
End Property

Property Get TopMost() As Boolean
    Dim exStyle As Long
    exStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    TopMost = ((exStyle And WS_EX_TOPMOST) = WS_EX_TOPMOST)
End Property

Sub CloseWindow()
    PostMessage hwnd, WM_CLOSE, 0, 0
End Sub

Property Get isValid() As Boolean
    isValid = Not (IsWindow(hwnd) = 0)
End Property

Function FindChild(ByVal className As String) As CWindow
    
    Dim hChild As Long
    Dim w As New CWindow
    
    w.hwnd = FindWindowEx(hwnd, 0&, className, vbNullString)
    Set FindChild = w 'can be 0 check isValid...
    
End Function

Property Get ProcessPID() As Long
    Dim pid As Long
    GetWindowThreadProcessId hwnd, pid
    ProcessPID = pid
End Property

Property Get Caption() As String
    Dim hWndlength As Long, hWndTitle As String, a As Long
    
    hWndlength = GetWindowTextLength(hwnd)
    hWndTitle = String$(hWndlength + 10, 0)
    a = GetWindowText(hwnd, hWndTitle, (hWndlength + 1))
    Caption = hWndTitle
    
End Property

Property Get className() As String
    Dim lpClassName As String, retVal As Long
    lpClassName = Space(256)
    retVal = GetClassName(hwnd, lpClassName, 256)
    className = Left$(lpClassName, retVal)
End Property

Property Get WindowState() As eWindowStates
    Dim wp As WINDOWPLACEMENT
    wp.Length = Len(wp)
    GetWindowPlacement hwnd, wp
    WindowState = wp.showCmd
End Property

Property Let WindowState(s As eWindowStates)
    Dim wp As WINDOWPLACEMENT
    wp.Length = Len(wp)
    GetWindowPlacement hwnd, wp
    wp.Length = Len(wp)
    wp.showCmd = s
    wp.Flags = WPF_ASYNCWINDOWPLACEMENT
    SetWindowPlacement hwnd, wp
End Property

Property Get WindowStateString() As String
    
    Dim e As eWindowStates, s As String
    
    e = WindowState
    
    If e = SW_HIDE Then s = "SW_HIDE"
    If e = SW_MAXIMIZE Then s = "SW_MAXIMIZE"
    If e = SW_MINIMIZE Then s = "SW_MINIMIZE"
    If e = SW_RESTORE Then s = "SW_RESTORE"
    If e = SW_SHOW Then s = "SW_SHOW"
    If e = SW_SHOWMAXIMIZED Then s = "SW_SHOWMAXIMIZED"
    If e = SW_SHOWMINIMIZED Then s = "SW_SHOWMINIMIZED"
    If e = SW_SHOWMINNOACTIVE Then s = "SW_SHOWMINNOACTIVE"
    If e = SW_SHOWNA Then s = "SW_SHOWNA"
    If e = SW_SHOWNOACTIVATE Then s = "SW_SHOWNOACTIVATE"
    If e = SW_SHOWNORMAL Then s = "SW_SHOWNORMAL"
    If Len(s) = 0 Then s = "Unknown: " & e
    
    WindowStateString = s
    
End Property

Property Get ProcessPath() As String
    Dim hProc As Long, pid As Long
    Dim hMods() As Long, cbAlloc As Long, ret As Long, retMax As Long
    Dim sPath As String
    
    GetWindowThreadProcessId hwnd, pid
    
    hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, False, pid)
    
    If hProc <> 0 Then
        cbAlloc = 200
        ReDim hMods(cbAlloc)
        ret = EnumProcessModules(hProc, hMods(0), (cbAlloc * 4), retMax)
                
        sPath = Space$(260)
        ret = GetModuleFileNameExA(hProc, hMods(0), sPath, 260)
        ProcessPath = Left$(sPath, ret)
        
        Call CloseHandle(hProc)
    End If
    
End Property

Public Sub MinimizeWindow()
    PostMessage hwnd, WM_SYSCOMMAND, SC_MINIMIZE, 0
End Sub




