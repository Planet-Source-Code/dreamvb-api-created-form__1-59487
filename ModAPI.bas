Attribute VB_Name = "ModAPI"
Public Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function SetMenu Lib "user32.dll" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Public Declare Function CreateMenu Lib "user32.dll" () As Long
Public Declare Function AppendMenu Lib "user32.dll" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function CreatePopupMenu Lib "user32.dll" () As Long

Public Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function MessageBox Lib "user32.dll" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
'Window API Stuff for createing and showing a window
Public Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function UpdateWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

'APIs used for dealing with the window messages
Public Declare Function GetMessage Lib "user32.dll" Alias "GetMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function TranslateMessage Lib "user32.dll" (lpMsg As Msg) As Long
Public Declare Function DispatchMessage Lib "user32.dll" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Public Declare Sub PostQuitMessage Lib "user32.dll" (ByVal nExitCode As Long)
Public Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Window Style consts see API Viewer for more
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_MINIMIZE As Long = &H20000000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_SIZEBOX As Long = WS_THICKFRAME
'API MessageBox consts
Public Const MB_OK As Long = &H0&
Public Const MB_ICONASTERISK As Long = &H40&
Public Const MB_ICONEXCLAMATION As Long = &H30&
Public Const MB_ICONQUESTION As Long = &H20&
Public Const MB_ICONINFORMATION As Long = MB_ICONASTERISK
Public Const MB_YESNO As Long = &H4&

'Default X and Y position of were the window is placed
Public Const CW_USEDEFAULT As Long = &H80000000

Public Const WindowStyle = WS_SYSMENU + WS_CAPTION + WS_MINIMIZEBOX + WS_MAXIMIZEBOX + WS_SIZEBOX
Public Const SW_NORMAL As Long = 1 'Show the window in normal mode also see API View for more

Public WinHwnd As Long, WndDC As Long ' Hangle for the window to be created and the DC Hangle
Public wc As WNDCLASS
'Public wc As WNDCLASS ' Class type information for our window to be created
Public WinMsg As Msg    ' Used to hold the messages of a window

Public WindowCaption As String ' Caption of our the window
Public WinClassName As String   'Our new window's ClassName

'Used for windows messages
Public Const WM_CLOSE As Long = &H10
Public Const WM_DESTROY As Long = &H2
Public Const WM_MOUSEMOVE As Long = &H200
Public Const WM_SIZE As Long = &H5
Public Const WM_CREATE As Long = &H1
Public Const WM_COMMAND As Long = &H111
' Public consts that Identify the menu items been clicked
Public Const DM_MENU_ABOUT = 1
Public Const DM_MENU_EXIT = 2
'Menu consts
Public Const MF_POPUP As Long = &H10&
Public Const MF_APPEND As Long = &H100&
Public Const MF_STRING As Long = &H0&
Public Const MF_SEPARATOR As Long = &H800&

Function LoWord(ByVal DWord As Long) As Integer
    If DWord And &H8000& Then
        LoWord = DWord Or &HFFFF0000
    Else
        LoWord = DWord And &HFFFF&
    End If
End Function

Function HiWord(ByVal DWord As Long) As Integer
    HiWord = (DWord And &HFFFF0000) \ 65536
End Function

