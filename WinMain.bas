Attribute VB_Name = "modWin"
Private Sub PrintToToWindow(lpStr As String, Optional xPos As Long = 10, Optional yPos As Long = 10)
    ' All the little function does is print out some text to the window
    TextOut WndDC, xPos, yPos, lpStr, Len(lpStr)
End Sub

Function WinProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' This function is used to hangle all the messages that the window will recive
    Dim sTmp As String
    Dim ans As Integer
    Dim hMenu As Long, hSubMenu As Long
    
    Select Case wMsg
        Case WM_CREATE ' Window is created
            hMenu = CreateMenu() ' Create a new menu for our window
            hSubMenu = CreatePopupMenu() ' Create a popup menu that we put out sub-items into
            AppendMenu hMenu, MF_STRING Or MF_POPUP, hSubMenu, "&File" ' Top Level Menu
            AppendMenu hSubMenu, MF_STRING, DM_MENU_ABOUT, "&About" ' Sub item
            AppendMenu hSubMenu, MF_STRING, DM_MENU_EXIT, "E&xit.." ' Sub Item
            AppendMenu hSubMenu, MF_SEPARATOR, -1, 0&
            AppendMenu hSubMenu, MF_STRING, 3, "Add other items here" ' Sub Item
            SetMenu hwnd, hMenu ' update the window with our menu
            MessageBox WinHwnd, "Your API Window has been created", "VB-API Window App", MB_OK Or MB_ICONINFORMATION
            Exit Function
        Case WM_COMMAND
            Select Case wParam
                Case DM_MENU_EXIT ' menu exit
                    SendMessage hwnd, WM_CLOSE, ByVal 0&, ByVal 0&
                    Exit Function
                Case DM_MENU_ABOUT
                    ' Display an about messagebox
                    MessageBox hwnd, "ALL API Window Created by Ben Jones", "About..", MB_ICONINFORMATION Or MB_OK
                    Exit Function
            End Select
            
        Case WM_CLOSE ' User has clicked the X on the form so we need to destroy the window
            ans = MessageBox(WinHwnd, "Do you want to quite this program now?", "Quit...", MB_ICONQUESTION Or MB_YESNO)
            If ans = vbNo Then Exit Function ' If the users answer was yes then we can then destroy the window
            DestroyWindow WinHwnd
        'Case WM_DESTROY ' Using this seems to close down the VB IDE so make sure you save any work first
            'PostQuitMessage 0
        Case WM_MOUSEMOVE
            UpdateWindow WinHwnd
            ' Use is moveing the mouse lets show the x and y positions
            sTmp = "x = " & LoWord(lParam) & ", " & " y = " & HiWord(lParam) ' Str(LoWord(lParam) & " , " & Str(HiWord(lParam)))
            PrintToToWindow sTmp, 10, 10
            sTmp = ""
            Exit Function
        Case WM_SIZE ' Window is resizeing
            MessageBox WinHwnd, "Windows has been resize.", "WM_SIZE", MB_OK Or MB_ICONINFORMATION
            Exit Function
        Case Else
            WinProc = DefWindowProc(hwnd, wMsg, wParam, lParam)  ' Keep fireing the messages back to the window
    End Select
    
End Function

Function GetAddress(Address) As Long
    GetAddress = Address
End Function
