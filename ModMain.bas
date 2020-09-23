Attribute VB_Name = "ModMain"
' Hi All
' This is some code I wipped up about an hour ago and shows you how to create a Window / Form from just API Calls
' Why I am I reiventing the wheel. since I been using C++ I noticed in VB we take eveything for granted.
' I mean almost eveything is done for you. I my self like to see exacly how something is working.
' And if you use C++ then you see exacly that.

' IN VB as we all know a from is simple you click add form and Bingo there you have a form.
' But how did it get there? how was it created by Magic umm I thing not, well with any look this example should expain
' how it is created.
' Ok at the moment the code only support some basic things such as

' Createing a new widnow from class information we give it
' Displaying the new window
' Identifying messages and dealing with them
' Displaying a Message Box with only API
' I also added a menu to the window that is also API Created

' Well hope you like this code I will try and get a new update mabe add some
' other controls to it such as editboxes. static lables and statusbars etc

' PS Note to beginners running this code always make sure you save your work when using this code
' If you want to close down the example do it in the correct way File->Exit or click the X
' Doing this by just by clicking the stop button in VB IDE will crash simple

Sub main()
' This is the part that loads our window first
    WindowCaption = "MY API Created Window" 'Caption for our new window
    WinClassName = "MyWinClass" 'Class name for our window
    
    'Fill the class struc with the information needed for the new window
    With wc
        .lpfnwndproc = GetAddress(AddressOf WinProc)
        .cbClsextra = 0
        .cbWndExtra2 = 0
        .hInstance = App.hInstance
        .lpszMenuName = vbNullString
        .style = 0
        .hbrBackground = 16
        .lpszClassName = WinClassName
    End With
    
    If RegisterClass(wc) = 0 Then ' Check if the windows class was registered
        MessageBox 0, "RegisterClass Faild.", "Error", MB_ICONEXCLAMATION Or MB_OK
        End
    Else
        ' Create the window
        WinHwnd = CreateWindowEx(0&, WinClassName, WindowCaption, _
        WindowStyle, CW_USEDEFAULT, CW_USEDEFAULT, 240, 120, 0, 0, App.hInstance, ByVal 0&)
    
        If WinHwnd = 0 Then ' Check if the window was created
            MessageBox 0, "CreateWindowEx Faild.", "Error", MB_ICONEXCLAMATION Or MB_OK
            Exit Sub
            End
        Else
            WndDC = GetDC(WinHwnd) ' Get the Windows DC
            ShowWindow WinHwnd, SW_NORMAL ' Show the window in normal mode
            UpdateWindow WinHwnd ' Update the new window
            
            'Do the Message Loop
            Do While GetMessage(WinMsg, WinHwnd, 0, 0) > 0
                TranslateMessage WinMsg
                DispatchMessage WinMsg
                DoEvents
            Loop
        End If
    End If
    
End Sub

