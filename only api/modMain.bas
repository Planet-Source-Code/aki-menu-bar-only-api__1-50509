Attribute VB_Name = "modMain"
Option Explicit
DefLng A-Z
' The original project in c++ was not written by me but I posted that project in vb section while ago in order to see
' if anybody can write some kind of the example in vb
' You can find that project here:
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=4601&lngWId=3

' Since nobody wrote an example I decided to do it and today
' I was impressed by creating controls with only API from this code:
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=44043&lngWId=1
' You will find almost nothing from that code here except it helped me to start coding
' and I must say thanks to the author, Viktor E

' Thanks also to www.vbaccelerator for three, ready to use, functions ( for menu selection, highlight color)
'...
' I didn't want to improve a lot in case someone wants to look and compare those two projects, and yes, you can cut it a lot
' Feel free to cut for your own needs

' In case you improve this project I would like to see it, if you don't mind
' You can contact me at: aniram@zahav.net.il

' For the end...
' There is one similar project in c++ I posted but is a bit more complicated and different
' I noticed that PSC deleted that code so i have no link to give you
' In case someone is interested I will email the project
' When I find some more free time I will translate it to vb

'  I also recommend visiting: www.allapi.net for API declarations
Private Enum IDM
    IDM_M1 = 601
    IDM_M2 = 602
    IDM_M3 = 603
    IDM_M4 = 604
    IDM_FILE_EXIT = 700
    IDM_EDIT_CUT = 701
    IDM_EDIT_COPY = 702
    IDM_EDIT_PASTE = 703
    IDM_WINDOW_CASCADE = 704
    IDM_WINDOW_TILE = 705
    IDM_HELP_ABOUT = 706
End Enum

Private TBPRESS As TBBUTTONINFO, TBUNPRESS As TBBUTTONINFO
Private hMenu As Long
Private KeyboardHook As Long, ProcKeyboardHook As Long
Private index As Long, bckIndex As Long, Hot As Long
Private sys As Boolean
Private TextBox As Long, Coolbar As Long, Toolbar  As Long, StatusBar  As Long
Private Main_Form As Long
Private ProcTextBox As Long

Public Function FormProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim NMH As NMHDR
Select Case iMsg

    Case WM_SHOWWINDOW
    Dim RC As RECT
    Dim x As Long, y As Long
        
        Main_Form = hwnd
        'center the form
        GetClientRect hwnd, RC
        x = GetSystemMetrics(SM_CXSCREEN) / 2
        x = x - (RC.Right - RC.Left) / 2
        
        y = GetSystemMetrics(SM_CYSCREEN) / 2
        y = y - (RC.Bottom - RC.Top) / 2
        
        SetWindowPos hwnd, 0, x, y, 0, 0, SWP_NOSIZE
    
    Case WM_CREATE
        Dim I As Long
        Dim BUTS(0 To 3) As TBBUTTON
        Dim BUT As TBBUTTON
        Dim CMDS(0 To 3) As Long
        Dim dwStyle As Long
        '=================================TOOLBAR=====================================================
        CMDS(0) = IDM.IDM_M1
        CMDS(1) = IDM.IDM_M2
        CMDS(2) = IDM.IDM_M3
        CMDS(3) = IDM.IDM_M4
        
        dwStyle = WS_CHILD Or WS_VISIBLE Or TBSTYLE_FLAT Or CCS_NODIVIDER Or CCS_NORESIZE Or CCS_BOTTOM Or TBSTYLE_AUTOSIZE
        Toolbar = CreateWindowEx(WS_EX_TOOLWINDOW, "ToolbarWindow32", "", _
                                                        dwStyle, 0, 0, 0, 0, hwnd, ICC_COOL_CLASSES, App.hInstance, 0&)
                                                        
        SendMessageLong Toolbar, TB_SETPARENT, hwnd, 0
        SendMessageLong Toolbar, TB_BUTTONSTRUCTSIZE, LenB(BUT), 0

            'NB: set the bitmap size and the padding between the buttons.
            'play with these values to see the various size effects.
            SendMessageLong Toolbar, TB_SETBITMAPSIZE, 0, modDefinitions.MAKELONG(0, -2)
            SendMessageLong Toolbar, TB_SETPADDING, 0, modDefinitions.MAKELONG(10, 3)
            
            'add some strings to be used with the toolbar buttons
            SendMessage Toolbar, TB_ADDSTRING, 0, ByVal "&File"
            SendMessage Toolbar, TB_ADDSTRING, 0, ByVal "&Edit"
            SendMessage Toolbar, TB_ADDSTRING, 0, ByVal "&Window"
            SendMessage Toolbar, TB_ADDSTRING, 0, ByVal "&Help"
            
            'add the buttons along with the strings above.
            For I = 0 To 3
                BUTS(I).fsState = TBSTATE_ENABLED
                BUTS(I).fsStyle = TBSTYLE_BUTTON Or TBSTYLE_AUTOSIZE
                BUTS(I).iBitmap = -2  'make sure this is a non-positive value to show that we do not want to display an image
                BUTS(I).idCommand = CMDS(I)  'command identifiers for the buttons - see the array above within WM_CREATE.
                BUTS(I).iString = I
            Next

                'now add the buttons
                SendMessage Toolbar, TB_ADDBUTTONS, 4, BUTS(0)
                'set the font
                modDefinitions.SetFont Toolbar
        
                'setup pressed and non pressed states.
                TBPRESS.cbSize = LenB(TBPRESS)
                TBPRESS.dwMask = TBIF_STATE
                TBPRESS.fsState = TBSTATE_ENABLED Or TBSTATE_PRESSED 'Or TBSTATE_CHECKED

                TBUNPRESS.cbSize = LenB(TBUNPRESS)
                TBUNPRESS.dwMask = TBIF_STATE
                TBUNPRESS.fsState = TBSTATE_ENABLED
                '=================================END TOOLBAR=====================================================
                '
                '=================================REBAR===========================================================
                'add a rebar control to house the toolbar.
                Dim rbbi As REBARBANDINFO
                Dim btnSize As Long
                btnSize = SendMessage(Toolbar, TB_GETBUTTONSIZE, 0, 0)
                dwStyle = WS_CHILD Or WS_VISIBLE Or RBS_AUTOSIZE Or WS_BORDER Or CCS_NOPARENTALIGN Or CCS_NODIVIDER Or RBS_BANDBORDERS
                
                Coolbar = CreateWindowEx(WS_EX_TOOLWINDOW, "ReBarWindow32", "", _
                                                              dwStyle, 0, 0, 0, 0, hwnd, ICC_COOL_CLASSES, App.hInstance, ByVal 0&)

                'setup the properties of the band to add
                rbbi.cbSize = LenB(rbbi)
                rbbi.fMask = RBBIM_STYLE Or RBBIM_CHILD Or RBBIM_CHILDSIZE Or RBBIM_TEXT
                rbbi.fStyle = RBBS_GRIPPERALWAYS Or RBBS_CHILDEDGE 'adds a nice gripper effect.

                rbbi.cxMinChild = 0
                rbbi.cyMinChild = (modDefinitions.HIWORD(btnSize))  'the height based on a toolbar button size.
                rbbi.hWndCHild = Toolbar 'the toolbar windows handle.
                rbbi.lpText = ""

                'insert that band into the rebar control.
                SendMessage Coolbar, RB_INSERTBAND, 0, rbbi
                SetParent Coolbar, hwnd
                '=================================END REBAR===========================================================
                '
                '=================================STATUSBAR AND TEXTBOX===========================================
                'add a statusbar to the window
                dwStyle = WS_CHILD Or WS_VISIBLE Or CCS_BOTTOM Or SBARS_SIZEGRIP
                StatusBar = CreateWindowEx(WS_EX_TOOLWINDOW, "msctls_statusbar32", "", dwStyle, 0, 0, 0, 0, hwnd, ICC_COOL_CLASSES, App.hInstance, 0)
                modDefinitions.SetFont StatusBar
                'add an edit control to the window
                dwStyle = WS_CHILD Or WS_VSCROLL Or WS_HSCROLL Or ES_MULTILINE Or WS_VISIBLE
                TextBox = CreateWindowEx(WS_EX_CLIENTEDGE, "edit", "", dwStyle, 0, 0, 0, 0, hwnd, 0, App.hInstance, 0)
                modDefinitions.SetFont TextBox
                ProcTextBox = SetWindowLong(TextBox, GWL_WNDPROC, AddressOf WindowProcTextBox)
                SetFocusAPI TextBox
                '=================================END STATUSBAR AND TEXTBOX===========================================
                FormProc = 0

'helps us keep track of our accelerator keys when say ALT+f etc is pressed.
Case WM_SYSCHAR
                If (wParam = Asc("f") Or wParam = Asc("F")) Then
                    bckIndex = 0
                    SendMessage hwnd, WM_CANCELMODE, 0, 0
                    PostMessage hwnd, WM_COMMAND, IDM_M1 + bckIndex, 0
                ElseIf (wParam = Asc("e") Or wParam = Asc("'E")) Then
                    bckIndex = 1
                    SendMessage hwnd, WM_CANCELMODE, 0, 0
                    PostMessage hwnd, WM_COMMAND, IDM_M1 + bckIndex, 0
                ElseIf (wParam = Asc("w") Or wParam = Asc("W")) Then
                    bckIndex = 2
                    SendMessage hwnd, WM_CANCELMODE, 0, 0
                    PostMessage hwnd, WM_COMMAND, IDM_M1 + bckIndex, 0
                ElseIf (wParam = Asc("h") Or wParam = Asc("H")) Then
                    bckIndex = 3
                    SendMessage hwnd, WM_CANCELMODE, 0, 0
                    SendMessage hwnd, WM_COMMAND, IDM_M1 + bckIndex, 0
                End If
                
                'save the index of the selected button.
                index = bckIndex
                FormProc = 0
                Exit Function
                        
                           'used to detect accelerator keys after the alt key is pressed and the toolbar
Case WM_CHAR 'has the focus and a hot item enabled as this is the only time the window will accept keystrokes when
                           'it has the focus, the text control has the focus most of the time.
                    Hot = SendMessage(Toolbar, TB_GETHOTITEM, 0, 0)
                
                If ((wParam = Asc("f") Or wParam = Asc("F")) And Not (Hot = -1)) Then
                    bckIndex = 0
                    SendMessage hwnd, WM_CANCELMODE, 0, 0
                    PostMessage hwnd, WM_COMMAND, IDM_M1 + bckIndex, 0
                ElseIf ((wParam = Asc("e") Or wParam = Asc("'E")) And Not (Hot = -1)) Then
                    bckIndex = 1
                    SendMessage hwnd, WM_CANCELMODE, 0, 0
                    PostMessage hwnd, WM_COMMAND, IDM_M1 + bckIndex, 0
                ElseIf ((wParam = Asc("w") Or wParam = Asc("W")) And Not (Hot = -1)) Then
                    bckIndex = 2
                    SendMessage hwnd, WM_CANCELMODE, 0, 0
                    PostMessage hwnd, WM_COMMAND, IDM_M1 + bckIndex, 0
                ElseIf ((wParam = Asc("h") Or wParam = Asc("H")) And Not (Hot = -1)) Then
                    bckIndex = 3
                    SendMessage hwnd, WM_CANCELMODE, 0, 0
                    SendMessage hwnd, WM_COMMAND, IDM_M1 + bckIndex, 0
                End If
                
                index = bckIndex
                FormProc = 0
                Exit Function
                
    Case WM_COMMAND
            'see which button is being pressed.
            Select Case (modDefinitions.LOWORD(wParam))
                Case IDM_M1
                        'change that buttons state.
                        SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M1, TBPRESS
                        index = 0
                        'position the popup menu relative to the buttons position on the screen.
                        Dim bR As RECT
                        Dim pt As POINTAPI
                        SendMessage Toolbar, TB_GETRECT, IDM_M1, bR
                        pt.x = bR.Left
                        pt.y = bR.Bottom - bR.Top
                        ClientToScreen Toolbar, pt      'ClientToScreen hwnd, pt  .... is original code  I don't understand why...pt.x +12, pt.y+....
                        SendMessage StatusBar, SB_SETTEXT, 0, ByVal "Button1"
                        sys = False

                        'create the popup menu
                        hMenu = CreatePopupMenu()
                        'add some stuff to it.
                        AppendMenu hMenu, MF_STRING, IDM.IDM_FILE_EXIT, "E&xit   press X"
                        SetMenuDefaultItem hMenu, 0, True 'bold the first item as a test.
                        
                        'setup our hook to capture some messages when the popup box is displayed.
                        ProcKeyboardHook = SetWindowsHookEx(WH_MSGFILTER, AddressOf Filter, App.hInstance, GetCurrentThreadId())

                        'show the popup menu
                        'from some reason when pt.x is < 0 it does not put the menu on pt.x = 0, only problem with left side
                        TrackPopupMenu hMenu, TPM_LEFTALIGN, IIf(pt.x < 0, pt.x = 0, pt.x), pt.y, 0, hwnd, 0
                        
                        'remove the hook.
                        UnhookWindowsHookEx ProcKeyboardHook
                        ProcKeyboardHook = 0
                        'unpress this button after the hook procedure exits.
                        SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M1, TBUNPRESS
                        SetFocusAPI TextBox  'return the focus back to the edit control.
                        'the remaining commands perform the same actions as the case above
                        'see the above case for statement details.
                        SendMessage StatusBar, SB_SETTEXT, 0, ByVal ""  'no button pressed
                        Exit Function
            Case IDM_M2:
                        SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M2, TBPRESS
                        index = 1
                        SendMessage Toolbar, TB_GETRECT, IDM_M2, bR
                        pt.x = bR.Left
                        pt.y = bR.Top + bR.Bottom
                        ClientToScreen Toolbar, pt
                        SendMessage StatusBar, SB_SETTEXT, 0, ByVal "Button2"
                        sys = False
                        
                        hMenu = CreatePopupMenu()
                        AppendMenu hMenu, MF_STRING, IDM.IDM_EDIT_CUT, "&Cut"
                        AppendMenu hMenu, MF_STRING, IDM.IDM_EDIT_COPY, "C&opy"
                        AppendMenu hMenu, MF_STRING, IDM.IDM_EDIT_PASTE, "&Paste"
                        
                        ProcKeyboardHook = SetWindowsHookEx(WH_MSGFILTER, AddressOf Filter, App.hInstance, GetCurrentThreadId())

                        TrackPopupMenu hMenu, TPM_LEFTALIGN, IIf(pt.x < 0, pt.x = 0, pt.x), pt.y, 0, hwnd, 0
                        UnhookWindowsHookEx ProcKeyboardHook
                        ProcKeyboardHook = 0
                        SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M2, TBUNPRESS
                        SetFocusAPI TextBox
                        SendMessage StatusBar, SB_SETTEXT, 0, ByVal ""
                        Exit Function
            Case IDM_M3:
                        SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M3, TBPRESS
                        index = 2
                        SendMessage Toolbar, TB_GETRECT, IDM_M3, bR
                        pt.x = bR.Left
                        pt.y = bR.Bottom
                        ClientToScreen Toolbar, pt
                        SendMessage StatusBar, SB_SETTEXT, 0, ByVal "Button3"
                        sys = False
                        
                        hMenu = CreatePopupMenu()
                        AppendMenu hMenu, MF_STRING, IDM.IDM_WINDOW_CASCADE, "&Cascade"
                        AppendMenu hMenu, MF_STRING, IDM.IDM_WINDOW_TILE, "&Tile"
                        
                        ProcKeyboardHook = SetWindowsHookEx(WH_MSGFILTER, AddressOf Filter, App.hInstance, GetCurrentThreadId())

                        TrackPopupMenu hMenu, TPM_LEFTALIGN, IIf(pt.x < 0, pt.x = 0, pt.x), pt.y, 0, hwnd, 0
                        UnhookWindowsHookEx ProcKeyboardHook
                        ProcKeyboardHook = 0
                        SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M3, TBUNPRESS
                        SetFocusAPI TextBox
                        SendMessage StatusBar, SB_SETTEXT, 0, ByVal ""
                        Exit Function
            Case IDM_M4:
                        SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M4, TBPRESS
                        index = 3
                        SendMessage Toolbar, TB_GETRECT, IDM_M4, bR
                        pt.x = bR.Left
                        pt.y = bR.Bottom
                        ClientToScreen Toolbar, pt
                        SendMessage StatusBar, SB_SETTEXT, 0, ByVal "Button4"
                        sys = False
                        
                        hMenu = CreatePopupMenu()
                        AppendMenu hMenu, MF_STRING, IDM.IDM_HELP_ABOUT, "&About..."
                        
                        ProcKeyboardHook = SetWindowsHookEx(WH_MSGFILTER, AddressOf Filter, App.hInstance, GetCurrentThreadId())

                        TrackPopupMenu hMenu, TPM_LEFTALIGN, IIf(pt.x < 0, pt.x = 0, pt.x), pt.y, 0, hwnd, 0
                        UnhookWindowsHookEx ProcKeyboardHook
                        ProcKeyboardHook = 0
                        SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M4, TBUNPRESS
                        SetFocusAPI TextBox
                        SendMessage StatusBar, SB_SETTEXT, 0, ByVal ""
                        Exit Function
            Case IDM.IDM_FILE_EXIT
                        SendMessage StatusBar, SB_SETTEXT, 0, ByVal "Exit"
                        FormProc = 0
                        Exit Function
            Case IDM.IDM_EDIT_CUT
                        SendMessage StatusBar, SB_SETTEXT, 0, ByVal "Cut"
                        FormProc = 0
                        Exit Function
            Case IDM.IDM_EDIT_COPY
                        SendMessage StatusBar, SB_SETTEXT, 0, ByVal "Copy"
                        FormProc = 0
                        Exit Function
            Case IDM.IDM_EDIT_PASTE
                        SendMessage StatusBar, SB_SETTEXT, 0, ByVal "Paste"
                        FormProc = 0
                        Exit Function
            Case IDM.IDM_WINDOW_CASCADE
                        SendMessage StatusBar, SB_SETTEXT, 0, ByVal "Cascade"
                        FormProc = 0
                        Exit Function
            Case IDM.IDM_WINDOW_TILE
                        SendMessage StatusBar, SB_SETTEXT, 0, ByVal "Tile"
                        FormProc = 0
                        Exit Function
            Case IDM.IDM_HELP_ABOUT
                        SendMessage StatusBar, SB_SETTEXT, 0, ByVal "About"
                        FormProc = 0
            Case Else
                    FormProc = 0
                    Exit Function
            End Select
        
        Case WM_MENUSELECT         'ignore this.
                FormProc = 0
'            {
'                if (LOWORD(wparam) == IDM_FILE_EXIT)
'                {
'
'                }
'
'                return 0;
'            }
                
        Case WM_KEYDOWN
                Hot = SendMessage(Toolbar, TB_GETHOTITEM, 0, 0)
                If (Hot = -1) Then bckIndex = 0
                'perfomrs various ALT key press actions
                'press and unpress.
                'to see what i mean run the proggy and press the alt key
                'and then again to remove it.
                If (wParam = STE_MENU) Then 'if alt key was pressed
                        SetFocusAPI hwnd
                        If (Hot > -1) Then 'if a hot item exists then...
                            SendMessage Toolbar, TB_SETHOTITEM, -1, 0 'TURN THEM ALL OFF
                            SetFocusAPI TextBox
                            UpdateWindow Toolbar
                        Else
                            bckIndex = 0
                            Hot = 0
                            SendMessage Toolbar, TB_SETHOTITEM, -1, 0 'TURN THEM ALL OFF
                            SendMessage Toolbar, TB_SETHOTITEM, 0, 0
                        End If
                            FormProc = 0
                        Exit Function
                    End If
                
                'capture the left and right arrow keys
                If (wParam = VK_RIGHT And Hot > -1) Then
                        bckIndex = bckIndex + 1
                        If (bckIndex > 3) Then bckIndex = 0
            
                ElseIf (wParam = VK_LEFT And Hot > -1) Then
                        bckIndex = bckIndex - 1
                        If (bckIndex < 0) Then bckIndex = 3
                
                'capture the enter key
                'to test press alt and then the enter key (13)
                'this has the effect of selecting that item.
                ElseIf (wParam = 13 And Hot > -1) Then
                    index = bckIndex 'starting from hot item point for the keydown event in the hook procedure
                    SendMessage hwnd, WM_CANCELMODE, 0, 0
                    PostMessage hwnd, WM_COMMAND, IDM_M1 + bckIndex, 0
                    FormProc = 0
                    Exit Function
                
                Else
                    FormProc = 0
                   Exit Function
                End If
                
                index = bckIndex
                SendMessage Toolbar, TB_SETHOTITEM, bckIndex, 0
                Dim d As DRAWITEMSTRUCT
                d.CtlType = ODT_MENU
                d.hwndItem = hwnd
                d.itemState = ODS_SELECTED

                PostMessage hwnd, WM_DRAWITEM, 0, LenB(d)
                FormProc = 0
                Exit Function
Case WM_SIZE
                If Coolbar <> 0 And TextBox <> 0 And StatusBar <> 0 And Toolbar <> 0 Then
                'play with the controls positions here to make them look neat on the window.
                Dim parts(2) As Long
                parts(0) = modDefinitions.LOWORD(lParam) / 2
                parts(1) = -1
                LockWindowUpdate GetDesktopWindow
                SendMessage StatusBar, SB_SETPARTS, 2, parts(0)
                SendMessage StatusBar, SB_SETTEXT, 0, ByVal "Ready"
                SendMessage StatusBar, WM_SIZE, 0, 0
                
                SendMessage Toolbar, TB_AUTOSIZE, 0, 0
                
                MoveWindow Coolbar, 0, 0, modDefinitions.LOWORD(lParam), modDefinitions.HIWORD(lParam), True

                'size of the statusbar
                Dim sr As RECT
                Dim sH As Long, tH As Long, newH As Long
                GetClientRect StatusBar, sr
                sH = sr.Bottom - sr.Top
                'size of the coolbar
                Dim tr As RECT
                GetClientRect Coolbar, tr
                tH = tr.Bottom - tr.Top
                newH = (modDefinitions.HIWORD(lParam) - sH) - tH

                'reposition the edit control correclty
                SetWindowPos TextBox, 0, 0, tH, modDefinitions.LOWORD(lParam), newH, SWP_NOZORDER
                LockWindowUpdate 0
                FormProc = 0
                Exit Function
                End If

Case WM_SYSCOMMAND
                If (wParam = SC_KEYMENU) Then
                    'give the main window the focus and send it a key message to enable the menu items
                    SendMessage hwnd, WM_KEYDOWN, STE_MENU, 0
                    FormProc = -1
                    Exit Function
                End If
                sys = True
                FormProc = -1
             
Case WM_NOTIFY 'used by the common controls to signal to the main window that it is executing an action.
                CopyMemory NMH, ByVal lParam, Len(NMH)
                If NMH.hwndFrom = Toolbar Then 'we need only messages for toolbar
                        Select Case NMH.code
                            Case NM_CUSTOMDRAW
                                If IsXp = True Then
                                    FormProc = XpStyle(NMH.hwndFrom, lParam)
                                    Exit Function
                                End If
                        End Select
                End If
'no need for this code ===============================================
 '       {
 '           LPNMHDR nm = (LPNMHDR)lparam;
 '           LPNMTOOLBAR nmTB = (LPNMTOOLBAR)lparam;
'
'            switch(nm->code)
'            {
'                case RBN_BEGINDRAG: //cancel any rebar gripper drag operations
'                {
'                    return -1;
'                }
'
'                break;
'            }
'
'            return 0;
'        }
'=================================================================
Case WM_DESTROY
    'just in case something gone wrong
    If KeyboardHook <> 0 Then UnhookWindowsHookEx KeyboardHook
    If ProcKeyboardHook <> 0 Then UnhookWindowsHookEx ProcKeyboardHook
    'destroy created windows
    If TextBox <> 0 Then DestroyWindow TextBox
    If StatusBar <> 0 Then DestroyWindow StatusBar
    If Toolbar <> 0 Then DestroyWindow Toolbar
    If Coolbar <> 0 Then DestroyWindow Coolbar
    'send the WM_QUIT message to the form
    PostQuitMessage 0 '&
    FormProc = 0
End Select
FormProc = DefWindowProc(hwnd, iMsg, wParam, lParam)
End Function

Private Function XpStyle(ByVal hwnd As Long, ByVal lParam As Long) As Long
Dim BBRUSH As Long, HHDC As Long, HHEIGHT As Long, HFONT As Long, HFONTOLD As Long
Dim NMTBCUST As NMTBCUSTOMDRAW
Dim strText As String
Dim IsHot As Boolean, IsSelected As Boolean, isFlat As Boolean
Dim RC As RECT

   CopyMemory NMTBCUST, ByVal lParam, Len(NMTBCUST)
   
        Select Case NMTBCUST.nmcd.dwDrawStage
                        Case CDDS_PREPAINT
                                XpStyle = CDRF_NOTIFYITEMDRAW
                        Case CDDS_ITEMPREPAINT
                                SystemParametersInfo SPI_GETFLATMENU, 0&, isFlat, 0&
                                    If isFlat Then
                                                HHDC = NMTBCUST.nmcd.hdc
                                                LSet RC = NMTBCUST.nmcd.RC
                                        
                                                IsHot = ((NMTBCUST.nmcd.uItemState And CDIS_HOT) = CDIS_HOT)
                                                IsSelected = ((NMTBCUST.nmcd.uItemState And CDIS_SELECTED) = CDIS_SELECTED)
                                                
                                                If IsHot Or IsSelected Then XpStyle = CDRF_SKIPDEFAULT
                                                    
                                                    If IsSelected Or IsHot Then
                                                            strText = String$(255, Chr(0))
                                                            SendMessage Toolbar, TB_GETBUTTONTEXT, NMTBCUST.nmcd.dwItemSpec, ByVal strText
                                                            strText = Replace(strText, Chr(0), "", 1, , vbTextCompare)
                                                            SetBkMode HHDC, TRANSPARENT
                                                            BBRUSH = CreateSolidBrush(ToolbarMenuColor)
                                                            FillRect HHDC, RC, BBRUSH
                                                            SetTextColor HHDC, TranslateColor(vbHighlightText)
                                                            HHEIGHT = -MulDiv(8, GetDeviceCaps(HHDC, LOGPIXELSY), 72)
                                                            HFONT = CreateFont(HHEIGHT, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "Trebuchet MS")
                                                            HFONTOLD = SelectObject(HHDC, HFONT)
                                                            RC.Top = RC.Top + 1
                                                            DrawText HHDC, strText, -1, RC, DT_CENTER Or DT_SINGLELINE
                                                            SelectObject HHDC, HFONTOLD
                                                            DeleteObject BBRUSH
                                                    End If
                            End If
        End Select
End Function

Private Function WindowProcTextBox(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'dedicated to messages sent to "b" CommandButton
Select Case iMsg
    Case WM_KEYDOWN
            If (wParam = STE_MENU) Then
                'give the main window the focus and send it an aly key message
                SendMessage Main_Form, WM_KEYDOWN, STE_MENU, lParam
                SetFocusAPI Main_Form
            End If
    Case WM_SYSCHAR 'handle alt and keydown event and pass it to the main windows procedure.
            SendMessage Main_Form, WM_SYSCHAR, wParam, lParam
    Case WM_SETFOCUS
            SendMessage Toolbar, TB_SETHOTITEM, -1, 0
    End Select
WindowProcTextBox = CallWindowProc(ProcTextBox, hwnd, iMsg, wParam, lParam)
End Function

'this filter function captures a couple of messages while the
'popup menu is being displayed the reason for this is that a popup menu is modal.
Private Function Filter(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim stat1 As Long, stat2 As Long, stat3 As Long, stat4 As Long
Dim r1 As RECT, r2 As RECT, r3 As RECT, r4 As RECT
Dim pt As POINTAPI
Dim MSGA As msg
   
    CopyMemory MSGA, ByVal lParam, 12

    If (MSGA.message = WM_KEYDOWN) Then
                    If (MSGA.wParam = VK_RIGHT) Then
                            index = index + 1
                            If (index > 3) Then index = 0
                       
                    ElseIf (MSGA.wParam = VK_LEFT) Then
                            index = index - 1
                            If (index < 0) Then index = 3
                        
                    ElseIf (MSGA.wParam = STE_MENU) Then
                            PostMessage Main_Form, WM_CANCELMODE, 0, 0
                            Filter = 0
                            Exit Function
                        Else
                            bckIndex = index
                            Filter = 0
                            Exit Function
                    End If
                    
                    SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M1 + index, TBPRESS
                    SendMessage Main_Form, WM_CANCELMODE, 0, 0
                    PostMessage Main_Form, WM_COMMAND, IDM_M1 + index, 0

                    bckIndex = index
                    SendMessage Toolbar, TB_SETHOTITEM, index, 0
                    Filter = 0
                    Exit Function

    ElseIf (MSGA.message = WM_LBUTTONDOWN) Then
                'check to see if the system menu is being displayed if so...
                If (sys = True) Then
                    KeyboardHook = 0
                    Filter = CallNextHookEx(KeyboardHook, nCode, wParam, lParam)
                    Exit Function
                End If
        
                'these if statements check the mouse cursor position to
                'see if it lies within a toolbars button
                'and if so displays the popup etc.
                SendMessage Toolbar, TB_GETRECT, IDM_M1, r1
                SendMessage Toolbar, TB_GETRECT, IDM_M2, r2
                SendMessage Toolbar, TB_GETRECT, IDM_M3, r3
                SendMessage Toolbar, TB_GETRECT, IDM_M4, r4
        
                GetCursorPos pt
                ScreenToClient Toolbar, pt

                    If (PtInRect(r1, pt.x, pt.y)) Then
                        stat1 = SendMessage(Toolbar, TB_ISBUTTONPRESSED, IDM_M1, 0)
                        If Not (stat1 = TBSTATE_PRESSED) Then
                            SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M1, TBUNPRESS
                            PostMessage Main_Form, WM_CANCELMODE, 0, 0
                            PostMessage Main_Form, WM_COMMAND, IDM_M1, 0
                            Filter = True
                            Exit Function
                        End If
            
                        SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M1, TBPRESS
                        PostMessage Main_Form, WM_CANCELMODE, 0, 0
                        Filter = True
                        Exit Function
                   
                    ElseIf (PtInRect(r2, pt.x, pt.y)) Then
                        stat2 = SendMessage(Toolbar, TB_ISBUTTONPRESSED, IDM_M2, 0)
                        If Not (stat2 = TBSTATE_PRESSED) Then
                            SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M2, TBUNPRESS
                            PostMessage Main_Form, WM_CANCELMODE, 0, 0
                            PostMessage Main_Form, WM_COMMAND, IDM_M2, 0
                            Filter = True
                            Exit Function
                       End If
                        
                        SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M2, TBPRESS
                        PostMessage Main_Form, WM_CANCELMODE, 0, 0
                        Filter = True
                        Exit Function
                        
                    ElseIf (PtInRect(r3, pt.x, pt.y)) Then
                        stat3 = SendMessage(Toolbar, TB_ISBUTTONPRESSED, IDM_M3, 0)
                        If Not (stat3 = TBSTATE_PRESSED) Then
                            SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M3, TBUNPRESS
                            PostMessage Main_Form, WM_CANCELMODE, 0, 0
                            PostMessage Main_Form, WM_COMMAND, IDM_M3, 0
                            Filter = True
                            Exit Function
                        End If
            
                        SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M3, TBPRESS
                        PostMessage Main_Form, WM_CANCELMODE, 0, 0
                        Filter = True
                        Exit Function
                    
                    ElseIf (PtInRect(r4, pt.x, pt.y)) Then
                        'solves click while hot item has changed to another button with arrow keys
                        'and the mouse is not moving and then the item the mouse is on is reclicked without
                        'mouse moving.
                        stat4 = SendMessage(Toolbar, TB_ISBUTTONPRESSED, IDM_M4, 0)
                        If Not (stat4 = TBSTATE_PRESSED) Then
                            SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M4, TBUNPRESS
                            PostMessage Main_Form, WM_CANCELMODE, 0, 0
                            PostMessage Main_Form, WM_COMMAND, IDM_M4, 0
                            Filter = True
                            Exit Function
                        End If
            
                        SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M4, TBPRESS
                        PostMessage Main_Form, WM_CANCELMODE, 0, 0
                        Filter = True
                        Exit Function
                    End If
                    
    ElseIf (MSGA.message = WM_MOUSEMOVE) Then
                        'prevents mousemove from running if the left or right keys are being
                        'pressed, avoids flickering which occurs
                
                        'to test remove this if statement and press a menubar button
                        'keep the mouse within that button and try to use the left and right arrow keys.
                        If (GetKeyState(VK_RIGHT) < 0 Or GetKeyState(VK_LEFT) < 0) Then Exit Function
                        
                        If (sys = True) Then 'check for system menu.
                            KeyboardHook = 0
                            Filter = CallNextHookEx(KeyboardHook, nCode, wParam, lParam)
                            Exit Function
                        End If
                        
                        SendMessage Toolbar, TB_GETRECT, IDM_M1, r1
                        SendMessage Toolbar, TB_GETRECT, IDM_M2, r2
                        SendMessage Toolbar, TB_GETRECT, IDM_M3, r3
                        SendMessage Toolbar, TB_GETRECT, IDM_M4, r4
                
                        GetCursorPos pt
                        ScreenToClient Toolbar, pt

                        If (PtInRect(r1, pt.x, pt.y)) Then
                                index = 0
                                stat1 = SendMessage(Toolbar, TB_GETSTATE, IDM_M1, 0)
                                
                                'if this button is pressed already exit the hook proc.
                                If ((stat1 And (TBSTATE_PRESSED))) Then
                                    KeyboardHook = 0
                                    Filter = CallNextHookEx(KeyboardHook, nCode, wParam, lParam)
                                    Exit Function
                                End If
                                
                                SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M1, TBPRESS
                                SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M2, TBUNPRESS
                                SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M3, TBUNPRESS
                                SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M4, TBUNPRESS
                    
                                'the reason we use postmessage here is so we can execute the next statement immediately
                                'unlike sendmessage which whats until the message is processed before returning
                                SendMessage Main_Form, WM_CANCELMODE, 0, 0
                                PostMessage Main_Form, WM_COMMAND, IDM_M1, 0
                
                        ElseIf (PtInRect(r2, pt.x, pt.y)) Then
                                stat2 = SendMessage(Toolbar, TB_GETSTATE, IDM_M2, 0)
                                index = 1
                                
                                If ((stat2 And (TBSTATE_PRESSED))) Then
                                    KeyboardHook = 0
                                    Filter = CallNextHookEx(KeyboardHook, nCode, wParam, lParam)
                                    Exit Function
                                End If
                                
                                SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M1, TBUNPRESS
                                SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M2, TBPRESS
                                SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M3, TBUNPRESS
                                SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M4, TBUNPRESS
                    
                                SendMessage Main_Form, WM_CANCELMODE, 0, 0
                                PostMessage Main_Form, WM_COMMAND, IDM_M2, 0
                            
                        ElseIf (PtInRect(r3, pt.x, pt.y)) Then
                                stat3 = SendMessage(Toolbar, TB_GETSTATE, IDM_M3, 0)
                                index = 2
                                
                                If ((stat3 And (TBSTATE_PRESSED))) Then
                                    KeyboardHook = 0
                                    Filter = CallNextHookEx(KeyboardHook, nCode, wParam, lParam)
                                    Exit Function
                                End If
                                
                                SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M1, TBUNPRESS
                                SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M2, TBUNPRESS
                                SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M3, TBPRESS
                                SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M4, TBUNPRESS
                    
                                SendMessage Main_Form, WM_CANCELMODE, 0, 0
                                PostMessage Main_Form, WM_COMMAND, IDM_M3, 0
                            
                        ElseIf (PtInRect(r4, pt.x, pt.y)) Then
                                stat4 = SendMessage(Toolbar, TB_GETSTATE, IDM_M4, 0)
                                index = 3
            
                                If ((stat4 And (TBSTATE_PRESSED))) Then
                                    KeyboardHook = 0
                                    Filter = CallNextHookEx(KeyboardHook, nCode, wParam, lParam)
                                    Exit Function
                                End If
                                
                                SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M1, TBUNPRESS
                                SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M2, TBUNPRESS
                                SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M3, TBUNPRESS
                                SendMessage Toolbar, TB_SETBUTTONINFO, IDM_M4, TBPRESS
                    
                                SendMessage Main_Form, WM_CANCELMODE, 0, 0
                                PostMessage Main_Form, WM_COMMAND, IDM_M4, 0
                    
                        End If

    End If
    SendMessage Toolbar, TB_SETHOTITEM, -1, 0
    KeyboardHook = 0
    Filter = CallNextHookEx(KeyboardHook, nCode, wParam, lParam)
End Function
