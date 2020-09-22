Attribute VB_Name = "modDefinitions"
Option Explicit
DefLng A-Z
Public Type CommonControlsEx
        dwSize As Long '// size of this structure
        dwICC As Long  '// flags indicating which classes to be initialized
End Type
 Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Public Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type
Public Type tagMSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As String
    pt As POINTAPI
End Type
Public Type WNDCLASS
   ' cbSize As Long 'Win 3.x
    Style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
    'hIconSm As Long 'Win 4.0
End Type
Public Type TBBUTTON
   iBitmap As Long
   idCommand As Long
   fsState As Byte
   fsStyle As Byte
   bReserved1 As Byte
   bReserved2 As Byte
   dwData As Long
   iString As Long
End Type
Public Type TBBUTTONINFO
   cbSize As Long
   dwMask As Long
   idCommand As Long
   iImage As Long
   fsState As Byte
   fsStyle As Byte
   cx As Integer
   lParam As Long
   pszText As Long
   cchText As Long
End Type
Public Type REBARBANDINFO
    cbSize As Long
    fMask As Long
    fStyle As Long
    clrFore As Long
    clrBack As Long
    lpText As String
    cch As Long
    iImage As Long
    hWndCHild As Long
    cxMinChild As Long
    cyMinChild As Long
    cx As Long
    hbmBack As Long
    wID As Long
End Type
Public Type NMCUSTOMDRAW
   hdr As NMHDR
   dwDrawStage As Long
   hdc As Long
   RC As RECT
   dwItemSpec As Long
   uItemState As Long
   lItemlParam As Long
End Type
Public Type NMTBCUSTOMDRAW
   nmcd As NMCUSTOMDRAW
   hbrMonoDither As Long
   hbrLines As Long
   hpenLines As Long
   clrText As Long
   clrMark As Long
   clrTextHighlight As Long
   clrBtnFace As Long
   clrBtnHighlight As Long
   clrHighlightHotTrack As Long
   rcText As RECT
   nStringBkMode As Long
   nHLStringBkMode As Long
End Type
Public Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As RECT
    ItemData As Long
End Type
Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long 'Structure size = 148
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformID As Long
        szCSDVersion As String * 128
End Type

Public Const WS_EX_TOOLWINDOW = &H80&
Public Const WM_USER = &H400
Public Const WM_CREATE = &H1
Public Const WM_SIZE = &H5
Public Const WM_NOTIFY = &H4E
Public Const GWL_WNDPROC = (-4)

'generic window (extended) styles:
Public Const WS_VISIBLE = &H10000000
Public Const WS_CHILD = &H40000000
Public Const WS_THICKFRAME = &H40000
Public Const WS_TABSTOP = &H10000
Public Const WS_BORDER = &H800000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_CAPTION = &HC00000 ' WS_BORDER Or WS_DLGFRAME
Public Const WS_SYSMENU = &H80000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
'uncomment the flags in the following constant to have Maximize and Minimize buttons
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_WINDOWEDGE = &H100&
Public Const WS_EX_STATICEDGE = &H20000
'constants used in creating the TextBox:
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000

Public Const ES_MULTILINE = &H4& 'textbox

Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_VSCROLL = &H115 'vertical scroll
Public Const WM_KEYUP = &H101 'emulate the end of _KeyPress or _Change
Public Const WM_LBUTTONUP = &H202 'emulate the end of _Click
Public Const WM_LBUTTONDOWN = &H201 'emulate the beginning of _Click
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_PARENTNOTIFY = &H210
Public Const WM_SHOWWINDOW = &H18
Public Const WM_DESTROY = &H2 'aka Form_Unload
Public Const WM_SETFONT = &H30 'used in building text font for the new controls
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302

Public Const SM_CXSCREEN = 0&
Public Const SM_CYSCREEN = 1&

Public Const CS_VREDRAW = &H1
Public Const CS_HREDRAW = &H2
Public Const COLOR_WINDOW = 5
Public Const IDC_ARROW = 32512&
Public Const IDI_APPLICATION = 32512&
Public Const SW_SHOWNORMAL = 1
Public Const CW_USEDEFAULT = &H80000000

Public Const ICC_BAR_CLASSES = &H4&      '// toolbar, statusbar, trackbar, tooltips
Public Const ICC_COOL_CLASSES = &H400&   '// rebar (coolbar) control

Public Const STE_MENU = 901
Public Const COLOR_BTNFACE = 15
Public Const SWP_NOSIZE = &H1

Public Const TB_PRESSBUTTON = (WM_USER + 3)
Public Const TB_ISBUTTONPRESSED = (WM_USER + 11)
Public Const TB_ISBUTTONHIGHLIGHTED = (WM_USER + 14)
Public Const TB_ADDBUTTONS = (WM_USER + 20)
Public Const TB_ADDSTRING = (WM_USER + 28)
Public Const TB_BUTTONSTRUCTSIZE = (WM_USER + 30)
Public Const TB_SETBUTTONSIZE = (WM_USER + 31)
Public Const TB_SETBITMAPSIZE = (WM_USER + 32)
Public Const TB_AUTOSIZE = (WM_USER + 33)
Public Const TB_SETPARENT = (WM_USER + 37)
Public Const TB_GETBUTTONTEXT = (WM_USER + 45)
Public Const TB_GETRECT = (WM_USER + 51)
Public Const TB_GETBUTTONSIZE = (WM_USER + 58)
Public Const TB_GETBUTTONINFO = (WM_USER + 65)
Public Const TB_SETBUTTONINFO = (WM_USER + 66)
Public Const TB_GETHOTITEM = (WM_USER + 71)
Public Const TB_SETHOTITEM = (WM_USER + 72)

Public Const TB_SETPADDING = (WM_USER + 87)

Public Const TBSTYLE_BUTTON = &H0
Public Const TBSTYLE_FLAT = &H800
Public Const TBSTYLE_AUTOSIZE = &H10         '// automatically calculate the cx of the button

Public Const TBIF_STYLE = &H8&
Public Const TBIF_STATE = &H4&
Public Const TB_SETSTATE = (WM_USER + 17)
Public Const TB_GETSTATE = (WM_USER + 18)
Public Const TBSTATE_CHECKED = &H1
Public Const TBSTATE_PRESSED = &H2
Public Const TBSTATE_ENABLED = &H4


Public Const RBS_AUTOSIZE = &H2000&
Public Const RBS_BANDBORDERS = &H400
Public Const RBBIM_STYLE = &H1
Public Const RBBIM_TEXT = &H4
Public Const RBBIM_CHILD = &H10
Public Const RBBIM_CHILDSIZE = &H20
Public Const RBBS_CHILDEDGE = &H4  '// edge around top & bottom of child window
Public Const RBBS_GRIPPERALWAYS = &H80      ' always show the gripper
Public Const RB_INSERTBAND = (WM_USER + 1)
Public Const RB_GETBARHEIGHT = (WM_USER + 27)


Public Const SB_SETTEXT = (WM_USER + 1)
Public Const SB_SETPARTS = (WM_USER + 4)

Public Const WM_CHAR = &H102
Public Const WM_SYSCHAR = &H106
Public Const WM_COMMAND = &H111
Public Const WM_SYSCOMMAND = &H112
Public Const WM_CANCELMODE = &H1F
Public Const WM_SETFOCUS = &H7
Public Const WM_MOUSEMOVE = &H200
Public Const WM_MENUSELECT = &H11F
Public Const SC_KEYMENU = &HF100&
Public Const WM_KEYDOWN = &H100

'//====== COMMON CONTROL STYLES ===========
Public Const CCS_BOTTOM = &H3
Public Const CCS_NORESIZE = &H4
Public Const CCS_NOPARENTALIGN = &H8
Public Const CCS_NODIVIDER = &H40

Public Const SBARS_SIZEGRIP = &H100

Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOZORDER = &H4

Public Const MF_STRING = &H0&
Public Const MF_HILITE = &H80&
Public Const MF_BYCOMMAND = &H0&

Public Const WH_MSGFILTER As Long = (-1)
Public Const TPM_LEFTALIGN = &H0&
Public Const VK_LEFT = &H25
Public Const VK_RIGHT = &H27

Public Const SPI_GETFLATMENU = &H1022

Public Const CDDS_PREPAINT = &H1
Public Const CDDS_ITEM = &H10000
Public Const CDDS_ITEMPREPAINT = (CDDS_ITEM Or CDDS_PREPAINT)
Public Const CDRF_SKIPDEFAULT = &H4
Public Const CDRF_NOTIFYITEMDRAW = &H20

Public Const CDIS_SELECTED = &H1
Public Const CDIS_HOT = &H40

Public Const LOGPIXELSY = 90
Public Const NM_FIRST = &HFFFF + 1
Public Const NM_CUSTOMDRAW = (NM_FIRST - 12)
Public Const COLOR_HIGHLIGHT = 13
Public Const TRANSPARENT = 1
Public Const DT_CENTER = &H1
Public Const DT_SINGLELINE = &H20
Public Const ODS_SELECTED = &H1
Public Const ODT_MENU = 1
Public Const WM_DRAWITEM = &H2B
Public Const TPM_RETURNCMD = &H100

Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As msg) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Public Declare Function HiliteMenuItem Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal wIDHiliteItem As Long, ByVal wHilite As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpFn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As Any) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SetParent Lib "user32" (ByVal hWndCHild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function InitCommonControlsEx Lib "COMCTL32" (LPINITCOMMONCONTROLSEX As CommonControlsEx) As Boolean
Public Declare Sub InitCommonControls Lib "COMCTL32" ()
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Private Const ClassName = "Stex"
Private Const ApplicationCaption = "Win32 MenuBar"
Public IsXp As Boolean
Dim tmpMainForm As Long

Sub Main()
Dim wMsg As msg
Call LoadCommCtls        'There is also Manifest file in the RES - full XP support
Call GetOSVersion         'this is for our menu, flat or 3d (if XP)
App.TaskVisible = False 'vb auto creates task so we need to hide it

'Register window class name
If (Not RegisterWindowClass) Then
    MsgBox "Oooops..." & vbCrLf & "Window not created.", vbExclamation + vbOKOnly, "Error"
    Exit Sub
Else
        If (Not CreateWindow) Then
            MsgBox "Oooops..." & vbCrLf & "Class not registered.", vbExclamation + vbOKOnly, "Error"
            Exit Sub
        End If
End If

'Loop will exit when WM_QUIT is sent to the window.
Do While GetMessage(wMsg, 0, 0, 0)
''TranslateMessage takes keyboard messages and converts them to WM_CHAR for easier processing
TranslateMessage wMsg
''Dispatchmessage calls the default window procedure to process the window message. (WndProc)
DispatchMessage wMsg
Loop
Call UnregisterClass(ClassName, App.hInstance)
End Sub

Private Function RegisterWindowClass() As Boolean
Dim wc As WNDCLASS
'Registers the new window so we can use its class name:
'redraw entire window when movement or size adjustments modify the client area's width/height:
With wc
    .Style = 0 'CS_HREDRAW Or CS_VREDRAW  'Specifies the class style(s). Styles can be combined by using the bitwise OR operator
    .lpfnwndproc = GetAddress(AddressOf FormProc) 'Address of (pointer to) the window procedure
    .hInstance = App.hInstance 'Handle to the instance that the window procedure of this class is within
    .hIcon = LoadIcon(0&, IDI_APPLICATION) 'Default application icon
    .hCursor = LoadCursor(0&, IDC_ARROW) 'Default arrow cursor
    .hbrBackground = GetSysColorBrush(COLOR_BTNFACE) 'Default color for window background
    .lpszClassName = ClassName 'Pointer to a null-terminated string or, if lpszClassName is a string, it specifies the window class name
End With
RegisterWindowClass = RegisterClass(wc) <> 0
End Function

Private Function CreateWindow() As Boolean
'Create actual window
tmpMainForm = CreateWindowEx(0, ClassName, ApplicationCaption, WS_OVERLAPPEDWINDOW, 0, 0, 500, 400, 0, 0, App.hInstance, ByVal 0&)
ShowWindow tmpMainForm, SW_SHOWNORMAL
CreateWindow = (tmpMainForm <> 0)
End Function

Private Function GetAddress(ByVal lngAddr As Long) As Long
'Used with AddressOf to return the address in memory of a procedure
GetAddress = lngAddr '&
End Function

Private Function LoadCommCtls() As Boolean
Dim ctEx As CommonControlsEx
On Error GoTo Hell
  ctEx.dwSize = LenB(ctEx)
  ctEx.dwICC = ICC_BAR_CLASSES Or ICC_COOL_CLASSES
  LoadCommCtls = InitCommonControlsEx(ctEx)
  If LoadCommCtls = False Then Call InitCommonControls
  Exit Function
Hell:
  Call InitCommonControls
End Function

Private Sub GetOSVersion()
Dim OS As OSVERSIONINFO
Dim mRes As Long

     OS.dwOSVersionInfoSize = Len(OS)
     mRes = GetVersionEx(OS)
        
        Select Case OS.dwMajorVersion
            Case 5
                Select Case OS.dwMinorVersion
                    Case Is >= 1 'Windows XP
                        IsXp = True
                    End Select
            End Select
End Sub
Public Sub SetFont(ByVal hwnd As Long, Optional ByVal fontName As String = "Trebuchet MS", Optional ByVal fontSize As Long = 8)
Dim HHDC As Long, HFONT As Long, HHEIGHT As Long
If fontSize < 0 Then fontSize = 8
If fontName = "" Then fontName = "Trebuchet MS"
    HHDC = GetDC(hwnd)
    HHEIGHT = -MulDiv(fontSize, GetDeviceCaps(HHDC, LOGPIXELSY), 72)
    HFONT = CreateFont(HHEIGHT, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, fontName)
    SendMessage hwnd, WM_SETFONT, HFONT, MAKELONG(True, 0)
    ReleaseDC hwnd, HHDC
End Sub

Public Function HIWORD(dwValue As Long) As Integer ' Returns the low 16-bit integer from a 32-bit long integer
    MoveMemory HIWORD, ByVal VarPtr(dwValue) + 2, 2
End Function

Public Function LOWORD(dwValue As Long) As Integer ' Returns the low 16-bit integer from a 32-bit long integer
  MoveMemory LOWORD, dwValue, 2
End Function

Public Function MAKELONG(wLow As Long, wHigh As Long) As Long ' Combines two integers into a long integer
  MAKELONG = wLow
  MoveMemory ByVal VarPtr(MAKELONG) + 2, wHigh, 2
End Function

'thanks to www.vbaccelerator.com====================================================================================
Public Function TranslateColor(ByVal oClr As Long, Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then TranslateColor = -1
End Function

Private Property Get BlendColor(ByVal oColorFrom As Long, ByVal oColorTo As Long, Optional ByVal alpha As Long = 128) As Long
Dim lCFrom As Long
Dim lCTo As Long
   lCFrom = TranslateColor(oColorFrom)
   lCTo = TranslateColor(oColorTo)
Dim lSrcR As Long
Dim lSrcG As Long
Dim lSrcB As Long
Dim lDstR As Long
Dim lDstG As Long
Dim lDstB As Long
   lSrcR = lCFrom And &HFF
   lSrcG = (lCFrom And &HFF00&) \ &H100&
   lSrcB = (lCFrom And &HFF0000) \ &H10000
   lDstR = lCTo And &HFF
   lDstG = (lCTo And &HFF00&) \ &H100&
   lDstB = (lCTo And &HFF0000) \ &H10000
     
   
   BlendColor = RGB(((lSrcR * alpha) / 255) + ((lDstR * (255 - alpha)) / 255), ((lSrcG * alpha) / 255) + ((lDstG * (255 - alpha)) / 255), ((lSrcB * alpha) / 255) + ((lDstB * (255 - alpha)) / 255))
      
End Property
Public Property Get ToolbarMenuColor() As Long
   ToolbarMenuColor = BlendColor(vbHighlight, vbHighlight, 0)
End Property
'=============================================================================================================
