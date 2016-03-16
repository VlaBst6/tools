Attribute VB_Name = "API1"
Option Explicit
'bmp dc stuff
'Public Declare Function GetDesktopWindow Lib "user32" () As Long
'Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
'Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
'Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
'Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
'Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
'Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
'Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
'Public Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
'Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
'Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'Public Declare Function PolyDraw Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, lpbTypes As Byte, ByVal cCount As Long) As Long
'Public Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
'Public Declare Function ArcTo Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long

Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Public Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long

'Regens
'Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
'Public Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
'Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
'Public Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
'Public Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
'Public Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
'Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'Public Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
'Public Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Public Declare Function PaintRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
'Public Declare Function RectInRegion Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
'Public Declare Function GetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
'    'this alway returns error but it works
    '(create compatible regen of desktop then combind the one u get from here)
    

'more window functions
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Public Declare Function GetWinMetaFileBits Lib "gdi32" (ByVal hemf As Long, ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal fnMapMode As Long, ByVal hdcRef As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Public Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function GetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Public Declare Function GetWindowExtEx Lib "gdi32" (ByVal hdc As Long, lpSize As Size) As Long
Public Declare Function GetWindowContextHelpId Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1
'Public Const SW_NORMAL = 1
'Public Const SW_PARENTCLOSING = 1
'Public Const SW_OTHERZOOM = 2
'Public Const SW_INVALIDATE = &H2
'Public Const SW_OTHERUNZOOM = 4
'Public Const SW_ERASE = &H4
'Public Const SW_MAX = 10
'Public Const SW_PARENTOPENING = 3
'Public Const SW_SCROLLCHILDREN = &H1
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowsHook Lib "user32" Alias "SetWindowsHookA" (ByVal nFilterType As Long, ByVal pfnFilterProc As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

'mouse
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


'for ontop
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const NULLREGION = 1
Public Const SIMPLEREGION = 2
Public Const COMPLEXREGION = 3



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

Public Enum eCombineMode
    RGN_AND = 1
    RGN_OR = 2
    RGN_XOR = 3
    RGN_DIFF = 4
    RGN_COPY = 5
    RGN_MAX = RGN_COPY
    RGN_MIN = RGN_AND
    RGN_SetTo = -1
End Enum








Public Const WM_USER = &H400

Public Const WM_ACTIVATE = &H6
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_ASKCBFORMATNAME = &H30C
Public Const WM_CANCELJOURNAL = &H4B
Public Const WM_CANCELMODE = &H1F
Public Const WM_CHANGECBCHAIN = &H30D
Public Const WM_CHAR = &H102
Public Const WM_CHARTOITEM = &H2F
Public Const WM_CHILDACTIVATE = &H22
Public Const WM_CHOOSEFONT_GETLOGFONT = (WM_USER + 1)
Public Const WM_CHOOSEFONT_SETFLAGS = (WM_USER + 102)
Public Const WM_CHOOSEFONT_SETLOGFONT = (WM_USER + 101)
Public Const WM_CLEAR = &H303
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_COMPACTING = &H41
Public Const WM_COMPAREITEM = &H39
Public Const WM_CONVERTREQUESTEX = &H108
Public Const WM_COPY = &H301
Public Const WM_COPYDATA = &H4A
Public Const WM_CREATE = &H1
Public Const WM_CTLCOLORBTN = &H135
Public Const WM_CTLCOLORDLG = &H136
Public Const WM_CTLCOLOREDIT = &H133
Public Const WM_CTLCOLORLISTBOX = &H134
Public Const WM_CTLCOLORMSGBOX = &H132
Public Const WM_CTLCOLORSCROLLBAR = &H137
Public Const WM_CTLCOLORSTATIC = &H138
Public Const WM_CUT = &H300
Public Const WM_DDE_FIRST = &H3E0
Public Const WM_DDE_ACK = (WM_DDE_FIRST + 4)
Public Const WM_DDE_ADVISE = (WM_DDE_FIRST + 2)
Public Const WM_DDE_DATA = (WM_DDE_FIRST + 5)
Public Const WM_DDE_EXECUTE = (WM_DDE_FIRST + 8)
Public Const WM_DDE_INITIATE = (WM_DDE_FIRST)
Public Const WM_DDE_LAST = (WM_DDE_FIRST + 8)
Public Const WM_DDE_POKE = (WM_DDE_FIRST + 7)
Public Const WM_DDE_REQUEST = (WM_DDE_FIRST + 6)
Public Const WM_DDE_UNADVISE = (WM_DDE_FIRST + 3)
Public Const WM_DDE_TERMINATE = (WM_DDE_FIRST + 1)
Public Const WM_DEADCHAR = &H103
Public Const WM_DELETEITEM = &H2D
Public Const WM_DESTROY = &H2
Public Const WM_DESTROYCLIPBOARD = &H307
Public Const WM_DEVMODECHANGE = &H1B
Public Const WM_DRAWCLIPBOARD = &H308
Public Const WM_DRAWITEM = &H2B
Public Const WM_DROPFILES = &H233
Public Const WM_ENABLE = &HA
Public Const WM_ENDSESSION = &H16
Public Const WM_ENTERIDLE = &H121
Public Const WM_ENTERMENULOOP = &H211
Public Const WM_ERASEBKGND = &H14
Public Const WM_EXITMENULOOP = &H212
Public Const WM_FONTCHANGE = &H1D
Public Const WM_GETDLGCODE = &H87
Public Const WM_GETFONT = &H31
Public Const WM_GETHOTKEY = &H33
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_HOTKEY = &H312
Public Const WM_HSCROLL = &H114
Public Const WM_HSCROLLCLIPBOARD = &H30E
Public Const WM_ICONERASEBKGND = &H27
Public Const WM_IME_CHAR = &H286
Public Const WM_IME_COMPOSITION = &H10F
Public Const WM_IME_COMPOSITIONFULL = &H284
Public Const WM_IME_ENDCOMPOSITION = &H10E
Public Const WM_IME_KEYDOWN = &H290
Public Const WM_IME_KEYLAST = &H10F
Public Const WM_IME_KEYUP = &H291
Public Const WM_IME_NOTIFY = &H282
Public Const WM_IME_SELECT = &H285
Public Const WM_IME_SETCONTEXT = &H281
Public Const WM_IME_STARTCOMPOSITION = &H10D
Public Const WM_INITDIALOG = &H110
Public Const WM_INITMENU = &H116
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYFIRST = &H100
Public Const WM_KEYLAST = &H108
Public Const WM_KEYUP = &H101
Public Const WM_KILLFOCUS = &H8
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDICASCADE = &H227
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDINEXT = &H224
Public Const WM_MDIREFRESHMENU = &H234
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDISETMENU = &H230
Public Const WM_MDITILE = &H226
Public Const WM_MEASUREITEM = &H2C
Public Const WM_MENUCHAR = &H120
Public Const WM_MENUSELECT = &H11F
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSELAST = &H209
Public Const WM_MOUSEMOVE = &H200
Public Const WM_MOVE = &H3
Public Const WM_NCACTIVATE = &H86
Public Const WM_NCCALCSIZE = &H83
Public Const WM_NCCREATE = &H81
Public Const WM_NCDESTROY = &H82
Public Const WM_NCHITTEST = &H84
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCMBUTTONDBLCLK = &HA9
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCPAINT = &H85
Public Const WM_NCRBUTTONDBLCLK = &HA6
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_NEXTDLGCTL = &H28
Public Const WM_NULL = &H0
Public Const WM_PAINT = &HF
Public Const WM_PAINTCLIPBOARD = &H309
Public Const WM_PAINTICON = &H26
Public Const WM_PALETTECHANGED = &H311
Public Const WM_PALETTEISCHANGING = &H310
Public Const WM_PARENTNOTIFY = &H210
Public Const WM_PASTE = &H302
Public Const WM_PENWINFIRST = &H380
Public Const WM_PENWINLAST = &H38F
Public Const WM_POWER = &H48
Public Const WM_PSD_ENVSTAMPRECT = (WM_USER + 5)
Public Const WM_PSD_FULLPAGERECT = (WM_USER + 1)
Public Const WM_PSD_GREEKTEXTRECT = (WM_USER + 4)
Public Const WM_PSD_MARGINRECT = (WM_USER + 3)
Public Const WM_PSD_MINMARGINRECT = (WM_USER + 2)
Public Const WM_PSD_PAGESETUPDLG = (WM_USER)
Public Const WM_PSD_YAFULLPAGERECT = (WM_USER + 6)
Public Const WM_QUERYDRAGICON = &H37
Public Const WM_QUERYENDSESSION = &H11
Public Const WM_QUERYNEWPALETTE = &H30F
Public Const WM_QUERYOPEN = &H13
Public Const WM_QUEUESYNC = &H23
Public Const WM_QUIT = &H12
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RENDERALLFORMATS = &H306
Public Const WM_RENDERFORMAT = &H305
Public Const WM_SETCURSOR = &H20
Public Const WM_SETFOCUS = &H7
Public Const WM_SETFONT = &H30
Public Const WM_SETHOTKEY = &H32
Public Const WM_SETREDRAW = &HB
Public Const WM_SETTEXT = &HC
Public Const WM_SHOWWINDOW = &H18
Public Const WM_SIZE = &H5
Public Const WM_SIZECLIPBOARD = &H30B
Public Const WM_SPOOLERSTATUS = &H2A
Public Const WM_SYSCHAR = &H106
Public Const WM_SYSCOLORCHANGE = &H15
Public Const WM_SYSCOMMAND = &H112
Public Const WM_SYSDEADCHAR = &H107
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const WM_TIMECHANGE = &H1E
Public Const WM_TIMER = &H113
Public Const WM_UNDO = &H304
Public Const WM_VKEYTOITEM = &H2E
Public Const WM_VSCROLL = &H115
Public Const WM_VSCROLLCLIPBOARD = &H30A
Public Const WM_WINDOWPOSCHANGED = &H47
Public Const WM_WINDOWPOSCHANGING = &H46





'Private Function SetFormsRgn(ByVal frmhWnd As Long, ByVal NewRgn As Long, Optional ByVal nCombineMode As eCombineMode = RGN_OR) As Boolean
'    Dim endRgn As Long
'    Dim frmRgn As Long
'    Dim t As Long
'    Dim wr As RECT
'
'
'    SetFormsRgn = False
'
'    t = GetWindowRect(frmhWnd, wr)
'
'
'    If nCombineMode = RGN_SetTo Then
'        SetFormsRgn = SetWindowRgn(frmhWnd, NewRgn, True)
'    Else
'        frmRgn = CreateRectRgn(0, 0, wr.Right - wr.Left, wr.Bottom - wr.Top)
'        t = GetWindowRgn(frmhWnd, frmRgn)
'
'        endRgn = CreateRectRgn(0, 0, 0, 0)
'        CombineRgn endRgn, frmRgn, NewRgn, nCombineMode
'
'        SetFormsRgn = SetWindowRgn(frmhWnd, endRgn, True)
'
'        If NewRgn <> 0 Then DeleteObject NewRgn
'    End If
'
'    If frmRgn <> 0 Then DeleteObject frmRgn
'
'End Function
'
'Public Function ApplyEllipticRgn(ByVal frmhWnd As Long, ByRef RgnRect As RECT, Optional ByVal nCombineMode As eCombineMode = RGN_OR) As Boolean
'    Dim NewRgn As Long
'
'    ApplyEllipticRgn = False
'
'    NewRgn = CreateEllipticRgn(RgnRect.Left, RgnRect.Top, RgnRect.Right, RgnRect.Bottom)
'    If NewRgn > 0 Then
'        ApplyEllipticRgn = SetFormsRgn(frmhWnd, NewRgn, nCombineMode)
'    End If
'
'
'End Function
'
'
'Public Function ApplyPolyRgn(ByVal frmhWnd As Long, ByRef NewPoins() As POINTAPI, ByVal PolyCount As Long, Optional ByVal pFillMode As Integer = 1, Optional ByVal nCombineMode As eCombineMode = RGN_OR) As Boolean
'    Dim NewRgn As Long
'
'    ApplyPolyRgn = False
'
'    NewRgn = CreatePolygonRgn(NewPoins(0), PolyCount, pFillMode)
'
'    If NewRgn > 0 Then
'        ApplyPolyRgn = SetFormsRgn(frmhWnd, NewRgn, nCombineMode)
'    End If
'End Function
'
'Public Function ApplyRectRgn(ByVal frmhWnd As Long, ByRef RgnRect As RECT, Optional ByVal nWidthEllips As Long = 0, Optional ByVal nHeightEllips As Long = 0, Optional Rounded As Boolean = False, Optional ByVal nCombineMode As eCombineMode = RGN_OR) As Boolean
'    Dim NewRgn As Long
'
'    ApplyRectRgn = False
'    If Rounded Then
'        NewRgn = CreateRoundRectRgn(RgnRect.Left, RgnRect.Top, RgnRect.Right, RgnRect.Bottom, nWidthEllips, nHeightEllips)
'    Else
'        NewRgn = CreateRectRgn(RgnRect.Left, RgnRect.Top, RgnRect.Right, RgnRect.Bottom)
'    End If
'
'    If NewRgn > 0 Then
'        ApplyRectRgn = SetFormsRgn(frmhWnd, NewRgn, nCombineMode)
'    End If
'
'End Function
'
'
'Public Function SetFormOnTopAll(ByVal frmhWnd As Long, ByVal YESNO As Boolean) As Long
'    Dim t&, t2&
'
'    If YESNO Then t2 = HWND_TOPMOST Else t2 = HWND_NOTOPMOST
'    t = SetWindowPos(frmhWnd, t2, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
'
'    SetFormOnTopAll = t
'
'End Function
'
'
'Public Sub StartFormMove(ByVal frmhWnd As Long)
'    ReleaseCapture
'    SendMessage frmhWnd, &HA1, 2, 0&
'End Sub
'
'
'Public Function isPointOnForm(ByVal frmhWnd As Long, ByVal X1 As Long, ByVal Y1 As Long) As Boolean
'    Dim frmRgn As Long
'    Dim t As Long
'    Dim wr As RECT
'
'
'
'    t = GetWindowRect(frmhWnd, wr)
'
'    frmRgn = CreateRectRgn(0, 0, wr.Right - wr.Left, wr.Bottom - wr.Top)
'    t = GetWindowRgn(frmhWnd, frmRgn)
'
'    isPointOnForm = PtInRegion(frmRgn, X1, Y1)
'
'
'End Function
'
