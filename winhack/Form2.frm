VERSION 5.00
Begin VB.Form frmMenus 
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   LinkTopic       =   "Form2"
   ScaleHeight     =   2790
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Do it"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   3900
      List            =   "Form2.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2400
      Width           =   2115
   End
   Begin VB.TextBox txtItem 
      Height          =   315
      Left            =   2880
      TabIndex        =   3
      ToolTipText     =   "item index"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtHandle 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "handle"
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   60
      Width           =   6975
   End
   Begin VB.Label Label1 
      Caption         =   "EnableMenuItem"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   2460
      Width           =   1275
   End
End
Attribute VB_Name = "frmMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpstring As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_COMMAND = &H111

Private Type MENUITEMINFO   ' 44 bytes
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
End Type

Private Const WM_USER = &H400
Private Const MIIM_CHECKMARKS = &H8
Private Const MIIM_DATA = &H20
Private Const MIIM_ID = &H2
Private Const MIIM_STATE = &H1
Private Const MIIM_SUBMENU = &H4
Private Const MIIM_TYPE = &H10

Private Const MF_INSERT = &H0
Private Const MF_CHANGE = &H80
Private Const MF_APPEND = &H100
Private Const MF_DELETE = &H200
Private Const MF_REMOVE = &H1000
Private Const MF_BYCOMMAND = &H0
Private Const MF_BYPOSITION = &H400
Private Const MF_SEPARATOR = &H800
Private Const MF_ENABLED = &H0
Private Const MF_GRAYED = &H1
Private Const MF_DISABLED = &H2
Private Const MF_UNCHECKED = &H0
Private Const MF_CHECKED = &H8
Private Const MF_USECHECKBITMAPS = &H200
Private Const MF_STRING = &H0
Private Const MF_BITMAP = &H4
Private Const MF_OWNERDRAW = &H100
Private Const MF_POPUP = &H10
Private Const MF_MENUBARBREAK = &H20
Private Const MF_MENUBREAK = &H40
Private Const MF_UNHILITE = &H0
Private Const MF_HILITE = &H80
Private Const MF_SYSMENU = &H2000
Private Const MF_RIGHTJUSTIFY = &H4000&
Private Const MF_HELP = &H4000
Private Const MF_MOUSESELECT = &H8000
Private Const MF_END = &H80
Private Const MFS_CHECKED = MF_CHECKED
Private Const MFS_ENABLED = MF_ENABLED
Private Const MFS_GRAYED = &H3&
Private Const MFS_HILITE = MF_HILITE
Private Const MFS_DISABLED = MFS_GRAYED
Private Const MFS_UNCHECKED = MF_UNCHECKED
Private Const MFS_UNHILITE = MF_UNHILITE
Private Const MFT_BITMAP = MF_BITMAP
Private Const MFT_MENUBARBREAK = MF_MENUBARBREAK
Private Const MFT_MENUBREAK = MF_MENUBREAK
Private Const MFT_OWNERDRAW = MF_OWNERDRAW
Private Const MFT_RADIOCHECK = &H200&
Private Const MFT_RIGHTJUSTIFY = MF_RIGHTJUSTIFY
Private Const MFT_RIGHTORDER = &H2000&
Private Const MFT_SEPARATOR = MF_SEPARATOR
Private Const MFT_STRING = MF_STRING

Private s As Reporter



Private Sub Form_Load()
    If GetMenuItemCount(GetMenu(CLng(frmWinHack.txthWnd))) < 0 Then
        MsgBox "No Menues In this window", vbInformation
        Unload Me
    Else
        Set s = New Reporter
        Me.Caption = "Menu Reporter for hWnd: " & frmWinHack.txthWnd
        ViewMenu GetMenu(CLng(frmWinHack.txthWnd))
        Combo1.ListIndex = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set s = Nothing
End Sub

Private Sub Command1_Click()
    Dim flags As Long
    On Error GoTo hell
    Select Case Combo1.ListIndex
        Case 0: flags = MF_BYPOSITION Or MF_ENABLED
        Case 1: flags = MF_BYPOSITION Or MF_DISABLED Or MF_GRAYED
        Case 2: flags = MF_BYCOMMAND Or MF_ENABLED
        Case 3: flags = MF_BYCOMMAND Or MF_DISABLED Or MF_GRAYED
    End Select
    
    If Len(txtHandle) = 0 Or Not IsNumeric(txtHandle) Then
        MsgBox "Handle must be numeric value", vbInformation
        txtHandle.SetFocus
    End If
    
    If Len(txtItem) = 0 Or Not IsNumeric(txtItem) Then
        MsgBox "Handle must be numeric value", vbInformation
        txtHandle.SetFocus
    End If
    
    flags = EnableMenuItem(CLng(txtHandle), CLng(txtItem), flags)
    If flags = -1 Then MsgBox "Menu Item not found :("
    
    Exit Sub
hell:     MsgBox Err.Description, vbInformation
End Sub

Private Sub ViewMenu(ByVal menuhnd&)
    Dim menulen&, i&, menuId&, currentpopup&, db%, menuflags%, flagstring$
    Dim menuinfo As MENUITEMINFO
    Dim menustring2 As String * 128
    Dim trackpopups&(32)
   
    currentpopup = 0
    menulen = GetMenuItemCount(menuhnd)
    
    s.add "Menu handle " & menuhnd & " : " & Str$(menulen) & " entries"
    
    For i = 0 To menulen - 1
        menuId = GetMenuItemID(menuhnd, i)
        Select Case menuId
            Case 0: s.add "Seperator - index: " & i & "  MenuId: " & menuId
            
            Case -1 ' It's a popup menu
                ' Save it in the list of popups
                trackpopups&(currentpopup) = i
                currentpopup = currentpopup + 1
                
                db = GetMenuString(menuhnd, i, menustring2, 127, MF_BYPOSITION)
                menuflags = GetMenuState(menuhnd, i, MF_BYPOSITION)
                
                s.add "Popup menu: " & Left$(menustring2, db) & _
                      "  Index: " & i & _
                      "  MenuId: " & Str$(menuId) & _
                      "  Handle: " & GetSubMenu(menuhnd, i) & _
                      "  Flags: " + GetFlagString$(menuflags)
                      
            Case Else ' A regular entry
                db = GetMenuString(menuhnd, menuId, menustring2, 127, MF_BYCOMMAND)
                menuflags = GetMenuState(menuhnd, menuId, MF_BYCOMMAND)
                s.add "Index: " & i & _
                      "  MenuId: " & Str$(menuId) & _
                      "  Text: " & Left$(menustring2, db) & _
                      "  Flags: " + GetFlagString$(menuflags)
        End Select
    Next i
    
If currentpopup > 0 Then ' At least one popup was found
    s.add "Sub menus:"
    For i = 0 To currentpopup - 1
        menuId = trackpopups&(i)
        ' Recursively analyze the popup menu.
        s.AddNode
        ViewMenu GetSubMenu(menuhnd, menuId)
        s.CloseNode
    Next
End If

Text1 = Text1 & s.Report
s.NewReport

End Sub

Private Function GetFlagString$(menuflags%)
    Dim f$
    If (menuflags% And MF_CHECKED) <> 0 Then
        f$ = f$ + "Checked "
    Else
        f$ = f$ + "Unchecked "
    End If
    If (menuflags% And MF_DISABLED) <> 0 Then
        f$ = f$ + "Disabled "
    Else
        f$ = f$ + "Enabled "
    End If
    If (menuflags% And MF_GRAYED) <> 0 Then f$ = f$ + "Grayed "
    If (menuflags% And MF_BITMAP) <> 0 Then f$ = f$ + "Bitmap "
    If (menuflags% And MF_MENUBARBREAK) <> 0 Then f$ = f$ + "Bar-break "
    If (menuflags% And MF_MENUBREAK) <> 0 Then f$ = f$ + "Break "
    If (menuflags% And MF_SEPARATOR) <> 0 Then f$ = f$ + "Seperator "
    GetFlagString$ = f$
End Function


'Private Declare Function GetLastError Lib "kernel32" () As Long
'Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpstring As String) As Long
'Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
'Private Declare Function LoadMenu Lib "user32" Alias "LoadMenuA" (ByVal hInstance As Long, ByVal lpstring As String) As Long
'Private Declare Function LoadMenuIndirect Lib "user32" Alias "LoadMenuIndirectA" (ByVal lpMenuTemplate As Long) As Long
'Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
'Private Declare Function HiliteMenuItem Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal wIDHiliteItem As Long, ByVal wHilite As Long) As Long
'Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
'Private Declare Function CreateMenu Lib "user32" () As Long
'Private Declare Function CreatePopupMenu Lib "user32" () As Long
'Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
'Private Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
'
'Private Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
'Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
'Private Declare Function AppendMenuBynum Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
'Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpstring As String) As Long
'Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
'Private Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long
'Private Declare Function TrackPopupMenuBynum Lib "user32" Alias "TrackPopupMenu" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Long) As Long
'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function ModifyMenuBynum Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpstring As Long) As Long
'
'Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Private Declare Function GetMenuContextHelpId Lib "user32" (ByVal hMenu As Long) As Long
'Private Declare Function GetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long
'Private Declare Function SetMenuContextHelpId Lib "user32" (ByVal hMenu As Long, ByVal dw As Long) As Long
'Private Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
'Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Long, lpcMenuItemInfo As MENUITEMINFO) As Long
