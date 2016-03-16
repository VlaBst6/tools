Attribute VB_Name = "modWinHack"
Option Explicit

Global MagnifierVisible As Boolean

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
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
'  STANDARD_RIGHTS_REQUIRED = &HF0000
  SYNCHRONIZE = &H100000
  PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
End Enum

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpstring As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassname As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long


Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, ByRef hINst As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long

Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long


Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const WM_PAINT = &HF

Public Type PointAPI
    x As Long
    Y As Long
End Type

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long

Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long



Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim s As String, h As String
    
    h = hwnd
    s = GetCaption(hwnd)
    s = IIf(Len(Trim(s)) > 0, " -- ", Empty) & s & " -- " & GetClass(hwnd)
    
    
    While Len(h) < 10: h = h & " ": Wend
    
    frmChildren.List1.AddItem h & s
    
    EnumChildProc = 1 'continue enum
End Function

'Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
'    On Error Resume Next
'    Dim sSave As String, x As Long, sHwnd As String
'    x = GetWindowTextLength(hwnd)
'    If x > 100 Then x = 100
'    sSave = Space$(x + 1)
'    GetWindowText hwnd, sSave, Len(sSave)
'    sSave = Left$(sSave, Len(sSave) - 1)
'    If sSave <> "" Then
'        sHwnd = hwnd
'        While Len(sHwnd) < 10: sHwnd = sHwnd & " ": Wend
'        frmChildren.List1.AddItem sHwnd & " - " & sSave
'    End If
'    'continue enumeration
'    EnumChildProc = 1
'End Function



Function GetCaption(hwnd As Long)
Dim hWndlength As Long, hWndTitle As String, a As Long

'Get the length of the caption
hWndlength = GetWindowTextLength(hwnd)

'Fill up a string with that amount of characters
hWndTitle = String$(hWndlength, 0)

'Fill the string with the real caption
a = GetWindowText(hwnd, hWndTitle, (hWndlength + 1))
GetCaption = hWndTitle
End Function


Public Function GetClass(hwnd As Long)
    Dim lpClassname As String, retVal As Long
    
    lpClassname = Space(256)
     
    retVal = GetClassName(hwnd, lpClassname, 256)
     
    GetClass = Left$(lpClassname, retVal)

End Function

Sub OutlineWindow(hwnd As Long)

    Dim hHBr As Long, r As RECT, hdc As Long
    Const HS_DIAGCROSS = 5
    
    If hwnd = 0 Then Exit Sub
    
    GetClientRect hwnd, r
    hHBr = CreateHatchBrush(HS_DIAGCROSS, vbRed)
    
    hdc = GetDC(hwnd)
    FrameRect hdc, r, hHBr
    
    ReleaseDC hwnd, hdc
    DeleteObject hHBr
    
End Sub

Sub RemoveRectOutline(hwnd As Long)
   
        RedrawWindow hwnd, ByVal 0&, ByVal 0&, 1 'RDW_INVALIDATE

End Sub
 

Function GetProcessPath(hwnd As Long) As String
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
        GetProcessPath = Left$(sPath, ret)
        'Debug.Print pId, Left$(sPath, ret)
        Call CloseHandle(hProc)
    'Else
    '        Debug.Print pId, Err.LastDllError
    End If
End Function
