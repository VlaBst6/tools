Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = -4

Private gOldWinProc As Long
Private HookedHwnd As Long  'dont let em loose this
Public gisHooked As Boolean
Public gChileEnumContinue As Boolean

Public Sub Hook(ByVal hwnd As Long)
   HookedHwnd = hwnd
   gOldWinProc = SetWindowLong(HookedHwnd, GWL_WNDPROC, AddressOf WindowProc)
   gisHooked = True
End Sub

Public Sub Unhook()
    Dim temp As Long
    
    If HookedHwnd <> 0 Then
        temp = SetWindowLong(HookedHwnd, GWL_WNDPROC, gOldWinProc)
    End If
    
    HookedHwnd = 0
    gisHooked = False
End Sub

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim s As String
    
    s = "hWnd=" & hwnd & ", uMsg= " & getMsg("&H" & Hex(uMsg)) & ", wParam=" & wParam & ", lParam=" & lParam
    If uMsg <> 32 And uMsg <> 132 And uMsg <> 512 And uMsg <> &HA0 And uMsg <> &H121 Then
        If uMsg < 307 Or uMsg > 312 Then
            'exclude paint, hittest, mousemove etc (spam)
            Form1.List1.AddItem s
            If Form1.Check1.value = 1 Then AppendFile App.path & "\log.txt", s
        End If
    End If
    WindowProc = CallWindowProc(gOldWinProc, hwnd, uMsg, wParam, lParam)
End Function

Function getMsg(key As String) As String
 On Error GoTo shit
 getMsg = msgs(key)
 Exit Function
shit: getMsg = key
End Function

Sub AppendFile(path As String, it As String)
    Dim f As Long
    f = FreeFile
    Open path For Append As #f
    Print #f, it
    Close f
End Sub
