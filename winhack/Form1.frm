VERSION 5.00
Begin VB.Form frmChildren 
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7155
   End
End
Attribute VB_Name = "frmChildren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Caption = "Child Windows of: " & frmWinHack.txthWnd
    x = EnumChildWindows(CLng(frmWinHack.txthWnd), AddressOf EnumChildProc, ByVal 0&)
    If x = 0 Then
        MsgBox "This window has no children", vbInformation
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    List1.Width = Me.Width - List1.Left - 150
    List1.Height = Me.Height - List1.Top - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    RemoveRectOutline CLng(frmWinHack.txthWnd)
End Sub

Private Sub List1_DblClick()
    t = GetHandleFromList()
    Screen.MousePointer = vbHourglass
    Clipboard.Clear
    Clipboard.SetText t
    Sleep 200
    Screen.MousePointer = vbDefault
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Dim hwnd As Long
    
    hwnd = CLng(GetHandleFromList())
    
    RemoveRectOutline CLng(frmWinHack.txthWnd)
    OutlineWindow hwnd
    
    If Button = 2 Then frmWinHack.txthWnd = hwnd
         
End Sub

Function GetHandleFromList()
    t = List1.List(List1.ListIndex)
    If t = Empty Then Exit Function
    t = Mid(t, 1, InStr(t, " "))
    GetHandleFromList = t
End Function
