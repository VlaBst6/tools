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
   Begin WinHack.ucFilterList lv 
      Height          =   2625
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   4630
   End
End
Attribute VB_Name = "frmChildren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    lv.SetColumnHeaders "hwnd,class,caption*", "800,1800"
    Me.Caption = "Child Windows of: " & frmWinHack.txthWnd
    x = EnumChildWindows(CLng(frmWinHack.txthWnd), AddressOf EnumChildProc, ByVal 0&)
    If x = 0 Then
        MsgBox "This window has no children", vbInformation
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lv.Move 0, 0, Me.Width - 200, Me.Height - 400
End Sub

Private Sub Form_Unload(cancel As Integer)
    On Error Resume Next
    RemoveRectOutline CLng(frmWinHack.txthWnd)
End Sub

Private Sub lv_DblClick()
    If lv.selItem Is Nothing Then Exit Sub
    Dim hwnd As Long
    hwnd = CLng(lv.selItem.Text)
    Screen.MousePointer = vbHourglass
    Clipboard.Clear
    Clipboard.SetText hwnd
    Sleep 200
    Screen.MousePointer = vbDefault
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error Resume Next
    Dim hwnd As Long
    
    If lv.selItem Is Nothing Then Exit Sub
    
    hwnd = CLng(lv.selItem.Text)
    
    RemoveRectOutline CLng(frmWinHack.txthWnd)
    OutlineWindow hwnd
    
    If Button = 2 Then frmWinHack.txthWnd = hwnd
End Sub
