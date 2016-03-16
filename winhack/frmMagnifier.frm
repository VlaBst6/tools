VERSION 5.00
Begin VB.Form frmMagnifier 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   Icon            =   "frmMagnifier.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMagnifier.frx":0442
   ScaleHeight     =   208
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   222
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkOn 
      Caption         =   "Check1"
      Height          =   195
      Left            =   420
      TabIndex        =   2
      Top             =   2340
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2280
      Top             =   2640
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2265
      Left            =   600
      ScaleHeight     =   151
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   156
      TabIndex        =   0
      Top             =   75
      Width           =   2340
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   240
   End
End
Attribute VB_Name = "frmMagnifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dhwnd As Long, dhdc As Long
Dim mouse As PointAPI
Dim ZF As Integer

'magnifier project isnt mine, its from pscode.com i just dig it and
'think its handy do here it is..will be useful for color spy too
'author:  | LOST |

Private Declare Function CreateEllipticRgn& Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Const RGN_OR = 2


Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

'private  Type PointAPI
'  x As Long
'  y As Long
'End Type
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'private  Declare Function GetDesktopWindow Lib "user32" () As Long
'private  Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
'private  Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Const SRCCOPY = &HCC0020

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

'declare for moving the form
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_SYSCOMMAND = &H112

'for translucent effect in win2k, remove this if run in win9x or NT
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1

Private Const VK_F09 = &H78
Private Const VK_F10 = &H79
Private Const SW_HIDE = 0
Private Const SW_NORMAL = 1
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
  Private Const HWND_TOPMOST = -1
  Private Const HWND_NOTOPMOST = -2
  Private Const SWP_NOACTIVATE = &H10
  Private Const SWP_SHOWWINDOW = &H40

Public Sub AlwaysOnTop(formname As Form, SetOnTop As Boolean)
    Dim lflag As Long
    If SetOnTop Then
        lflag = HWND_TOPMOST
    Else
        lflag = HWND_NOTOPMOST
    End If
    SetWindowPos formname.hwnd, lflag, _
    formname.Left / Screen.TwipsPerPixelX, _
    formname.Top / Screen.TwipsPerPixelY, _
    formname.Width / Screen.TwipsPerPixelX, _
    formname.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub
Public Sub SetAutoRgn(hForm As Form, Optional transColor As Long = vbNull)
  Dim x As Long, y As Long
  Dim Rgn1 As Long, Rgn2 As Long
  Dim SPos As Long, EPos As Long
  Dim wID As Long, Hgt As Long
  Dim xoff As Long, yoff As Long
  Dim DIB As New cDIBSection
  Dim bDib() As Byte
  Dim tSA As SAFEARRAY2D
  
    'get the picture size of the form
  DIB.CreateFromPicture hForm.Picture
  wID = DIB.Width
  Hgt = DIB.Height
  
  With hForm
    .ScaleMode = vbPixels
    'compute the title bar's offset
    xoff = (.ScaleX(.Width, vbTwips, vbPixels) - .ScaleWidth) / 2
    yoff = .ScaleY(.Height, vbTwips, vbPixels) - .ScaleHeight - xoff
    'change the form size
    .Width = (wID + xoff * 2) * Screen.TwipsPerPixelX
    .Height = (Hgt + xoff + yoff) * Screen.TwipsPerPixelY
  End With
  
  ' have the local matrix point to bitmap pixels
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = DIB.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = DIB.BytesPerScanLine
        .pvData = DIB.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
        
      
' if there is no transColor specified, use the first pixel as the transparent color
  If transColor = vbNull Then transColor = RGB(bDib(0, 0), bDib(1, 0), bDib(2, 0))
  
  Rgn1 = CreateRectRgn(0, 0, 0, 0)
  
  For y = 0 To Hgt - 1 'line scan
    x = -3
    Do
     x = x + 3
     
     While RGB(bDib(x, y), bDib(x + 1, y), bDib(x + 2, y)) = transColor And (x < wID * 3 - 3)
       x = x + 3 'skip the transparent point
     Wend
     SPos = x / 3
     While RGB(bDib(x, y), bDib(x + 1, y), bDib(x + 2, y)) <> transColor And (x < wID * 3 - 3)
       x = x + 3 'skip the nontransparent point
     Wend
     EPos = x / 3
     
     'combine the region
     If SPos <= EPos Then
         Rgn2 = CreateRectRgn(SPos + xoff, Hgt - y + yoff, EPos + xoff, Hgt - 1 - y + yoff)
         CombineRgn Rgn1, Rgn1, Rgn2, RGN_OR
         DeleteObject Rgn2
     End If
    Loop Until x >= wID * 3 - 3
  Next y
  
  SetWindowRgn hForm.hwnd, Rgn1, True  'set the final shap region
  DeleteObject Rgn1
 
End Sub

Private Sub Form_Load()
Dim lReigon&, lResult&, n$
    MagnifierVisible = True

    If Me.Picture <> 0 Then Call SetAutoRgn(Me)
    lReigon& = CreateEllipticRgn(0, 0, 157, 153)
    lResult& = SetWindowRgn(Picture1.hwnd, lReigon&, True)
    dhwnd = GetDesktopWindow
    dhdc = GetDC(dhwnd)
    AlwaysOnTop Me, True ' Use this as the call to this fuction.
    ZF = 200

    n = "frmMagnifier"
    Me.Left = GetSetting(App.Title, n, "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, n, "MainTop", 1000)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
       
      ReleaseCapture
      SendMessage Me.hwnd, WM_NCLBUTTONDOWN, 2, 0&
       
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim n$
    n = "frmMagnifier"
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, n, "MainLeft", Me.Left
        SaveSetting App.Title, n, "MainTop", Me.Top
    End If
    MagnifierVisible = False
    Unload Me
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
       
      ReleaseCapture
      SendMessage Me.hwnd, WM_NCLBUTTONDOWN, 2, 0&
       
    End If
End Sub


Sub Magnify(ptx, pty)
    Dim w As Long, h As Long, sw As Long, sh As Long, x As Long, y As Long
    If chkOn.value = 1 Then
      'GetCursorPos mouse        ' capture mouse-position
      mouse.x = ptx
      mouse.y = pty
      w = Picture1.ScaleWidth   ' destination width
      h = Picture1.ScaleHeight  ' destination height
      sw = w * (1 / (ZF / 100))
      sh = h * (1 / (ZF / 100))
      x = mouse.x - sw / 2       ' x source position (center to destination)
      y = mouse.y - sh / 2       ' y source position (center to destination)
      Picture1.Cls                ' clean picturebox
      StretchBlt Picture1.hdc, 0, 0, w, h, dhdc, x, y, sw, sh, SRCCOPY
      ' copy desktop (source) and strech to picturebox (destination)
    End If
End Sub


Private Sub Timer2_Timer()
Dim p As Long
If GetAsyncKeyState(VK_F09) Then p = ShowWindow(Me.hwnd, SW_HIDE)
If GetAsyncKeyState(VK_F10) Then p = ShowWindow(Me.hwnd, SW_NORMAL)
If Me.Visible Then If GetAsyncKeyState(&H26) Then If ZF < 500 Then ZF = ZF + 1
If Me.Visible Then If GetAsyncKeyState(&H28) Then If ZF > 1 Then ZF = ZF - 1
End Sub


