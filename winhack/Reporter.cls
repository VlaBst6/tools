VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Reporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strReport As String
Private nestlevel As Long

Property Get Report() As String
    Report = br(strReport)
End Property

Public Sub NewReport()
    strReport = Empty
    nestlevel = -1
End Sub

Public Sub AddNode()
    nestlevel = nestlevel + 1
End Sub

Public Sub CloseNode()
    If nestlevel > 0 Then nestlevel = nestlevel - 1
End Sub

Sub add(it As String, Optional newLine As Boolean = True)
    Dim s As String, i As Long
    For i = 0 To nestlevel: it = "\t" & it: Next
    strReport = strReport & IIf(newLine, vbCrLf, Empty) & it
End Sub

Private Function br(it As String) As String
    it = Replace(it, "\t", vbTab)
    br = Replace(it, "\n", vbCrLf)
End Function

Private Sub Class_Initialize()
    nestlevel = -1
End Sub
