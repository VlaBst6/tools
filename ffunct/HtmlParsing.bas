Attribute VB_Name = "HtmlParsing"

Public Function parseScript(info) As String  'cut scripts out of html
  Dim EndOfScript As Integer, scripData As String, trimpage As String
  info = filt(info, "javascript,vbscript,mocha,createobject,activex,onload,onerror,onclick,onmove,onscroll,onmouse,onresize,onkey,iframe")
  script = Split(info, "<script", , vbTextCompare)
  If UBound(script) = 0 Then parseScript = info: Exit Function _
  Else: trimpage = script(0)
  For i = 1 To UBound(script)
    EndOfScript = InStr(1, script(i), "</script", vbTextCompare)
    trimpage = trimpage & Mid(script(i), EndOfScript + 10, Len(script(i)))
  Next
  parseScript = trimpage
End Function


Public Function parseHtml(info) As String
     Dim temp As String, EndOfTag As Integer
     fmat = Replace(info, "&nbsp;", " ")
     cut = Split(fmat, "<")

   For i = 0 To UBound(cut)  'cut at all html start tags
     EndOfTag = InStr(1, cut(i), ">")
        If EndOfTag > 0 Then
          EndOfText = Len(cut(i))
          NL = False
          If Left(cut(i), 2) = "br" Then NL = True
          cut(i) = Mid(cut(i), EndOfTag + 1, EndOfText)
          If NL Then cut(i) = vbCrLf & cut(i)
          If cut(i) = vbCrLf Then cut(i) = ""
        End If
     temp = temp & cut(i)
    Next
    
    parseHtml = temp
End Function

Public Function parseAnds(info)  'trims out &amp; type html for text
  Dim temp As String
  cut = Split(info, "&")
  If UBound(cut) > 0 Then
    For i = 0 To UBound(cut)            'cut at all start tags (&)
      EndOfTag = InStr(1, cut(i), ";")
        If EndOfTag > 0 Then
           EndOfText = Len(cut(i))
           cut(i) = Mid(cut(i), EndOfTag + 1, EndOfText)
        End If
      temp = temp & cut(i)
    Next
   parseAnds = temp
  Else: parseAnds = info
  End If
End Function


