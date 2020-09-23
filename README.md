<div align="center">

## Advanced Like


</div>

### Description

Compare using wildcards like * and ?, ranges like "at[0-99].gif", and a new wildcard %. Which is like *, but goes only at the end. "at%" would be like "at", and also "atquaz".
 
### More Info
 
filter - a pattern to compare with

expression - the text to check

boolean


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Techni Rei Myoko](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/techni-rei-myoko.md)
**Level**          |Advanced
**User Rating**    |3.8 (15 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/techni-rei-myoko-advanced-like__1-47095/archive/master.zip)





### Source Code

```
Option Explicit
Public Function advlike(filter As String, expression As String) As Boolean
Dim curr_filter As Long, curr_text As Long, buffer As Boolean, temp As Long, tempstr As String, temp2 As Long, tempstr2 As String
curr_text = 1
buffer = True
Do Until curr_filter = Len(filter) Or buffer = False
  curr_filter = curr_filter + 1
  Select Case Mid(filter, curr_filter, 1)
    Case "*"
      If curr_filter = Len(filter) Then
        curr_text = Len(expression) - 1
      Else
        curr_text = InStr(curr_text, expression, Mid(filter, curr_filter + 1, 1)) - 1
        If curr_text <= 0 Then buffer = False
      End If
    Case "%": curr_text = Len(expression) - 1
    Case "?" 'should just skip right over this with no problem at all
    Case "["
      temp = InStr(curr_filter, filter, "]") 'contains the ending ("]") delimeter for qualifications
      tempstr = Mid(filter, curr_filter + 1, temp - curr_filter - 1) 'contains qualifications
      'curr_text contains the start of the expression
      If curr_filter = Len(filter) Then
        temp2 = Len(expression) ' contains the end of the expression
      Else
        tempstr2 = Mid(filter, InStr(curr_filter, filter, "]") + 1, 1) ' contains the end of the expression
        temp2 = InStr(curr_text, expression, tempstr2)
      End If
      If temp2 = 0 Then
        buffer = False
      Else
        tempstr2 = Mid(expression, curr_text, temp2 - curr_text) 'contains expression
        If multicompare(tempstr2, tempstr) = False Then
          buffer = False
        Else
          curr_text = curr_text + Len(tempstr2) - 1
          curr_filter = curr_filter + Len(tempstr) + 1
        End If
      End If
    Case Else: If Mid(filter, curr_filter, 1) <> Mid(expression, curr_text, 1) Then buffer = False
  End Select
  curr_text = curr_text + 1
  'if current text loc is past the end of the expression when there is still untested filter chars
  If curr_text > Len(expression) And curr_filter + 1 < Len(filter) Then buffer = False
Loop
advlike = buffer
End Function
Public Function multicompare(text As String, qualifications As String) As Boolean
qualifications = Replace(qualifications, " ", Empty)
If InStr(qualifications, ",") = 0 Then
  multicompare = compare(text, qualifications)
Else
  Dim temp As Long, tempstr() As String
  tempstr = Split(qualifications, ",")
  For temp = 0 To UBound(tempstr)
    If compare(text, tempstr(temp)) Then multicompare = True
  Next
End If
End Function
Public Function compare(text As String, qualifier As String)
  Dim tempstr() As String
  If InStr(qualifier, "-") > 0 Then
    tempstr = Split(qualifier, "-")
    If isnumeric2(tempstr(0)) And isnumeric2(tempstr(1)) Then
      compare = Val(text) >= Val(tempstr(0)) And Val(text) <= Val(tempstr(1))
    Else
      compare = text >= tempstr(0) And text <= tempstr(1)
    End If
  Else
    If isnumeric2(qualifier) Then
      compare = Val(text) = Val(qualifier)
    Else
      compare = text = qualifier
    End If
  End If
End Function
Public Function islike(filter As String, expression As String) As Boolean
  On Error Resume Next
  Dim tempstr() As String, count As Long
  If Replace(filter, ";", Empty) <> filter Then
    tempstr = Split(filter, ";")
    islike = False
    For count = LBound(tempstr) To UBound(tempstr)
      If advlike(tempstr(count), expression) Then islike = True
    Next
  Else
    If advlike(filter, expression) Then islike = True
  End If
End Function
Public Function isnumeric2(text As String) As Boolean
isnumeric2 = IsNumeric(Replace(Replace(text, "-", Empty), ".", Empty))
End Function
```

