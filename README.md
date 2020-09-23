<div align="center">

## exact copy of mircs $gettok function


</div>

### Description

the same as the mirc gettok function - very usefull
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Hans Otto der dritte](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/hans-otto-der-dritte.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/hans-otto-der-dritte-exact-copy-of-mircs-gettok-function__1-57315/archive/master.zip)





### Source Code

```
' mIRC gettok func for vb
' (c) by diChter, www.diChtbox.de.vu
' usage: gettok(text, N, asciival char)
' example:
' MsgBox gettok(this-really-owns-yo-momma, "2-4", 45)
'  would return 'really-owns-yo'
' MsgBox gettok(this-really-owns-yo-momma, "3", 45)
'  would return 'owns'
Public Function gettok(t As String, n As String, c As Integer)
 On Error Resume Next             ' just in case..
 Dim splitted
 splitted = Split(t, Chr(c))          ' splits text by token
 If n = "0" Then                ' if n is 0 return num of tokens
  gettok = UBound(splitted) + 1
 ElseIf InStr(n, "-") Then           ' if '-' is in n..
  Dim x As Integer, r As String
  If Right(n, 1) = "-" Then          ' if n format = x-
   n = Left(n, Len(n) - 1)
   For x = Int(n) To UBound(splitted) + 1  ' all tokens started from x
    r = r & Chr(c) & splitted(x - 1)
   Next
   gettok = Mid(r, 2)
  ElseIf Not Left(n, 1) = "-" Then      ' if format not = x- and not -x it should be x-x
   Dim splittedN
   splittedN = Split(n, "-")         ' split n to get x1 'n x2
   For x = splittedN(0) To splittedN(1)   ' all tokens from x1 to x2
    r = r & Chr(c) & splitted(x - 1)
   Next
   gettok = Mid(r, 2)
  End If
 Else                     ' 1 token
  gettok = splitted(Int(n) - 1)
 End If
End Function
```

