<div align="center">

## A Function that removes all of a chr from a string


</div>

### Description

This is a fairly basic code, but I had someone ask me how to do it, so I figured since I made it I might as well post it. Basicly this is a simple function which you can add to either a bas, or a form, and then you call it, and it removes all occurenses of a character within a string. IE.(My name is steve) if you wanted to remove all the spaces you could use this.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Steve](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steve.md)
**Level**          |Unknown
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steve-a-function-that-removes-all-of-a-chr-from-a-string__1-3838/archive/master.zip)

### API Declarations

```
'Check out my website at http://www.vbtutor.com
'Thanks
```


### Source Code

```
Function RemoveChar(sText As String, sChar As String) As String
  Dim iPos As Integer, iStart As Integer
  Dim sTemp As String
  iStart = 1
  Do
    iPos = InStr(iStart, sText, sChar)
    If iPos <> 0 Then
      sTemp = sTemp & Mid(sText, iStart, (iPos - iStart))
      iStart = iPos + 1
    End If
  Loop Until iPos = 0
  sTemp = sTemp & Mid(sText, iStart)
  RemoveChar = sTemp
End Function
'The code could then be called like this
Call RemoveChar(Text1.text, " ")
'This will rmove all the spaces from the textbox
'named Text1
'I hope this helps some people out. I have actualy
'surprisignly enought had 37 requests from visitors
'to my site for this code.
```

