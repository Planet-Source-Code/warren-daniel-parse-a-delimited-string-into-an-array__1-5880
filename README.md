<div align="center">

## Parse a delimited string into an array


</div>

### Description

This code with scan through a string looking for a delimiter of your choice, and will put the text inbetween the delimiters into seperate elements of an array.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Warren Daniel](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/warren-daniel.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/warren-daniel-parse-a-delimited-string-into-an-array__1-5880/archive/master.zip)





### Source Code

```
Option Explicit
Dim Parsed() As String
Dim DelimitChr As String
Dim DelimitNum As Integer
Private Sub Form_Load()
Dim X As Integer
DelimitChr = Chr(1)
Dim ExampleString As String
ExampleString = "1" & DelimitChr & "2" & DelimitChr & "3" & DelimitChr
Call CountDelimit(ExampleString)
Call ParseData(ExampleString)
Call DisplayInfo
End Sub
Private Sub CountDelimit(StrData As String)
Dim X As Integer
Dim NxtPos As Integer
DelimitNum = 0
Do
X = X + 1
NxtPos = InStr(NxtPos + 1, StrData, DelimitChr)
If NxtPos = 0 Then ReDim Parsed(DelimitNum): Exit Sub
DelimitNum = DelimitNum + 1
Loop
End Sub
Private Sub ParseData(StrData As String)
Dim X As Integer
Dim PrevPos As Integer
Dim NxtPos As Integer
For X = 1 To DelimitNum
PrevPos = NxtPos
NxtPos = InStr(NxtPos + 1, StrData, DelimitChr)
Parsed(X - 1) = Mid(StrData, PrevPos + 1, NxtPos - PrevPos - 1)
Next X
End Sub
Private Sub DisplayInfo()
Dim X As Integer
Dim RetVal As String
For X = 0 To DelimitNum
RetVal = RetVal & Parsed(X) & vbCrLf
Next X
MsgBox RetVal
End Sub
```

