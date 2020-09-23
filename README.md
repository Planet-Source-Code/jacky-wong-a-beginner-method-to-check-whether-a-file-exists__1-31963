<div align="center">

## A Beginner method to check whether a file exists


</div>

### Description

A Beginner method to check whether the file exists or not. It is simple, easy to understand and should able to run on different Window OS.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jacky Wong](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jacky-wong.md)
**Level**          |Beginner
**User Rating**    |3.9 (27 globes from 7 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jacky-wong-a-beginner-method-to-check-whether-a-file-exists__1-31963/archive/master.zip)





### Source Code

Public Function CheckFile(InFileName As String) As Boolean <br>
 On Error GoTo ErrHandler <br>
 CheckFile = False <br>
 'Check whether the file or folder exist <br>
 If Dir(InFileName) <> "" Then <br>
 'Check is it a directory(folder) <br>
 If (GetAttr(InFileName) And vbDirectory) = 0 Then <br>
  CheckFile = True <br>
 Else <br>
  MsgBox "File doesn't exist!", vbCritical <br>
  Exit Function <br>
 End If <br>
 Else <br>
 MsgBox "File doesn't exist!", vbCritical <br>
 Exit Function <br>
 End If <br>
ErrHandler: <br>
End Function

