<div align="center">

## Add User Initials to a block of selected text


</div>

### Description

If you have to add your username or initals tag to each change you make in code, this lets you select a block of text and add it to the end of each line.
 
### More Info
 
add this code to the mnuHandler_Click

It doesn't delete any old comments that were there.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ustes](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ustes.md)
**Level**          |Intermediate
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ustes-add-user-initials-to-a-block-of-selected-text__1-5835/archive/master.zip)





### Source Code

```

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Dim MyPane As CodePane
  Dim lngStartLine As Long
  Dim lngEndLine As Long
  Dim lngStartCol As Long
  Dim lngEndCol As Long
  Dim strLine As String
  Dim tmpLine As String
  Dim i As Integer
  Dim LineLengths() As Integer
  Dim intLongestLine As Integer
  Dim intTotalLines As Integer
  Dim intLinecount As Integer
  Dim intDiff As Integer
  If strUser = "" Then
    strUser = "'"
    strUser = strUser & InputBox("Enter User Initials.", "Block Initials")
    strUser = strUser & " - " & Format(Now, "mm/dd/yy hh:mm")
  End If
  Set MyPane = VBInstance.ActiveCodePane
  MyPane.GetSelection lngStartLine, lngStartCol, lngEndLine, lngEndCol
  intTotalLines = lngEndLine - lngStartLine
  ReDim LineLengths(intTotalLines)
  intLinecount = 0
  For i = lngStartLine To lngEndLine - 1
    strLine = MyPane.CodeModule.Lines(i, 1)
    If intLongestLine < Len(strLine) Then
      LineLengths(intLinecount) = Len(strLine)
      intLongestLine = LineLengths(intLinecount)
    End If
    intLinecount = intLinecount + 1
  Next i
  For i = lngStartLine To lngEndLine - 1
    strLine = MyPane.CodeModule.Lines(i, 1)
    tmpLine = strLine
    If Trim(tmpLine) <> "" Then
      intDiff = intLongestLine - Len(strLine)
      MyPane.CodeModule.ReplaceLine i, strLine & Space(intDiff + 5) & strUser
    End If
  Next i
End Sub
```

