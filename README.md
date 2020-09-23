<div align="center">

## Determine if File Is Old


</div>

### Description

Determines if a file is old. I use this when I loop through the files in a "temp" directory to determine if I should delete old files on a website. Take note - the function looks at the last modified date rather then the date created.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lewis E\. Moten III](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lewis-e-moten-iii.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lewis-e-moten-iii-determine-if-file-is-old__1-24318/archive/master.zip)





### Source Code

```
Private Sub Form_Load()
  MsgBox IIf(FileIsOld("C:\AutoExec.bat"), "The file is old", "The file is new")
End Sub
Function FileIsOld(ByRef pStrFilePath As String) As Boolean
  Dim llngMinutesOld As Long
  Dim ldtmLastModified As Date
  Dim llngFileAttr As VbFileAttribute
  Const llngMinutesOldAfter As Long = 10
  On Error Resume Next
  llngFileAttr = FileSystem.GetAttr(pStrFilePath)
  If Err Then
    MsgBox "File does not exist."
    Exit Function ' file doesn't exist
  End If
  On Error GoTo 0
  If Len(FileSystem.Dir(pStrFilePath, llngFileAttr)) = 0 Then Exit Function
  ldtmLastModified = FileSystem.FileDateTime(pStrFilePath)
  llngMinutesOld = DateDiff("n", ldtmLastModified, Now())
  FileIsOld = llngMinutesOld > pLngMinutesOldAfter
End Function
```

