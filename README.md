<div align="center">

## A Cool Directory Map


</div>

### Description

This simple sub routine is used to map out a directory tree starting with any given path. Can be easily modified to perform any task that requires scanning folders.

THIS IS RECURSIVE!
 
### More Info
 
Full starting path

ex: C:\windows\

ex: \\computername\share\

This routine is generic. As it stands it will just show the directory structure in the immediate window. However, it can be modified easily for file search.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Coolwick](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/coolwick.md)
**Level**          |Advanced
**User Rating**    |4.8 (86 globes from 18 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/coolwick-a-cool-directory-map__1-9655/archive/master.zip)





### Source Code

```
Private Sub CommandButton1_Click()
 Call DirMap("C:\Windows\")
 'Must have "\" at the end of the path
End Sub
Sub DirMap(ByVal Path As String)
On Error Resume Next
 Dim i, j, x As Integer 'All used as counters
 Dim Fname(), CurrentFolder, Temp As String
 Temp = Path
 If Dir(Temp, vbDirectory) = "" Then Exit Sub 'if there arent any sub directories the exit
 CurrentFolder = Dir(Temp, vbDirectory)
 'First get number of folders (Stored in i)
 Do While CurrentFolder <> ""
 If GetAttr(Temp & CurrentFolder) = vbDirectory Then
  If CurrentFolder <> "." And CurrentFolder <> ".." Then
  i = i + 1
  End If
 End If
 CurrentFolder = Dir
 Loop
 ReDim Fname(i) 'Redim the array with number of folders
 'now store the folder names
 CurrentFolder = Dir(Temp, vbDirectory)
 Do While CurrentFolder <> ""
 If GetAttr(Temp & CurrentFolder) = vbDirectory Then
  If CurrentFolder <> "." And CurrentFolder <> ".." Then
  j = j + 1
  Fname(j) = CurrentFolder
  Debug.Print Temp & Fname(j)
  End If
 End If
 CurrentFolder = Dir
 Loop
 ' For each folder check to see there are sub folders
 For x = 1 To i
 Call DirMap(Temp & Fname(x) & "\")
 Next
End Sub
```

