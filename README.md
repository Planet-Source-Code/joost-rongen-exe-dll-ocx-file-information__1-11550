<div align="center">

## EXE / DLL / OCX  File \- information


</div>

### Description

This sample project shows how to get the following information out of EXE / DLL / OCX files: CompanyName , FileDescription , fileVersion , InternalName , LegalCopyright , OriginalFileName , ProductName , ProductVersion
 
### More Info
 
Name of file to inquire

CompanyName , FileDescription , fileVersion , InternalName , LegalCopyright , OriginalFileName , ProductName , ProductVersion

none (tested under W2K)


<span>             |<span>
---                |---
**Submitted On**   |2000-09-17 13:39:40
**By**             |[Joost Rongen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/joost-rongen.md)
**Level**          |Intermediate
**User Rating**    |4.8 (63 globes from 13 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD99879192000\.zip](https://github.com/Planet-Source-Code/joost-rongen-exe-dll-ocx-file-information__1-11550/archive/master.zip)

### API Declarations

```
Declare Function GetFileVersionInfo Lib "Version.dll" Alias _
 "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal _
 dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias _
 "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, _
 lpdwHandle As Long) As Long
Declare Function VerQueryValue Lib "Version.dll" Alias _
 "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, _
 lplpBuffer As Any, puLen As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias _
 "GetSystemDirectoryA" (ByVal Path As String, ByVal cbBytes As _
 Long) As Long
Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
 dest As Any, ByVal Source As Long, ByVal Length As Long)
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" ( _
 ByVal lpString1 As String, ByVal lpString2 As Long) As Long
' -----------------------------------
```





