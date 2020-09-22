Attribute VB_Name = "AppPath"
Option Explicit

Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
    (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
    lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, _
    lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long

Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const KEY_ALL_ACCESS = &H3F
Private Const SE_ERR_FNF = 2&
Private Const SE_ERR_PNF = 3&
Private Const SE_ERR_ACCESSDENIED = 5&
Private Const SE_ERR_OOM = 8&
Private Const SE_ERR_DLLNOTFOUND = 32&
Private Const SE_ERR_SHARE = 26&
Private Const SE_ERR_ASSOCINCOMPLETE = 27&
Private Const SE_ERR_DDETIMEOUT = 28&
Private Const SE_ERR_DDEFAIL = 29&
Private Const SE_ERR_DDEBUSY = 30&
Private Const SE_ERR_NOASSOC = 31&
Private Const ERROR_BAD_FORMAT = 11&
Private Const SW_HIDE = 0
Private Const SW_NORMAL = 1
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_RESTORE = 9
Public Function NameNoExtension(ByVal Name As String) As String
NameNoExtension = Left(NameFromPath(Name), InStr(1, NameFromPath(Name), ".") - 1)
End Function


Public Function WeekdayName(Weekday As Integer) As String
Select Case Weekday
Case 1
WeekdayName = "Sunday"
Case 2
WeekdayName = "Monday"
Case 3
WeekdayName = "Tuesday"
Case 4
WeekdayName = "Wednesday"
Case 5
WeekdayName = "Thursday"
Case 6
WeekdayName = "Friday"
Case 7
WeekdayName = "Saturday"
End Select
End Function

Public Function DrivePath(FullPath As String) As String
DrivePath = Left(FullPath, 3)
End Function


Public Function NameFromPath(FullPath As String) As String
NameFromPath = Right(FullPath, Len(FullPath) - InStrRev(FullPath, "\"))
End Function

Public Sub BuildPath(ByVal Path As String)
Dim FileAttribs As VbFileAttribute
Dim DirAttribs As VbFileAttribute
Dim NewPath As String
NewPath = Path
FileAttribs = vbNormal + vbReadOnly + vbHidden + vbSystem + vbArchive
DirAttribs = vbDirectory
If Len(Dir(Path, FileAttribs + DirAttribs)) > 0 Then
Exit Sub
ElseIf Right(Path, 1) <> "\" Then
NewPath = Path & "\"
Else
NewPath = Path
End If
MakeSureDirectoryPathExists NewPath
End Sub

Public Function SlashPath(FullPath As String) As String
SlashPath = Left(FullPath, InStrRev(FullPath, "\"))
End Function


Public Function FolderPath(FullPath As String) As String
FolderPath = Left(FullPath, InStrRev(FullPath, "\") - 1)
End Function


Public Function LongName(ShortName As String) As String
    
Dim temp As String
Dim NewString As String
Dim Searched As Boolean
Dim i As Integer

If Len(ShortName) = 0 Then Exit Function

temp = ShortName
If Right(temp, 1) = "\" Then
temp = Left(temp, Len(temp) - 1)
Searched = True
End If

On Error GoTo NoFile:
If InStr(temp, "\") Then
    NewString = ""
    Do While InStr(temp, "\")
        If Len(NewString) Then
        NewString = Dir(temp, 55) & "\" & NewString
        Else
        NewString = Dir(temp, 55)
            If Len(NewString) = 0 Then
            LongName = ShortName
            Exit Function
            End If
        End If
        On Error Resume Next
        For i = Len(temp) To 1 Step -1
            If ("\" = Mid(temp, i, 1)) Then
            Exit For
            End If
        Next
        temp = Left(temp, i - 1)
    Loop
    NewString = temp & "\" & NewString
Else
NewString = Dir(temp, 55)
End If

Here:
If Searched Then
NewString = NewString & "\"
End If

LongName = PrettyPath(NewString)
Exit Function
NoFile:
NewString = ""
Resume Here:
End Function

Public Function GetAppPath(ByVal AppName As String) As String
'from the registry if it exists
'returns vbNullstring if not

On Error GoTo TheEnd:
Dim TheResult As Long
Dim Index As Long
Dim TheEntry As String
Dim EntryLength As Long
Dim TheDataType As Long
Dim TheByteArray(1 To 1024) As Byte
Dim DataLength As Long
Dim ByteValue As String
Dim i As Integer
Dim MainKey As Long
Dim SubKey As String
Dim mKey As Long

If LCase(Right(AppName, 4)) <> ".exe" Then
AppName = AppName & ".exe"
End If

MainKey = HKEY_LOCAL_MACHINE
SubKey = "Software\Microsoft\Windows\CurrentVersion\App Paths\" & AppName

TheResult = RegOpenKeyEx(MainKey, SubKey, 0, KEY_ALL_ACCESS, mKey)

If TheResult <> 0 Then Exit Function
'looked for it and failed

Index = 0
Do
EntryLength = 1024
DataLength = 1024
TheEntry = Space(EntryLength)
TheResult = RegEnumValue(mKey, Index, TheEntry, EntryLength, 0, _
     TheDataType, TheByteArray(1), DataLength)
'looks like we just have to pass just the first element
'of the array to have it filled...
If TheResult <> 0 Then Exit Do

TheEntry = Left(TheEntry, EntryLength)

If Len(TheEntry) = 0 Then
'looking for (Default), empty string

    ByteValue = ""
    For i = 1 To DataLength - 1
    ByteValue = ByteValue & Chr(TheByteArray(i))
    Next
    
    If Len(ByteValue) > 0 Then
    GetAppPath = LongName(ByteValue)
    RegCloseKey mKey
    Exit Function
    End If

End If
Index = Index + 1
Loop
GetAppPath = ""
RegCloseKey mKey
Exit Function
TheEnd:
GetAppPath = ""
End Function

Public Function ParseAndShell(TheCommand As String) As String
'Returns detailed path and error information for debugging
Dim Scr_hDC As Long
Dim StartDoc As Long
Dim Parameters As String
On Error Resume Next
'Important: assumes all parameters are in quotes
'to use quotes in code use double quotes
If InStr(1, TheCommand, """") <> 0 Then
Parameters = Right(TheCommand, Len(TheCommand) - InStr(1, TheCommand, """"))
TheCommand = Left(TheCommand, Len(TheCommand) - Len(Parameters) - 2)
Parameters = Left(Parameters, Len(Parameters) - 1)
End If

If Len(Dir(TheCommand, 55)) = 0 Then
    If Len(EnvironPath(TheCommand)) > 0 Then
    TheCommand = EnvironPath(TheCommand)
    ElseIf Len(GetAppPath(TheCommand)) > 0 Then
    TheCommand = GetAppPath(TheCommand)
    End If
End If
Scr_hDC = GetDesktopWindow()
StartDoc = ShellExecute(Scr_hDC, "Open", TheCommand, _
    Parameters, "C:\", SW_SHOWNORMAL)

If Len(Dir(Parameters, 39)) > 0 Then Parameters = PrettyPath(LongName(Parameters))
ParseAndShell = TheCommand & " " & Parameters & " : " & ErrMessage(StartDoc)

End Function
Private Function ErrMessage(r As Long) As String
Select Case r 'for ShellExecute
Case 0
    ErrMessage = "Out of memory"
Case 1
    ErrMessage = "Operation successful"
Case Is > 32
    ErrMessage = "Operation successful"
Case SE_ERR_FNF
    ErrMessage = "File not found"
Case SE_ERR_PNF
    ErrMessage = "Path not found"
Case SE_ERR_ACCESSDENIED
    ErrMessage = "Access denied"
Case SE_ERR_OOM
    ErrMessage = "Out of memory"
Case SE_ERR_DLLNOTFOUND
    ErrMessage = "DLL not found"
Case SE_ERR_SHARE
    ErrMessage = "A sharing violation occurred"
Case SE_ERR_ASSOCINCOMPLETE
    ErrMessage = "Incomplete or invalid file association"
Case SE_ERR_DDETIMEOUT
    ErrMessage = "DDE Time out"
Case SE_ERR_DDEFAIL
    ErrMessage = "DDE transaction failed"
Case SE_ERR_DDEBUSY
    ErrMessage = "DDE busy"
Case SE_ERR_NOASSOC
    ErrMessage = "No association for file extension"
Case ERROR_BAD_FORMAT
    ErrMessage = "Invalid EXE file or error in EXE image"
Case Else
    ErrMessage = "Unknown error"
End Select

End Function



Public Function PrettyName(TheName As String) As String
On Error GoTo TheEnd:
'This is designed for human names not path names.
'Middle initial not supported but could be added.
Dim Name As String
Dim Spot As Integer
Dim First As String
Dim Last As String
Dim i As Integer

Name = TheName
Name = Replace(Name, "    ", " ", , , vbTextCompare)
Name = Replace(Name, "   ", " ", , , vbTextCompare)
Name = Replace(Name, "  ", " ", , , vbTextCompare)

First = Left(Name, 1)
Name = LCase(Name)
First = UCase(First)


If InStr(1, Name, ",") > 0 Then
Spot = InStr(1, Name, ",")
Last = Mid(Name, Spot + 1, 1)
    Do
        i = i + 1
        If Last = " " Then
        Last = Mid(Name, Spot + i, 1)
        End If
    Loop Until Last <> " " Or i = 3
Last = UCase(Last)
Name = First & Mid(Name, 2, Spot - 2 + i) & Last & Mid(Name, Spot + 1 + i, Len(Name) - Spot + i)
PrettyName = Name

Else
    If InStr(1, Name, " ") = 0 Then
    Name = First & Mid(Name, 2)
    PrettyName = Name
    Else
    Spot = InStr(1, Name, " ")
    Last = Mid(Name, Spot + 1, 1)
    Last = UCase(Last)
    Name = First & Mid(Name, 2, Spot - 1) & Last & Mid(Name, Spot + 2, Len(Name) - Spot + 1)
    PrettyName = Name
    End If
End If
Exit Function
TheEnd:
PrettyName = TheName
End Function

Public Function PrettyPath(ThePath As String) As String

On Error GoTo TheEnd:

Dim Path As String
Dim Start As Integer
Dim temp As String

Path = ThePath
Path = LCase(Path)

temp = Left(Path, 1)
temp = UCase(temp)
Path = temp & Right(Path, Len(Path) - 1)
'got drive letter pretty

Start = 1
Do
Start = InStr(Start, Path, "\")
If Start = 0 Then Exit Do
Mid(Path, Start + 1, 1) = UCase(Mid(Path, Start + 1, 1))
Start = Start + 1
Loop While Start < Len(ThePath)
'put a cap after each backslash

Start = 1
Do
Start = InStr(Start, Path, " ")
If Start = 0 Then Exit Do
Mid(Path, Start + 1, 1) = UCase(Mid(Path, Start + 1, 1))
Start = Start + 1
Loop While Start < Len(Path)
'put a cap after each space

PrettyPath = Path
Exit Function
TheEnd:
PrettyPath = ThePath
'just in case
End Function

Public Function EnvironPath(AppName As String) As String
'gets full path from environment if it exists
'returns vbNullstring if not
On Error GoTo TheEnd:
Dim i As Integer
Dim EnvString As String
Dim Test As String
Dim Paths() As String
Dim Start As Long
Dim length As Long
Dim UpPaths As Integer

EnvString = Environ("Path")
EnvString = EnvString & ";" 'to tell us where the end is

Start = 1
length = InStr(1, EnvString, ";", vbBinaryCompare) - 1
Do 'start parsing the Path environment
i = i + 1
ReDim Preserve Paths(i)
Paths(i) = Mid(EnvString, Start, length)
Start = Start + Len(Paths(i)) + 1
length = InStr(Start, EnvString, ";", vbBinaryCompare) - Start
Loop While length > 0
UpPaths = UBound(Paths)
For i = 1 To UpPaths
    If InStr(1, AppName, ".") = 0 Then
        If Dir(Paths(i) & "\" & AppName & ".exe", 39) <> "" Then
        EnvironPath = Paths(i) & "\" & AppName & ".exe"
        ElseIf Dir(Paths(i) & "\" & AppName & ".com", 39) <> "" Then
        EnvironPath = Paths(i) & "\" & AppName & ".com"
        ElseIf Dir(Paths(i) & "\" & AppName & ".bat", 39) <> "" Then
        EnvironPath = Paths(i) & "\" & AppName & ".bat"
        ElseIf Dir(Paths(i) & "\" & AppName & ".pif", 39) <> "" Then
        EnvironPath = Paths(i) & "\" & AppName & ".pif"
        ElseIf Dir(Paths(i) & "\" & AppName & ".scr", 39) <> "" Then
        EnvironPath = Paths(i) & "\" & AppName & ".scr"
        End If
    Else
        If Dir(Paths(i) & "\" & AppName, 39) <> "" Then
        EnvironPath = Paths(i) & "\" & AppName
        End If
    End If
Next
EnvironPath = PrettyPath(EnvironPath)
Exit Function
TheEnd:

End Function


