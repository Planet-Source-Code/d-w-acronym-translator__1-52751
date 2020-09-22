Attribute VB_Name = "SubMain"
Option Explicit

Private Declare Function GetFocus Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" _
    Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32.dll" _
    Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32.dll" Alias _
    "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
    (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, _
    Source As Any, ByVal length As Long)
Private Declare Function LocalAlloc Lib "kernel32" _
    (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" _
    (ByVal hMem As Long) As Long
Private Declare Function lstrcpyA Lib "kernel32" _
    (lpString1 As Any, lpString2 As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" _
    (lpString As Any) As Long
Public Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SHBrowseForFolder Lib "Shell32" _
    (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32" _
    (ByVal pidList As Long, _
    ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
    (ByVal lpString1 As String, ByVal _
    lpString2 As String) As Long
Public Declare Function GetSystemMetrics Lib "user32" ( _
     ByVal nIndex As Long) As Long
Public Declare Function SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CXSCREEN = 0
Public Const SM_CXFULLSCREEN = 16
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const lPtr = (LMEM_FIXED Or LMEM_ZEROINIT)
Private Const BFFM_INITIALIZED = 1
Public Const WM_USER As Long = &H400
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Public Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
Private Const ES_NUMBER = &H2000&
Private Const GWL_STYLE = (-16)
Private Const ES_UPPERCASE = &H8
Private Const ES_LOWERCASE = &H10
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115
Private Const SB_LINEUP = 0
Private Const SB_LINELEFT = 0
Private Const SB_LINEDOWN = 1
Private Const SB_LINERIGHT = 1
Private Const SB_PAGEUP = 2
Private Const SB_PAGELEFT = 2
Private Const SB_PAGEDOWN = 3
Private Const SB_PAGERIGHT = 3
Private Const SB_THUMBPOSITION = 4
Private Const SB_THUMBTRACK = 5
Private Const SB_TOP = 6
Private Const SB_LEFT = 6
Private Const SB_BOTTOM = 7
Private Const SB_RIGHT = 7
Private Const SB_ENDSCROLL = 8

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Enum FieldType
    dbBoolean = 1
    dbByte = 2
    dbInteger = 3
    dbLong = 4
    dbCurrency = 5
    dbSingle = 6
    dbDouble = 7
    dbDate = 8
    dbBinary = 9
    dbText = 10
    dbLongBinary = 11
    dbMemo = 12
    dbGUID = 15
    dbBigInt = 16
    dbVarBinary = 17
    dbChar = 18
    dbNumeric = 19
    dbDecimal = 20
    dbFloat = 21
    dbTime = 22
    dbTimeStamp = 23
End Enum

Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Public Enum Folder
    Windows = vbNull
    WINSYSTEM = -1
    DESKTOP = 0
    PROGRAMS = 2
    Documents = 5
    FAVORITES = 6
    STARTUP = 7
    RECENT = 8
    SENDTO = 9
    STARTMENU = 11
    DESKTOPUSER = 16
    NETHOOD = 19
    FONTFOLDER = 20
    SHELLNEW = 21
    APPDATA = 26
    PRINTHOOD = 27
    TEMPINTERNET = 32
    COOKIES = 33
    HISTORY = 34
    temp = 99

End Enum
Public TDB As Database
Public TableName As String
Public Rated As Boolean
Public LastSearch As String
Public ReadingAll As Boolean
Public FindSearch As Boolean
Public Selection As Long
Public Sub AddTable(TableName As String)
Dim i As Integer
Dim FieldSize As Integer
Dim Find As TableDef
Dim TTbl As TableDef
Dim TFld1 As Field
Dim TFld2 As Field
Dim TFld3 As Field
Set Find = New TableDef
    For Each Find In TDB.TableDefs
        If LCase(Find.Name) = LCase(TableName) Then
        Set Find = Nothing
        GoTo Skip:
        End If
    Next
Set TTbl = TDB.CreateTableDef(TableName)
Set TFld1 = TTbl.CreateField("Rating", dbText, 255)
TFld1.AllowZeroLength = True
Set TFld2 = TTbl.CreateField("Acronym", dbText, 255)
TFld2.AllowZeroLength = True
Set TFld3 = TTbl.CreateField("Definition", dbText, 255)
TFld3.AllowZeroLength = True
TTbl.Fields.Append TFld1
TTbl.Fields.Append TFld2
TTbl.Fields.Append TFld3
TDB.TableDefs.Append TTbl
Exit Sub
Skip:
MsgBox "Cannot add table because it already exists.", vbOKOnly + vbInformation
End Sub
Public Function CreateBlankDB()
Set TDB = DBEngine.Workspaces(0).CreateDatabase(App.Path & "\Acronyms.mdb", dbLangGeneral)
SaveSetting "Translator", "Paths", "Data", App.Path & "\Acronyms.mdb"
OpenData
End Function


Public Sub UpperOnly(TheBox As TextBox)
Dim lValue As Long
Dim lAlign As Long
Dim lReturn As Long
lAlign = ES_UPPERCASE
lValue = GetWindowLong(TheBox.hWnd, GWL_STYLE)
lReturn = SetWindowLong(TheBox.hWnd, GWL_STYLE, lValue Or lAlign)
TheBox.Refresh
End Sub




Public Sub CloseDatabase()
On Error Resume Next
TDB.Close
Set TDB = Nothing
On Error GoTo 0
End Sub
Public Function SelectedFolder(frm As Form, sPath As String)
Dim IdList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo
Dim lpPath As Long

szTitle = "Select folder for Translator files. " _
& "This cannot be the program folder. You may need to " _
& "create a New Folder from Explorer."
'If you use a windows installer package, they are self repairing
'and the repair process deletes the entire app.path folder and
'replaces it with original installation and you will lose your changes
'to the database. It inventories this folder each time program is run and
'initiates repair if just one file is missing, even unnecessary ones.
With tBrowseInfo
    .hwndOwner = frm.hWnd
    .pIDLRoot = 0
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    .lpfnCallback = FarProc(AddressOf BrowseCallbackProcStr)
    lpPath = LocalAlloc(lPtr, Len(sPath) + 1)
    CopyMemory ByVal lpPath, ByVal sPath, Len(sPath) + 1
    .lParam = lpPath
End With

IdList = SHBrowseForFolder(tBrowseInfo)

If IdList Then
sBuffer = Space(260)
SHGetPathFromIDList IdList, sBuffer
sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
SelectedFolder = sBuffer
Else
SelectedFolder = ""
End If
End Function

Public Function BrowseCallbackProcStr(ByVal hWnd As Long, _
    ByVal uMsg As Long, ByVal lParam As Long, _
    ByVal lpData As Long) As Long
'used by SelectedFolder function
Select Case uMsg
Case BFFM_INITIALIZED
SendMessage hWnd, BFFM_SETSELECTIONA, True, ByVal lpData
Case Else
End Select

End Function
Public Function FarProc(pfn As Long) As Long
FarProc = pfn 'to avoid a direct api call
End Function
Public Sub OpenData()

Dim DataPath As String

DataPath = GetSetting("Translator", "Paths", "Data", "")
If Len(DataPath) = 0 Then
SetDataPath
End If
DataPath = GetSetting("Translator", "Paths", "Data", "")
On Error GoTo Out:
If Len(DataPath) = 0 Or Dir(DataPath) = "" Then GoTo Out:

Set TDB = OpenDatabase(DataPath)
On Error GoTo 0
Exit Sub
Out:
On Error Resume Next
Kill TheIni
If MsgBox("Database removed, renamed or corrupted.", vbRetryCancel + vbCritical, "DATA SOURCE REQUIRED") = vbRetry Then
OpenData
Else
End
End If

End Sub

Public Sub SetDataPath()

Dim TheFile As String
Dim DataPath As String

DataPath = GetSetting("Translator", "Paths", "Data", "")
If DataPath = "" Or Dir(DataPath) = "" Then
    If Dir(App.Path & "\Acronyms.mdb") <> "" Then
    DataPath = App.Path & "\Acronyms.mdb"
    SaveSetting "Translator", "Paths", "Data", DataPath
    GoTo Skip:
    ElseIf Dir(FolderPath(App.Path) & "\Acronyms.mdb") <> "" Then
    DataPath = FolderPath(App.Path) & "\Acronyms.mdb"
    SaveSetting "Translator", "Paths", "Data", DataPath
    GoTo Skip:
    End If
End If

On Error GoTo Out:
    With Main.Filebox
    .FileName = IIf(Len(DataPath) = 0, "", NameFromPath(DataPath))
    .InitDir = IIf(Len(DataPath) > 0, DataPath, FolderPath(App.Path))
    .DialogTitle = "Set Data Source"
    .Filter = "Acronyms Database (Acronyms.mdb)|Acronyms.mdb"
    .flags = cdlOFNHideReadOnly
    .CancelError = True
    .ShowOpen
    End With

TheFile = Main.Filebox.FileName


If Len(TheFile) > 0 And Len(Dir(TheFile)) > 0 Then
SaveSetting "Translator", "Paths", "Data", TheFile
End If
Skip:
On Error Resume Next
TDB.Close
On Error GoTo 0
OpenData
Exit Sub
Out:

End Sub

Public Sub SaveSetting(AppName As String, Section As String, Key As String, Setting As String)
    'AppName not used here it's just for cod
    '     e compatability.
    'TheIni function contains the path to th
    '     e ini which serves the same purpose.
    WritePrivateProfileString Section, Key, Setting, TheIni
End Sub



Public Function GetSetting(AppName As String, Section As String, Key As String, Optional Default As String = "") As String
    'AppName not used here it's just for cod
    '     e compatability
    'TheIni function contains the path to th
    '     e ini which serves the same purpose
    Dim StringLength As Long
    Dim Buffer As String * 155
    StringLength = GetPrivateProfileString(Section, _
    Key, Default, Buffer, Len(Buffer), TheIni)
    GetSetting = Left(Buffer, StringLength)
End Function


Public Function TheIni() As String
TheIni = App.Path & "\" & App.EXEName & ".ini"
End Function


Public Function SpecialFolder(Optional TheFolder As Folder = vbNull) As String
Dim ID As ITEMIDLIST
Dim LngRet As Long
Dim ThePath As String
Dim TheLength As Long
ThePath = Space(255)
Select Case TheFolder
Case Windows
TheLength = GetWindowsDirectory(ThePath, 255)
ThePath = Left(ThePath, TheLength)

Case WINSYSTEM
TheLength = GetSystemDirectory(ThePath, 255)
ThePath = Left(ThePath, TheLength)

Case temp
TheLength = GetTempPath(255, ThePath)
ThePath = Left(ThePath, TheLength - 1)

Case Else
LngRet = SHGetSpecialFolderLocation(0, TheFolder, ID)
    If LngRet = 0 Then
        If SHGetPathFromIDList(ID.mkid.cb, ThePath) <> 0 Then
        ThePath = Left(ThePath, InStr(ThePath, vbNullChar) - 1)
        End If
    End If
End Select
SpecialFolder = Trim(ThePath)
End Function


