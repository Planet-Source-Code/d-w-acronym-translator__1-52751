Attribute VB_Name = "Find"
Option Explicit

Private Declare Function FindText Lib "comdlg32.dll" Alias "FindTextA" (pFindreplace As Long) As Long
Private Declare Function ReplaceText Lib "comdlg32.dll" Alias "ReplaceTextA" (pFindreplace As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Private Declare Function IsDialogMessage Lib "user32" Alias "IsDialogMessageA" (ByVal hDlg As Long, lpMsg As Msg) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CopyPointer2String Lib "kernel32" Alias "lstrcpyA" (ByVal NewString As String, ByVal OldString As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetProcessHeap& Lib "kernel32" ()
Private Declare Function HeapAlloc& Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long)
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare Function EndDialog Lib "user32" (ByVal hDlg As Long, ByVal nResult As Long) As Long


Private Type FINDREPLACE
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    flags As Long
    lpstrFindWhat As Long
    lpstrReplaceWith As Long
    wFindWhatLen As Integer
    wReplaceWithLen As Integer
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type Msg
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    ptX As Long
    ptY As Long
End Type

Private Const GWL_WNDPROC = (-4)
Private Const HEAP_ZERO_MEMORY = &H8
Public Const FR_HIDEUPDOWN = &H4000
Public Const FR_HIDEWHOLEWORD = &H10000
Public Const FR_HIDEMATCHCASE = &H8000
Public Const FR_NOMATCHCASE = &H800
Public Const FR_NOUPDOWN = &H400
Public Const FR_NOWHOLEWORD = &H1000
Public Const FR_REPLACE = &H10
Public Const FR_WHOLEWORD = &H2
Public Const FR_SHOWHELP = &H80
Public Const FR_FINDNEXT = &H8
Public Const FR_ENABLETEMPLATEHANDLE = &H2000
Public Const FR_ENABLETEMPLATE = &H200
Public Const FR_DIALOGTERM = &H40
Public Const FR_DOWN = &H1
Public Const FR_ENABLEHOOK = &H100

Private Const WM_DESTROY = &H2
Private Const FINDMSGSTRING = "commdlg_FindReplace"
Private Const HELPMSGSTRING = "commdlg_help"
Private Const BufLength = 256

Public hDialog As Long
Public OldProc As Long
Public RetFrs As FINDREPLACE
Public TMsg As Msg
Public Up As Boolean
Public TheTextBox As TextBox
Public FindStarted As Boolean
Private LastFind As String
Public DefSearch As Boolean
Private PrevSearch As Boolean
Private lFlags As Long
Private sFind As String
Private uFindMsg As Long
Private uHelpMsg As Long
Private lHeap As Long
Private arrFind() As Byte
Private arrReplace() As Byte

Private Sub StartFind()

Dim TRS As Recordset
Dim i As Long
Dim FieldName As String
If DefSearch Then
FieldName = "Definition"
Else
FieldName = "Acronym"
End If
Main.LeftList.Clear
Main.RightList.Clear

If Rated Then
Set TRS = TDB.OpenRecordset("SELECT * FROM " & TableName _
    & " WHERE Rating = '0' ORDER BY Acronym")
Else
Set TRS = TDB.OpenRecordset("SELECT * FROM " & TableName _
    & " ORDER BY Acronym")
End If
For i = 1 To TRS.RecordCount
    If InStr(1, UCase(TRS.Fields(FieldName)), UCase(sFind)) > 0 Then
    Main.LeftList.AddItem TRS.Fields("Acronym")
    Main.RightList.AddItem TRS.Fields("Definition")
    End If
If Not TRS.EOF Then TRS.MoveNext
Next

If Main.LeftList.ListCount > 0 Then
Main.LeftList.ListIndex = 0
End If

ReadingAll = False
LastSearch = sFind
FindSearch = True
End Sub
Public Function FindNextWord() As Boolean
If Main.RightList.ListCount > Selection + 1 Then
Main.RightList.ListIndex = Selection + 1
Selection = Main.RightList.ListIndex
FindNextWord = True
Else
FindNextWord = False
MsgBox "Cannot find """ & sFind & """", vbInformation
End If
End Function

Public Function FindPrevWord() As Boolean
If Selection > 0 Then
Main.RightList.ListIndex = Selection - 1
Selection = Main.RightList.ListIndex
FindPrevWord = True
Else
FindPrevWord = False
MsgBox "Cannot find """ & sFind & """", vbInformation
End If
End Function

Public Sub ShowFind(fOwner As Form, ATextBox As TextBox)
Dim FRS As FINDREPLACE
Dim i As Integer
Set TheTextBox = ATextBox
sFind = TheTextBox.SelText
If Len(sFind) > 30 Then
sFind = Left(sFind, 30)
End If
lFlags = FR_DOWN Or FR_HIDEWHOLEWORD Or FR_NOMATCHCASE
If hDialog > 0 Then Exit Sub
arrFind = StrConv(sFind & Chr$(0), vbFromUnicode)
With FRS
    .lStructSize = LenB(FRS)
    .lpstrFindWhat = VarPtr(arrFind(0))
    .wFindWhatLen = BufLength
    .hwndOwner = fOwner.hWnd
    .flags = lFlags
    .hInstance = App.hInstance
End With
lHeap = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, FRS.lStructSize)
CopyMemory ByVal lHeap, FRS, Len(FRS)
uFindMsg = RegisterWindowMessage(FINDMSGSTRING)
uHelpMsg = RegisterWindowMessage(HELPMSGSTRING)
OldProc = SetWindowLong(fOwner.hWnd, GWL_WNDPROC, AddressOf WndProc)
hDialog = FindText(ByVal lHeap)
MessageLoop
End Sub

Private Sub MessageLoop()
Do While GetMessage(TMsg, 0&, 0&, 0&) And hDialog > 0
    If IsDialogMessage(hDialog, TMsg) = False Then
    TranslateMessage TMsg
    DispatchMessage TMsg
    End If
Loop
End Sub

Public Function WndProc(ByVal hOwner As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Select Case wMsg
      Case uFindMsg
           CopyMemory RetFrs, ByVal lParam, Len(RetFrs)
           If (RetFrs.flags And FR_DIALOGTERM) = FR_DIALOGTERM Then
              SetWindowLong hOwner, GWL_WNDPROC, OldProc
              HeapFree GetProcessHeap(), 0, lHeap
              hDialog = 0
              lHeap = 0
              OldProc = 0
              FindStarted = False
              sFind = ""
              Erase arrFind
           Else
              DoFind RetFrs
           End If
      
      Case Else
           If wMsg = WM_DESTROY Then
              EndDialog hDialog, 0&
              SetWindowLong hOwner, GWL_WNDPROC, OldProc
              HeapFree GetProcessHeap(), 0, lHeap
              hDialog = 0
              lHeap = 0
              OldProc = 0
              FindStarted = False
              sFind = ""
              Erase arrFind
              Exit Function
           End If
           WndProc = CallWindowProc(OldProc, hOwner, wMsg, wParam, lParam)
   End Select
End Function

Private Sub DoFind(FR As FINDREPLACE)

If CheckFlags(FR_FINDNEXT, FR.flags) Then
sFind = PointerToString(FR.lpstrFindWhat)
    
    If sFind <> LastFind Then
    FindStarted = False
    LastFind = sFind
    End If

    If PrevSearch <> DefSearch Then
    FindStarted = False
    PrevSearch = DefSearch
    End If

lFlags = FR.flags

    If Not FindStarted Then
    StartFind
    FindStarted = True
    Exit Sub
    End If

    If CheckFlags(FR_DOWN, FR.flags) Then
    FindNextWord
    Else
    FindPrevWord
    End If
End If
End Sub


Private Function PointerToString(p As Long) As String
Dim s As String
s = String(BufLength, Chr$(0))
CopyPointer2String s, p
PointerToString = Left(s, InStr(s, Chr$(0)) - 1)
End Function

Private Function CheckFlags(flag As Long, flags As Long) As Boolean
CheckFlags = ((flags And flag) = flag)
End Function

Private Function ValidateWholeWord(PrevLetter As String, NextLetter As String) As Boolean
Dim sLetters As String
ValidateWholeWord = True
sLetters = "abcdefghijklmnoprqstuvwxyz1234567890"
If InStr(1, sLetters, PrevLetter, vbTextCompare) Or InStr(1, sLetters, NextLetter, vbTextCompare) Then ValidateWholeWord = False
End Function

Private Function InStrR(Optional lStart As Long, Optional sTarget As String, Optional sFind As String, Optional iCompare As Integer) As Long
Dim cFind As Long
Dim i As Long
cFind = Len(sFind)
For i = lStart - cFind + 1 To 1 Step -1
    If StrComp(Mid$(sTarget, i, cFind), sFind, iCompare) = 0 Then
    InStrR = i
    Exit Function
    End If
Next
End Function
