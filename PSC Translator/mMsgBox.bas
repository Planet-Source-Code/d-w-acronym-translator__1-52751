Attribute VB_Name = "mMsgBox"
Option Explicit

Private Const HCBT_ACTIVATE = 5
Private Const WH_CBT = 5
Private Const HCBT_DESTROYWND = 4
Private Const AppEXEName = "TRANSLATOR"
Private Declare Function SetWindowsHookEx Lib "user32" Alias _
  "SetWindowsHookExA" (ByVal IDHook As Long, ByVal lpfn As Long, _
  ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const WS_VISIBLE = &H10000000
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private m_HookID As Long
Private m_NoButton As Boolean
Private m_Captions() As Variant
Private m_CaptionFound As Boolean
Private m_Icon As PictureBox
Private m_PrevWnd As Long

Dim i As Integer

Public Function MsgBox(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
  Optional Title As String = AppEXEName, Optional HelpFile, Optional Context, _
  Optional Captions As Variant, _
  Optional Icon As PictureBox, _
  Optional NoButton As Boolean = False) As VbMsgBoxResult
  Set Icon = Main.Pic
  'Save the parameter cause we will call out enumerate function
  If Not IsMissing(Captions) Then
    If IsArray(Captions) Then
      m_Captions = Captions
      m_CaptionFound = True
    End If
  End If
  Set m_Icon = Icon
  If Not m_Icon Is Nothing Then
    m_PrevWnd = m_Icon.Parent.hWnd
  End If
  m_NoButton = NoButton
  'Place a hook to catch msgbox activation
  SubClassMsgBox
  'Call the intrinsic VB Msgbox and get the retval
  MsgBox = VBA.MsgBox(Prompt, Buttons, Title, HelpFile, Context)
End Function

Public Sub SubClassMsgBox()
  If m_HookID = 0& Then
    m_HookID = SetWindowsHookEx(WH_CBT, AddressOf MsgBoxProc, 0&, GetCurrentThreadId())
  Else
    RemoveMsgBoxHook
  End If
End Sub

Private Function MsgBoxProc(ByVal lMsg As Long, ByVal wParam As Long, _
   ByVal lParam As Long) As Long

  Select Case lMsg
    Case HCBT_ACTIVATE
    'wparam is the MsgBox Wnd handle
    'Now modify the msgbox
      EnumChildWindows wParam, AddressOf EnumChildProc, ByVal 0&
    Case HCBT_DESTROYWND
    'remove the hook when msgbox closed
      RemoveMsgBoxHook
  End Select
  'return 0
  MsgBoxProc = False
End Function

Public Sub RemoveMsgBoxHook()
  If m_HookID <> 0& Then
    UnhookWindowsHookEx m_HookID
    m_HookID = 0&
    Erase m_Captions
    m_CaptionFound = False
    If Not m_Icon Is Nothing Then
      'restore our icon to it's previous parent or else it will also get destroy
      'along with msgbox
      SetParent m_Icon.hWnd, m_PrevWnd
      m_Icon.Visible = False
      m_PrevWnd = 0
    End If
    Set m_Icon = Nothing
    i = 0
  End If
End Sub

Private Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
  Dim sTemp As String
  Dim lRetval As Long
  Dim sClass As String
  Dim hdc As Long
  
  'Get the window class name
  
  sClass = GetWndName(hWnd)
  
  Select Case sClass
    Case "Button"
      'This child is a button
      If m_NoButton Then
        'A better way is just reduce the msgbox height so the button won't be visible
        'so we won't have to enumerate the buttons
        
        'Set visible to false
        
        SetWindowLong hWnd, GWL_STYLE, GetWindowLong(hWnd, GWL_STYLE) And Not WS_VISIBLE
      Else
      'Change Button's caption
        If m_CaptionFound Then
          'Get the button caption
          On Error Resume Next
          SetWindowText hWnd, m_Captions(i)
          i = i + 1
        End If
        
      End If
      
    Case "Static"
      If Not m_Icon Is Nothing Then
        sTemp = String(255, Chr(0))
        lRetval = GetWindowText(hWnd, sTemp, 255)
        'For win98 the Icon does not have a blank caption
        If lRetval > 0 Then
          'Get the first character
          sTemp = Left$(sTemp, 1)
        Else
          sTemp = "o" 'just fill it with anything you like we only need this for comparison
        End If
        
        If lRetval = 0 Or Asc(sTemp) = 255 Then
          'This is the icon so hide it
          SetWindowLong hWnd, GWL_STYLE, GetWindowLong(hWnd, GWL_STYLE) And Not WS_VISIBLE
          'Mpve our icon to msgbox and show it
          Call SetParent(m_Icon.hWnd, GetParent(hWnd))
          m_Icon.Move 120, 120
          m_Icon.Visible = True
        End If
      End If
  End Select
  'continue enumeration
  EnumChildProc = 1
End Function

Private Function GetWndName(hWnd As Long) As String
  Dim lRetval As Long
  GetWndName = String(255, Chr(0))
  lRetval = GetClassName(hWnd, GetWndName, 255)
  GetWndName = VBA.Left$(GetWndName, lRetval)
End Function


