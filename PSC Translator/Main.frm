VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Main 
   Caption         =   "Acronym Translator"
   ClientHeight    =   5580
   ClientLeft      =   1650
   ClientTop       =   660
   ClientWidth     =   6585
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Main"
   ScaleHeight     =   5580
   ScaleWidth      =   6585
   Begin VB.Timer FocusTimer 
      Interval        =   100
      Left            =   4680
      Top             =   1935
   End
   Begin VB.PictureBox TopPane 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   0
      ScaleHeight     =   1185
      ScaleWidth      =   6585
      TabIndex        =   4
      Top             =   0
      Width           =   6585
      Begin VB.CommandButton CopyButton 
         Caption         =   "c"
         Height          =   255
         Index           =   0
         Left            =   1170
         TabIndex        =   14
         ToolTipText     =   "Copy Identifier To Clipboard"
         Top             =   15
         Width           =   240
      End
      Begin VB.CommandButton CopyButton 
         Caption         =   "c"
         Height          =   255
         Index           =   3
         Left            =   6330
         TabIndex        =   13
         ToolTipText     =   "Copy Definition To Clipboard"
         Top             =   645
         Width           =   240
      End
      Begin VB.CommandButton CopyButton 
         Caption         =   "c"
         Height          =   255
         Index           =   2
         Left            =   6330
         TabIndex        =   12
         ToolTipText     =   "Copy Definition To Clipboard"
         Top             =   337
         Width           =   240
      End
      Begin VB.CommandButton CopyButton 
         Caption         =   "c"
         Height          =   255
         Index           =   1
         Left            =   6330
         TabIndex        =   11
         ToolTipText     =   "Copy Definition To Clipboard"
         Top             =   30
         Width           =   240
      End
      Begin VB.TextBox Results 
         Height          =   285
         Index           =   3
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   615
         Width           =   4710
      End
      Begin VB.TextBox Results 
         Height          =   285
         Index           =   2
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   315
         Width           =   4710
      End
      Begin VB.TextBox Results 
         Height          =   285
         Index           =   1
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   15
         Width           =   4710
      End
      Begin VB.TextBox SearchBox 
         Height          =   285
         Left            =   90
         TabIndex        =   7
         ToolTipText     =   "Type first few letters of acronym..."
         Top             =   0
         Width           =   1005
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "E&dit Entry"
         Enabled         =   0   'False
         Height          =   300
         Left            =   135
         TabIndex        =   42
         Top             =   615
         Width           =   945
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New Entry"
         Height          =   300
         Left            =   135
         TabIndex        =   41
         Top             =   300
         Width           =   945
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   300
         Left            =   135
         TabIndex        =   44
         Top             =   300
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "C&ancel"
         Height          =   300
         Left            =   135
         TabIndex        =   43
         Top             =   615
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Image Images 
         Height          =   240
         Index           =   4
         Left            =   6330
         Picture         =   "Main.frx":030A
         Top             =   660
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Images 
         Height          =   240
         Index           =   3
         Left            =   6330
         Picture         =   "Main.frx":064C
         Top             =   345
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Images 
         Height          =   240
         Index           =   2
         Left            =   6330
         Picture         =   "Main.frx":098E
         Top             =   30
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Images 
         Height          =   240
         Index           =   1
         Left            =   1170
         Picture         =   "Main.frx":0CD0
         Top             =   45
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Letter 
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   25
         Left            =   3585
         TabIndex        =   40
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   24
         Left            =   3480
         TabIndex        =   39
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   23
         Left            =   3330
         TabIndex        =   38
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   22
         Left            =   3150
         TabIndex        =   37
         Top             =   930
         Width           =   165
      End
      Begin VB.Label Letter 
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   21
         Left            =   3006
         TabIndex        =   36
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   20
         Left            =   2865
         TabIndex        =   35
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   19
         Left            =   2724
         TabIndex        =   34
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   18
         Left            =   2583
         TabIndex        =   33
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   17
         Left            =   2442
         TabIndex        =   32
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "Q"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   16
         Left            =   2301
         TabIndex        =   31
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   15
         Left            =   2160
         TabIndex        =   30
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   14
         Left            =   2019
         TabIndex        =   29
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   13
         Left            =   1878
         TabIndex        =   28
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   12
         Left            =   1737
         TabIndex        =   27
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   11
         Left            =   1596
         TabIndex        =   26
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   10
         Left            =   1455
         TabIndex        =   25
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   9
         Left            =   1314
         TabIndex        =   24
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   8
         Left            =   1173
         TabIndex        =   23
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   7
         Left            =   1032
         TabIndex        =   22
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   6
         Left            =   891
         TabIndex        =   21
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   5
         Left            =   750
         TabIndex        =   20
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   4
         Left            =   609
         TabIndex        =   19
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   468
         TabIndex        =   18
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   327
         TabIndex        =   17
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   186
         TabIndex        =   16
         Top             =   930
         Width           =   135
      End
      Begin VB.Label Letter 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   15
         Top             =   930
         Width           =   135
      End
   End
   Begin MSComDlg.CommonDialog Filebox 
      Left            =   4065
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Pic 
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   4110
      Picture         =   "Main.frx":1012
      ScaleHeight     =   525
      ScaleWidth      =   510
      TabIndex        =   3
      Top             =   1860
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox RightPane 
      Align           =   3  'Align Left
      BackColor       =   &H80000005&
      Height          =   4395
      Left            =   1500
      ScaleHeight     =   4335
      ScaleWidth      =   1395
      TabIndex        =   2
      Top             =   1185
      Width           =   1455
      Begin VB.ListBox RightList 
         Height          =   2220
         IntegralHeight  =   0   'False
         Left            =   0
         TabIndex        =   6
         Top             =   -30
         Width           =   1215
      End
   End
   Begin VB.PictureBox SplitterBar 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   4395
      Left            =   1455
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4395
      ScaleWidth      =   45
      TabIndex        =   1
      Tag             =   "50"
      ToolTipText     =   "Click and drag to resize."
      Top             =   1185
      Width           =   45
   End
   Begin VB.PictureBox LeftPane 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   4395
      Left            =   0
      ScaleHeight     =   4335
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   1185
      Width           =   1455
      Begin VB.ListBox LeftList 
         Height          =   2220
         IntegralHeight  =   0   'False
         Left            =   0
         TabIndex        =   5
         Top             =   -15
         Width           =   1140
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Entry"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "&Cancel"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuChange 
         Caption         =   "&Change Entry"
      End
      Begin VB.Menu mnuRated 
         Caption         =   "&Show X-Rated"
      End
   End
   Begin VB.Menu mnuCat 
      Caption         =   "&Categories"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add New..."
      End
      Begin VB.Menu mnuCategory 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuFind 
      Caption         =   "&Search"
      Begin VB.Menu mnuDef 
         Caption         =   "&Definitions"
      End
      Begin VB.Menu mnuAcro 
         Caption         =   "&Acronyms"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BarPosition As Integer
Dim Selecting As Boolean
Dim NewRecord As Boolean


Private Sub EditMode()
Dim Ctrl As Control
Dim i As Integer
Selecting = True
For Each Ctrl In Main
If TypeOf Ctrl Is ListBox Or TypeOf Ctrl Is Label Then
Ctrl.Enabled = False
End If
Next
For i = 0 To 3
CopyButton(i).Visible = False
Next
For i = 1 To 3
Results(i).Locked = False
Next
For i = 1 To 4
Images(i).Visible = True
Next
If NewRecord Then
SearchBox = ""
    For i = 1 To 3
    Results(i) = ""
    Next
cmdEdit.Enabled = False
On Error Resume Next
SearchBox.SetFocus
On Error GoTo 0
Else
SearchBox.Locked = True
Results(1).SelStart = 0
Results(1).SelLength = Len(Results(1))
On Error Resume Next
Results(1).SetFocus
On Error GoTo 0
End If
cmdNew.Visible = False
cmdEdit.Visible = False
cmdSave.Visible = True
cmdCancel.Visible = True
mnuCat.Visible = False
mnuFind.Visible = False
mnuEdit.Visible = False
mnuNew.Visible = False
mnuDelete.Visible = False
mnuSave.Visible = True
mnuCancel.Visible = True
SearchBox.ToolTipText = "Type Identifier Here"
Selecting = False
End Sub

Private Sub FillResults()
Dim FullString As String
Dim i As Integer
Dim Start As Integer
Dim Spot As Integer
Dim StringLen As Integer

SearchBox = LeftList.Text
Selection = LeftList.ListIndex
cmdEdit.Enabled = True
mnuDelete.Enabled = True
ReDim Parts(0) As String
FullString = RightList.Text
Spot = InStr(1, FullString, ";")

If Spot > 0 Then
i = 1
Start = 1
ReDim Preserve Parts(1) As String
Parts(1) = Mid(FullString, Start, Spot - 1)
    Do
    DoEvents
    i = i + 1
    Start = Spot + 2
    Spot = InStr(Start, FullString, ";")
    StringLen = Spot - Start
    ReDim Preserve Parts(i) As String
    
    If Spot > 0 Then
    Parts(i) = Mid(FullString, Start, StringLen)
    Else
    Parts(i) = Mid(FullString, Start)
    End If
    
    Loop While Spot > 0 And UBound(Parts) < 3
    
    For i = 1 To 3
    Results(i) = ""
    Next
    
    For i = 1 To UBound(Parts)
    Results(i) = Parts(i)
    Next
Else
Results(1) = FullString
End If

End Sub


Private Function IsRated(Def As String) As Boolean

If InStr(1, UCase(Def), "FUCK") > 0 Or _
    InStr(1, UCase(Def), "SHIT") > 0 Or _
    InStr(1, UCase(Def), "PISS") > 0 Or _
    InStr(1, UCase(Def), " ASS ") > 0 Then
IsRated = True
End If

End Function

Private Sub NormalMode()
Dim Ctrl As Control
Dim i As Integer
Selecting = True
For Each Ctrl In Main
If TypeOf Ctrl Is ListBox Or TypeOf Ctrl Is Label Then
Ctrl.Enabled = True
End If
Next
For i = 0 To 3
CopyButton(i).Visible = True
Next
For i = 1 To 4
Images(i).Visible = False
Next
For i = 1 To 3
Results(i) = ""
Results(i).Locked = True
Next
SearchBox = ""
SearchBox.Locked = False
cmdNew.Visible = True
cmdEdit.Visible = True
cmdEdit.Enabled = False
cmdSave.Visible = False
cmdCancel.Visible = False
mnuCat.Visible = True
mnuFind.Visible = True
mnuEdit.Visible = True
mnuNew.Visible = True
mnuDelete.Visible = True
mnuDelete.Enabled = False
mnuSave.Visible = False
mnuCancel.Visible = False
SearchBox.ToolTipText = "Type first few letters of acronym..."
NewRecord = False
Selecting = False
End Sub

Private Sub RateDefinitions()
Dim TRS As Recordset
Dim i As Integer
Set TRS = TDB.OpenRecordset("SELECT * FROM Computer " _
    & "ORDER BY Acronym")
For i = 1 To TRS.RecordCount
    TRS.Edit
    If InStr(1, UCase(TRS.Fields("Definition")), "FUCK") > 0 Or _
        InStr(1, UCase(TRS.Fields("Definition")), "SHIT") > 0 Or _
        InStr(1, UCase(TRS.Fields("Definition")), "PISS") > 0 Or _
        InStr(1, UCase(TRS.Fields("Definition")), " ASS ") > 0 Then
    TRS.Fields("Rating") = 1
    Else
    TRS.Fields("Rating") = 0
    End If
    TRS.Update
If Not TRS.EOF Then TRS.MoveNext
Next
End Sub
Private Sub ReadAll(TableName As String)

Dim TRS As Recordset
Dim i As Integer
LeftList.Clear
RightList.Clear
If Rated Then
Set TRS = TDB.OpenRecordset("SELECT * FROM  " & TableName _
    & " WHERE Rating = '0' ORDER BY Acronym")
Else
Set TRS = TDB.OpenRecordset("SELECT * FROM  " & TableName _
    & " ORDER BY Acronym")
End If
For i = 1 To TRS.RecordCount
    If Not TRS.BOF Then
        If Len(TRS.Fields("Acronym")) > 0 Then
        LeftList.AddItem TRS.Fields("Acronym")
        RightList.AddItem TRS.Fields("Definition")
        End If
    End If
If Not TRS.EOF Then TRS.MoveNext
Next
TRS.Close
ReadingAll = True
FindSearch = False
End Sub











Private Sub SpaceLetters()
Dim i As Integer
For i = 1 To 25
Letter(i).Left = Letter(i - 1).Left + ScaleWidth / 26
Letter(i).Top = Letter(0).Top
Next
End Sub

Private Sub cmdCancel_Click()
NormalMode
End Sub

Private Sub cmdEdit_Click()
EditMode
End Sub
Private Sub cmdNew_Click()
NewRecord = True
EditMode
End Sub

Private Sub cmdSave_Click()
Dim i As Long
Dim n As Long
Dim TRS As Recordset
Dim FullDefine As String
Dim FieldName As String

If DefSearch Then
FieldName = "Definition"
Else
FieldName = "Acronym"
End If

If Len(Trim(Results(3))) > 0 Then
FullDefine = Trim(Replace(Results(1), ";", ":")) _
    & "; " & Trim(Replace(Results(2), ";", ":")) _
    & "; " & Trim(Replace(Results(3), ";", ":"))
ElseIf Len(Trim(Results(2))) > 0 Then
FullDefine = Trim(Replace(Results(1), ";", ":")) _
    & "; " & Trim(Replace(Results(2), ";", ":"))
Else
FullDefine = Trim(Replace(Results(1), ";", ":"))
End If

If Len(FullDefine) > 255 Then
FullDefine = Left(FullDefine, 255)
End If
If FindSearch And Not NewRecord Then GoTo Find:
If Len(Trim(SearchBox)) = 0 Then Exit Sub
If ReadingAll Then
    If Rated Then
    Set TRS = TDB.OpenRecordset("SELECT * FROM  " & TableName _
        & " WHERE Rating = '0' ORDER BY Acronym")
    Else
    Set TRS = TDB.OpenRecordset("SELECT * FROM  " & TableName _
        & " ORDER BY Acronym")
    End If
Else
    If Rated Then
    Set TRS = TDB.OpenRecordset("SELECT * FROM " & TableName _
        & " WHERE Acronym Like '" & LastSearch & "*' " _
        & " AND Rating = '0' ORDER BY Acronym")
    Else
    Set TRS = TDB.OpenRecordset("SELECT * FROM " & TableName _
        & " WHERE Acronym Like '" & LastSearch & "*' " _
        & " ORDER BY Acronym")
    End If
End If
If NewRecord Then
TRS.AddNew
Else
TRS.Move Selection
TRS.Edit
End If

TRS.Fields("Acronym") = Trim(SearchBox)
TRS.Fields("Definition") = FullDefine
    If IsRated(FullDefine) Then
    TRS.Fields("Rating") = 1
    Else
    TRS.Fields("Rating") = 0
    End If
TRS.Update
TRS.Close
NormalMode
ReadAll TableName
Exit Sub
Find:
If Rated Then
Set TRS = TDB.OpenRecordset("SELECT * FROM " & TableName _
    & " WHERE Rating = '0' ORDER BY Acronym")
Else
Set TRS = TDB.OpenRecordset("SELECT * FROM " & TableName _
    & " ORDER BY Acronym")
End If
For i = 1 To TRS.RecordCount
    If InStr(1, UCase(TRS.Fields(FieldName)), UCase(LastSearch)) > 0 Then
    n = n + 1
    If n = Selection + 1 Then Exit For
    End If
If Not TRS.EOF Then TRS.MoveNext
Next
TRS.Edit
TRS.Fields("Definition") = FullDefine
    If IsRated(FullDefine) Then
    TRS.Fields("Rating") = 1
    Else
    TRS.Fields("Rating") = 0
    End If
TRS.Update
TRS.Close
NormalMode
ReadAll TableName
End Sub

Private Sub CopyButton_Click(Index As Integer)
Clipboard.Clear
If Index > 0 Then
Clipboard.SetText Results(Index)
Else
Clipboard.SetText SearchBox
End If
End Sub

Private Sub FocusTimer_Timer()
FocusTimer = False
On Error Resume Next
SearchBox.SetFocus
On Error GoTo 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Unload Me
End If
End Sub

Private Sub Form_Load()
Dim DontShow As Boolean
Dim Msg As VbMsgBoxResult
DontShow = CBool(GetSetting("Translator", "UserPref", "DontShow", "False"))
Rated = CBool(GetSetting("Translator", "UserPref", "Rated", "True"))
TableName = GetSetting("Translator", "UserPref", "LastOpen", "Chat")
mnuRated.Checked = Not Rated
BarPosition = CInt(GetSetting("Translator", "Metrics", "Bar", 30))
Me.Width = CInt(GetSetting("Translator", "Metrics", "Width", 6705))
Me.Height = CInt(GetSetting("Translator", "Metrics", "Height", 7800))
Me.Top = CInt(GetSetting("Translator", "Metrics", "Top", 0))
Me.Left = CInt(GetSetting("Translator", "Metrics", "Left", 1590))

UpperOnly SearchBox
SpaceLetters
OpenData
LoadCategories
ReadAll TableName
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode < 2 Then
    If MsgBox("Quit Translator?", vbOKCancel + vbInformation) = vbCancel Then
    Cancel = True
    Exit Sub
    End If
End If
SaveSetting "Translator", "Metrics", "Width", CStr(Me.Width)
SaveSetting "Translator", "Metrics", "Height", CStr(Me.Height)
SaveSetting "Translator", "Metrics", "Bar", CStr(BarPosition)
SaveSetting "Translator", "Metrics", "Top", CStr(Me.Top)
SaveSetting "Translator", "Metrics", "Left", CStr(Me.Left)
CloseDatabase
End Sub

Private Sub Form_Resize()

If Me.WindowState <> vbMinimized Then
LeftPane.Width = Me.ScaleWidth * (BarPosition / 100) - 30
RightPane.Width = Me.ScaleWidth - LeftPane.Width - 60
SpaceLetters
End If
    
End Sub

Private Sub LeftList_Click()
RightList.TopIndex = LeftList.TopIndex
RightList.ListIndex = LeftList.ListIndex
End Sub

Private Sub LeftPane_Resize()
LeftList.Left = LeftPane.ScaleLeft
LeftList.Width = LeftPane.ScaleWidth
LeftList.Height = LeftPane.ScaleHeight
LeftList.Top = LeftPane.ScaleTop
SearchBox.Width = LeftPane.ScaleWidth - 300
CopyButton(0).Left = LeftPane.Left + LeftPane.ScaleWidth - 200
Images(1).Left = CopyButton(0).Left
End Sub


Private Sub Letter_Click(Index As Integer)
SearchBox = Letter(Index).Caption

End Sub

Private Sub Letter_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Letter(Index).ForeColor = vbRed
End Sub




Private Sub mnuAcro_Click()
DefSearch = False
ShowFind Main, SearchBox
End Sub



Private Function LoadCategories()
Dim i As Integer
Dim TBL As TableDef
Dim FLD As Field
Dim FieldName As String
Dim TheTable As String
Dim FoundAcro As Boolean
Dim FoundDef As Boolean
Dim FoundRat As Boolean
Dim CatIndex As Integer
Dim FoundLast As Boolean
CatIndex = 1
For i = 0 To TDB.TableDefs.Count - 1
TheTable = TDB.TableDefs(i).Name
    If Left(TheTable, 4) <> "MSys" Then
        For Each FLD In TDB.TableDefs(i).Fields
        FieldName = FLD.Name
        If FieldName = "Acronym" Then FoundAcro = True
        If FieldName = "Definition" Then FoundDef = True
        If FieldName = "Rating" Then FoundRat = True
        Next
        
        If FoundAcro And FoundDef And FoundRat Then
        Load mnuCategory(CatIndex)
        mnuCategory(CatIndex).Visible = True
        mnuCategory(CatIndex).Caption = TheTable
            If TheTable = TableName Then
            mnuCategory(CatIndex).Checked = True
            FoundLast = True
            End If
        CatIndex = CatIndex + 1
        End If
    End If
Next
Set FLD = Nothing
Set TBL = Nothing
If Not FoundLast Then
TableName = mnuCategory(1).Caption
mnuCategory(1).Checked = True
SaveSetting "Translator", "UserPref", "LastOpen", TableName
End If
End Function



Private Sub mnuAdd_Click()
Dim Cat As String
Cat = InputBox("Type name for new category.", "CATEGORY NAME", "")
If Len(Trim(Cat)) > 0 Then
    If Len(Trim(Cat)) < 13 Then
    AddTable Trim(Cat)
    Load mnuCategory(mnuCategory.UBound + 1)
    mnuCategory(mnuCategory.UBound).Visible = True
    mnuCategory(mnuCategory.UBound).Caption = Trim(Cat)
    mnuCategory_Click mnuCategory.UBound
    Else
        If MsgBox("Please make name less than 13 characters.", vbOKCancel + vbInformation) = vbOK Then
        mnuAdd_Click
        End If
    End If
End If
End Sub

Private Sub mnuCancel_Click()
cmdCancel_Click
End Sub

Private Sub mnuCategory_Click(Index As Integer)
Dim i As Integer
mnuCategory(Index).Checked = Not mnuCategory(Index).Checked
For i = 1 To mnuCategory.UBound
    If i <> Index Then
        If mnuCategory(i).Checked Then
        mnuCategory(i).Checked = False
        End If
    End If
Results(i) = ""
Next
SearchBox = ""
mnuCategory(Index).Checked = True
TableName = mnuCategory(Index).Caption
ReadAll TableName
SaveSetting "Translator", "UserPref", "LastOpen", mnuCategory(Index).Caption
End Sub


Private Sub mnuChange_Click()
cmdEdit_Click
End Sub

Private Sub mnuDef_Click()
DefSearch = True
ShowFind Main, SearchBox
End Sub

Private Sub mnuDelete_Click()
Dim TRS As Recordset
If MsgBox("Are you sure you want to delete the acronym " _
    & SearchBox & "? Cannot be undone!", vbOKCancel + vbInformation) = vbOK Then
If ReadingAll Then
    If Rated Then
    Set TRS = TDB.OpenRecordset("SELECT * FROM  " & TableName _
        & " WHERE Rating = '0' ORDER BY Acronym")
    Else
    Set TRS = TDB.OpenRecordset("SELECT * FROM  " & TableName _
        & " ORDER BY Acronym")
    End If
Else
    If Rated Then
    Set TRS = TDB.OpenRecordset("SELECT * FROM " & TableName _
        & " WHERE Acronym Like '" & LastSearch & "*' " _
        & " AND Rating = '0' ORDER BY Acronym")
    Else
    Set TRS = TDB.OpenRecordset("SELECT * FROM " & TableName _
        & " WHERE Acronym Like '" & LastSearch & "*' " _
        & " ORDER BY Acronym")
    End If
End If
TRS.Move Selection
TRS.Delete
TRS.Close
NormalMode
ReadAll TableName
End If
End Sub



Private Sub mnuEdit_Click()
mnuChange.Enabled = cmdEdit.Enabled
End Sub

Private Sub mnuExit_Click()
CloseDatabase
End
End Sub


Private Sub mnuFile_Click()
mnuDelete.Enabled = cmdEdit.Enabled
End Sub

Private Sub mnuNew_Click()
cmdNew_Click
End Sub


Private Sub mnuRated_Click()
mnuRated.Checked = Not mnuRated.Checked

If mnuRated.Checked Then
Rated = False
Else
Rated = True
End If

SaveSetting "Translator", "UserPref", "Rated", CStr(Rated)
ReadAll "Computer"

End Sub

Private Sub mnuSave_Click()
cmdSave_Click
End Sub





Private Sub RightList_Click()
LeftList.TopIndex = RightList.TopIndex
LeftList.ListIndex = RightList.ListIndex
Selecting = True
FillResults
Selecting = False
End Sub

Private Sub RightPane_Resize()
Dim i As Integer
RightList.Left = RightPane.ScaleLeft
RightList.Width = RightPane.ScaleWidth
RightList.Height = RightPane.ScaleHeight
RightList.Top = RightPane.ScaleTop
For i = 1 To 3
Results(i).Width = RightPane.ScaleWidth - 200
Results(i).Left = RightPane.Left
CopyButton(i).Left = RightPane.Left + RightPane.ScaleWidth - 200
Images(i + 1).Left = CopyButton(i).Left
Next

End Sub


Private Sub SearchBox_Change()

Dim i As Integer

For i = 1 To 3
Results(i) = ""
Next

If Not Selecting Then
LeftList.Clear
RightList.Clear
    If Len(SearchBox) = 0 Then
    ReadAll TableName
    Else
    FindNames TableName
    End If
End If
End Sub

Private Function FindNames(TableName As String) As Boolean
On Local Error Resume Next
Dim Trial As String
Dim Leng As Integer
Dim i As Long
Dim TRS As Recordset

Screen.MousePointer = vbArrowHourglass
Trial = Replace(SearchBox, "  ", " ")
Trial = Replace(SearchBox, ", ", ",")
Trial = Replace(SearchBox, "'", "' & Chr(39) & '")
Leng = Len(Trial)

If Leng = 0 Then
Screen.MousePointer = vbDefault
Exit Function
End If
If Rated Then
Set TRS = TDB.OpenRecordset("SELECT * FROM " & TableName _
    & " WHERE Acronym Like '" & Trial & "*' " _
    & " AND Rating = '0' ORDER BY Acronym")
Else
Set TRS = TDB.OpenRecordset("SELECT * FROM " & TableName _
    & " WHERE Acronym Like '" & Trial & "*' " _
    & " ORDER BY Acronym")
End If
For i = 1 To TRS.RecordCount
    LeftList.AddItem TRS.Fields("Acronym")
    RightList.AddItem TRS.Fields("Definition")
If Not TRS.EOF Then TRS.MoveNext
Next

TRS.Close
ReadingAll = False
LastSearch = Trial
FindSearch = False
Screen.MousePointer = vbDefault
End Function


Private Sub SearchBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SearchBox.SelStart = 0
SearchBox.SelLength = Len(SearchBox)
End Sub


Private Sub SplitterBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim MousePosition As Integer

If Button = vbLeftButton Then
MousePosition = (SplitterBar.Left + X) / Me.ScaleWidth * 100
    If MousePosition > 16 And MousePosition < 90 Then
    LeftPane.Width = SplitterBar.Left + X
    RightPane.Width = Me.ScaleWidth - LeftPane.Width - 40
    End If
    
    If MousePosition < 16 Then
    BarPosition = 17
    ElseIf MousePosition > 89 Then
    BarPosition = 89
    Else
    BarPosition = MousePosition
    End If
End If
End Sub




Private Sub TopPane_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 25
    If Letter(i).ForeColor = vbRed Then
    Letter(i).ForeColor = &HFF0000
    Exit Sub
    End If
Next
End Sub


