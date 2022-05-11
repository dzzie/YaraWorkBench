VERSION 5.00
Object = "{2668C1EA-1D34-42E2-B89F-6B92F3FF627B}#5.0#0"; "scivb2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Yara Workbench"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   15120
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   ScaleHeight     =   9915
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame hSplit 
      BackColor       =   &H00808080&
      Height          =   90
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   15
      Top             =   7440
      Width           =   14715
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   9120
      TabIndex        =   8
      Top             =   60
      Width           =   5775
      Begin VB.CommandButton cmdScan 
         Caption         =   "Scan"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2640
         TabIndex        =   14
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmdAbort 
         Caption         =   "Abort"
         Height          =   315
         Left            =   4200
         TabIndex        =   13
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   315
         Left            =   900
         TabIndex        =   12
         Top             =   0
         Width           =   435
      End
      Begin VB.CommandButton cmdBrowseDir 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   11
         Top             =   0
         Width           =   495
      End
      Begin VB.CheckBox chkRecursive 
         Caption         =   "R"
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         ToolTipText     =   "Recursive Scan (Folders)"
         Top             =   0
         Width           =   435
      End
      Begin VB.TextBox txtFilter 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   9
         Text            =   "*"
         ToolTipText     =   "File Filter"
         Top             =   0
         Width           =   795
      End
   End
   Begin YaraWorkBench.ucFilterList lv 
      Height          =   6795
      Left            =   0
      TabIndex        =   7
      Top             =   540
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   11986
   End
   Begin VB.Frame splitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   5865
      Left            =   4680
      MousePointer    =   9  'Size W E
      TabIndex        =   6
      Top             =   540
      Width           =   90
   End
   Begin VB.Timer tmrInit 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   12480
      Top             =   360
   End
   Begin YaraWorkBench.HistoryCombo txtSample 
      Height          =   315
      Left            =   1020
      TabIndex        =   5
      Top             =   60
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   556
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   195
      Left            =   4800
      TabIndex        =   4
      Top             =   420
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtDemo 
      Height          =   1575
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "frmMain.frx":08CA
      Top             =   1560
      Visible         =   0   'False
      Width           =   2475
   End
   Begin sci2.SciSimple sci 
      Height          =   6795
      Left            =   4740
      TabIndex        =   2
      Top             =   600
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   11986
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   0
      TabIndex        =   0
      Top             =   7560
      Width           =   14655
   End
   Begin VB.Label Label1 
      Caption         =   "Sample"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   795
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewRule 
         Caption         =   "New Rule File"
      End
      Begin VB.Menu mnuLoadRuleFile 
         Caption         =   "Load Rule File"
      End
      Begin VB.Menu mnuSaveRules 
         Caption         =   "Save Rule"
      End
      Begin VB.Menu mnuSaveRulesAs 
         Caption         =   "Save Rules As"
      End
      Begin VB.Menu mnuSpacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecent 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu mnuSpacer22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewInstance 
         Caption         =   "New Instance"
      End
      Begin VB.Menu mnuBuildSampleSet 
         Caption         =   "Build Sample Set"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuWordWrap 
         Caption         =   "Word Wrap"
      End
      Begin VB.Menu mnuCommentBlock 
         Caption         =   "Comment Block"
      End
      Begin VB.Menu mnuUncommentBlock 
         Caption         =   "Uncomment Block"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuLibrary 
         Caption         =   "Library"
      End
      Begin VB.Menu mnuRuleNavigator 
         Caption         =   "Rule Navigator"
      End
      Begin VB.Menu mnuModDump 
         Caption         =   "Dump Module Info"
      End
      Begin VB.Menu mnuHexEditFile 
         Caption         =   "Hexedit File"
      End
      Begin VB.Menu mnuSpacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTestOld 
         Caption         =   "Old Versions"
         Begin VB.Menu mnuTestOldV 
            Caption         =   "About"
            Index           =   0
         End
      End
      Begin VB.Menu mnuShowDbg 
         Caption         =   "Debug Msgs"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuRegister 
         Caption         =   "Register"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuYaraHelp 
         Caption         =   "Yara Help"
      End
      Begin VB.Menu mnuYaramods 
         Caption         =   "Yara Modules"
         Begin VB.Menu mnuYaraMod 
            Caption         =   "mnuMods0"
            Index           =   0
         End
      End
      Begin VB.Menu mnuAboutSci 
         Caption         =   "About Scintinilla"
      End
      Begin VB.Menu mnuAboutYWB 
         Caption         =   "About Yara Workbench"
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowBenchMarks 
         Caption         =   "Show File Bench Marks"
      End
      Begin VB.Menu mnuSetDisasmPath 
         Caption         =   "Set Disassembler Path"
      End
      Begin VB.Menu mnuExternalHexeditor 
         Caption         =   "Set External Hexeditor"
      End
      Begin VB.Menu mnuRegFileExt 
         Caption         =   "Register .yara File Ext"
      End
      Begin VB.Menu mnuSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDemoSample 
         Caption         =   "Demo Sample"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuCopyAll 
         Caption         =   "Copy All"
      End
      Begin VB.Menu mnuCopyLine 
         Caption         =   "Copy Line"
      End
      Begin VB.Menu mnuCopyPath 
         Caption         =   "Copy Path"
      End
      Begin VB.Menu mnuGenerateReport 
         Caption         =   "Generate Report"
      End
      Begin VB.Menu mnuHexEdit 
         Caption         =   "Hex Edit"
      End
      Begin VB.Menu mnuDisassemble 
         Caption         =   "Disassemble"
      End
      Begin VB.Menu mnuSqlWhereInParent 
         Caption         =   "Sql Where In"
         Begin VB.Menu mnuSqlWhereIn 
            Caption         =   "Hits"
            Index           =   0
         End
         Begin VB.Menu mnuSqlWhereIn 
            Caption         =   "Misses"
            Index           =   1
         End
      End
      Begin VB.Menu mnuMoveFiles 
         Caption         =   "Move"
         Begin VB.Menu mnuDoMove 
            Caption         =   "Matches"
            Index           =   0
         End
         Begin VB.Menu mnuDoMove 
            Caption         =   "Misses"
            Index           =   1
         End
         Begin VB.Menu mnuDoMove 
            Caption         =   "Selected"
            Index           =   2
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public curFile As String
Public isInitilized As Boolean
Public WithEvents curScan As CYaraScan
Attribute curScan.VB_VarHelpID = -1
Public curResults As Collection
Public isense As New CollectionEx
Public testFiles As New CollectionEx 'part of suble demo reset license scheme Test.yar indicates demo key generated (over test.yar)
Private crashLatter As Boolean 'they failed demo reset 2
Private disassembler As String
Private hexeditor As String

Dim deltaSplitterTop As Long
Private Capturing As Boolean
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Const MAX_RECENTS = 9

Private Sub highLightErrorLine(x)
    On Error Resume Next
     Dim lineNo As Long, a As Long, tmp As String
     x = Replace(x, "ERROR:", Empty)
     a = InStr(x, ":")
     If a > 0 Then
        tmp = Replace(Mid(x, 1, a - 1), "line ", Empty)
        lineNo = CLng(tmp)
        If lineNo > 0 Then sci.GotoLineCentered lineNo
     End If
End Sub

Private Sub cmdScan_Click()

    Dim c As Collection
    Dim misses As New Collection
    Dim li As ListItem
    Dim f As CYaraFile
    Dim tmp As String
    Dim t As Long
    Dim matches As Long
    Dim x, pcent As Double
    Dim isFolderScan As Boolean
    
    List1.Clear
    lv.ListItems.Clear
    txtSample.RecordIfNew
    List1.AddItem "Starting scan " & Now & " of " & txtSample
    Set curResults = Nothing
    
    txtFilter = trim(txtFilter)
    If Len(txtFilter) = 0 Then txtFilter = "*.*"
    If Left(txtFilter, 1) <> "*" Then txtFilter = "*" & txtFilter
    
    Set curScan = New CYaraScan
    If chkRecursive.value = vbChecked Then curScan.recursive = True

    If fso.FolderExists(txtSample) Then isFolderScan = True
    
    StartBenchMark t
    curScan.fileFilter = txtFilter
    Set c = curScan.scan(txtSample, sci.Text)
    Set curResults = c
    pb.value = 0 'in case of user abort
    
    If curScan.errors.count > 0 Then
        For Each x In curScan.errors
            List1.AddItem "ERROR: " & x
            highLightErrorLine x
        Next
        Me.Caption = curScan.errors.count & " Errors"
    Else
        Me.Caption = "Scan Complete: " & EndBenchMark(t)
    End If
        
    If c.count > 0 Then List1.AddItem String(30, "-") & "[ Matches " & curScan.matchCount & " ]" & String(30, "-")
    
    For Each f In c
        If f.found Then
            If f.isMemStr Then
                tmp = "Str: 0x" & Hex(Len(f.File))
            Else
                If fso.FolderExists(txtSample) And chkRecursive.value = vbChecked Then
                    tmp = Replace(f.File, txtSample, ".")
                Else
                    tmp = fso.FileNameFromPath(f.File)
                End If
            End If
            Set li = lv.ListItems.Add(, , lpad(f.TotalMatches, 3))  'f.matches.count is rules that hit..
            Set li.Tag = f
            li.ToolTipText = f.File
            li.subItems(1) = f.MatchNames()
            li.subItems(2) = tmp
            
            If Not isFolderScan Then
                List1.AddItem " " & f.MatchNames & " - " & Replace(f.File, fso.GetParentFolder(txtSample), ".")
            Else
                List1.AddItem " " & f.MatchNames & " - " & Replace(f.File, txtSample, ".")
            End If
            
            matches = matches + 1
        Else
            misses.Add f
        End If
    Next
        
    If misses.count > 0 Then
        'List1.AddItem String(70, "-"): List1.AddItem "Misses: " & misses.count
        List1.AddItem String(30, "-") & "[ Misses " & misses.count & " ]" & String(30, "-")
        For Each f In misses
            If Not isFolderScan Then
                List1.AddItem " " & Replace(f.File, fso.GetParentFolder(txtSample), ".")
            Else
                List1.AddItem " " & Replace(f.File, txtSample, ".")
            End If
        Next
    End If
    
    On Error Resume Next
    pcent = Round((matches / c.count) * 100, 2)
    Me.Caption = Me.Caption & " " & c.count & " files scanned " & matches & " matches ( " & pcent & "% )"
    
    
End Sub

Private Sub curScan_info(msgType As cb_type, msg As String)
    On Error Resume Next
    If msgType = cb_update Then
        List1.List(List1.ListCount - 1) = List1.List(List1.ListCount - 1) & " - " & msg
    Else
        List1.AddItem msg
    End If
End Sub

Sub LoadTestFiles()
    On Error Resume Next
    Dim x() As String, f, demoFound As Boolean, fName As String
    x = fso.GetFolderFiles(App.path, "*.yar*")
    For Each f In x
       fName = fso.FileNameFromPath(CStr(f))
       testFiles.Add f, fName
       If LCase(fName) = test & YaraExt Then demoFound = True
    Next
    If Not demoFound Then
        fso.writeFile App.path & "\" & test & YaraExt, txtDemo.Text
    End If
End Sub

Sub LoadRecents()
    
    On Error Resume Next
    
    Dim recents() As String
    Dim i As Long
    
    For i = 1 To MAX_RECENTS
        Load mnuRecent(i)
        mnuRecent(i).Visible = False
    Next
    
    recents = Split(GetSetting("Yara WorkBench", "Options", "Recents", ",,,"), ",")
    
    For i = 0 To MAX_RECENTS
        If i > UBound(recents) Then Exit For
        If fso.FileExists(recents(i)) Then
            mnuRecent(i).Visible = True
            mnuRecent(i).Tag = recents(i)
            mnuRecent(i).Caption = AbbreviatedPathForDisplay(recents(i))
        Else
            mnuRecent(i).Visible = False
        End If
    Next

End Sub

Sub AddToRecentList(filename As String, FileTitle As String)

    'On Error GoTo errHandle
    
    Dim x, i
    Dim c As New Collection
    c.Add filename 'new one is always first...
    
    For i = 0 To MAX_RECENTS
        If Len(mnuRecent(i).Tag) > 0 Then
            If mnuRecent(i).Tag <> filename Then           'no duplicate entries
                If fso.FileExists(mnuRecent(i).Tag) Then   'only keep files which still exist
                    c.Add mnuRecent(i).Tag
                End If
            End If
        End If
        mnuRecent(i).Tag = Empty                   'out with the old
        mnuRecent(i).Caption = Empty
        mnuRecent(i).Visible = False
    Next
    
    For i = 0 To MAX_RECENTS
        If i > c.count - 1 Then Exit For
        mnuRecent(i).Tag = c(i + 1)                'in with the new..
        mnuRecent(i).Caption = AbbreviatedPathForDisplay(c(i + 1))
        mnuRecent(i).Visible = True
    Next
    
    For i = 0 To MAX_RECENTS
        x = x & mnuRecent(i).Tag & ","
    Next

    SaveSetting "Yara WorkBench", "Options", "Recents", x

Exit Sub
errHandle:
    MsgBox "Error_frmMain_AddToRecentList: " & Err.Description

End Sub

Function AbbreviatedPathForDisplay(ByVal FullPath) As String
    Dim tmp() As String, abbrivate As Boolean, fName As String, ext As String, shortPath
    Const maxLen = 50
    
    If Len(FullPath) = 0 Then Exit Function
    
    If InStr(FullPath, "\") > 0 Then
        If Len(FullPath) < maxLen Then
            AbbreviatedPathForDisplay = FullPath
        Else
            tmp = Split(FullPath, "\")
            fName = tmp(UBound(tmp))
            FullPath = Replace(FullPath, fName, Empty, , 1)
            ext = fso.GetExtension(fName)
            fName = Replace(fName, ext, Empty)
            
            If Len(ext) > 6 Then ext = ".." & Right(ext, 4)
            If Len(fName) > 25 Then fName = Mid(fName, 1, 20) & "..."
            fName = fName & ext
            
            If Len(FullPath) > maxLen Then
                  shortPath = Mid(FullPath, 1, 18) & "..." & Right(FullPath, 18)
                  AbbreviatedPathForDisplay = shortPath & fName
            Else
                  AbbreviatedPathForDisplay = FullPath & fName
            End If
            
        End If
    End If
    
End Function



Private Sub mnuDoMove_Click(Index As Integer)
    Dim f As CYaraFile
    Dim i As Long
    Dim pth As String
    Dim doit As Boolean
    Dim li As ListItem
    Dim errs() As String
    
    On Error Resume Next
    
    If curResults Is Nothing Then
        MsgBox "No current scan?", vbInformation
        Exit Sub
    End If
    
    pth = dlg.FolderDialog2(txtSample)
    If Len(pth) = 0 Then Exit Sub
    
    If Index = 2 Then 'selected items
        Err.Clear
        For Each li In lv.selItems
            Set f = li.Tag
            fso.Move f.File, pth
            If Err.Number <> 0 Then push errs, f.File Else i = i + 1
        Next
    Else
        Err.Clear
        For Each f In curResults
            doit = False
            If Index = 0 And f.matches.count > 0 Then doit = True 'matches
            If Index = 1 And f.matches.count = 0 Then doit = True 'misses
            If doit Then
                fso.Move f.File, pth
                If Err.Number <> 0 Then push errs, f.File Else i = i + 1
            End If
        Next
    End If
    
'    If Index = 0 Then
'        lv.ListItems.Clear
'        txtSample = pth
'    End If

    Dim msg As String, e
    
    msg = i & " files moved"
    
    If Not AryIsEmpty(errs) Then
        msg = msg & " " & UBound(errs) & " Errors (see list)"
        For Each e In errs
            List1.AddItem "Move error: " & e
        Next
    End If
    
    MsgBox msg, vbInformation
    
End Sub

Private Sub mnuGenerateReport_Click()
    
    Dim li As ListItem
    Dim c As New CollectionEx
    Dim f As CYaraFile
    Dim m As CYaraMatch
    Dim tmp() As String
    Dim i As Long
    Dim k As String
    Dim v As String
    
    On Error Resume Next
    
    If lv.ListItems.count = 0 Then
        MsgBox "No results generated yet", vbInformation
        Exit Sub
    End If
    
    For Each li In lv.ListItems
        Set f = li.Tag
        For Each m In f.matches
            If c.keyExists(m.name) Then
                c(m.name, 1) = c(m.name, 1) & vbCrLf & f.File
            Else
                c.Add f.File, m.name
            End If
        Next
    Next
    
    push tmp, "This is a temporary file, Save As if you wish to keep it" & vbCrLf
    push tmp, Me.Caption & vbCrLf
    
    For i = 1 To c.count
        k = c.keyForIndex(i)
        v = c(k, 1)
        push tmp, "Rule: " & k & " (" & CountOccurances(v, vbCrLf) & ")" & vbCrLf & String(75, "-")
        push tmp, v & vbCrLf
    Next
    
    k = fso.GetFreeFileName(Environ("temp"))
    fso.writeFile k, Join(tmp, vbCrLf)
    Shell "notepad.exe """ & k & """", vbNormalFocus
        
End Sub

Private Sub mnuNewInstance_Click()
    On Error Resume Next
    Shell App.path & "\ywb.exe", vbNormalFocus
End Sub

Private Sub mnuRecent_Click(Index As Integer)
    Dim path As String
    Dim v As VbMsgBoxResult
    
    path = mnuRecent(Index).Tag
    
    If fso.FileExists(path) Then

        If sci.isDirty Then
            v = MsgBox("Changes not saved save now?", vbYesNo)
            If v = vbYes Then mnuSaveRules_Click
        End If
        
        curFile = Empty
        If sci.LoadFile(path) Then
            curFile = path
            load_ywbPath
        End If
        
    End If
    
End Sub


Private Sub Form_Load()
    Dim tmp
    Dim rv As Long
    Dim cmd As String
    Dim ext As String
    
    LoadRecents
    FormPos Me, True
    splitter.Top = sci.Top
    lv.Top = sci.Top
    splitter.Width = 66
    lv.SetFont "Courier", 10
    lv.SetColumnHeaders "n,Matches*,File", "500,2700,*"
    lv.MultiSelect = True
    
    sci.LoadHighlighter App.path & "\java.hilighter", True
    mnuPopup.Visible = False
    disassembler = GetSetting("yara", "settings", "disassembler", "")
    txtSample.LoadHistory App.path & "\hc.dat"
    txtSample.AutoCompletePaths = True
    mnuShowBenchMarks.Checked = GetSetting("yara", "settings", "mnuShowBenchMarks", 1)
    opt.ShowBenchMarks = IIf(mnuShowBenchMarks.Checked = vbChecked, True, False)
    
'    #If isBuilder = 1 Then
'        'this way no mistaking which build it is...
'        frmRegister.Show 1
'        End
'    #End If
'
'    If InStr(1, Command, "/register", vbTextCompare) > 0 Then
'        frmRegister.Show 1
'        End
'    End If
    
    cmd = Replace(Command, """", Empty)
    If fso.FileExists(cmd) Then
        ext = LCase(fso.GetExtension(cmd))
        If ext = ".txt" Or Left(ext, 4) = ".yar" Then
            If sci.LoadFile(cmd) Then curFile = cmd
        Else
            'txtSample_GotFocus
            txtSample.Text = cmd
        End If
    ElseIf fso.FolderExists(cmd) Then
        'txtSample_GotFocus
        txtSample.Text = cmd
    End If
    
    cmdScan.Enabled = InitLibYara()
    txtSample.AutoCompletePaths = True
    'SHAutoComplete txtSample.hwnd, SHACF_FILESYSTEM
    
    LoadTestFiles 'part of license scheme
    cmdScan.Enabled = True
    
    SyntaxColor sci
    'sci.misc.GutterWidth(gut0) = 60
    sci.WordWrap = False
    sci.DirectSCI.SendEditor SCI_AUTOCSETIGNORECASE, 1, 1
    
    If False And isIde Then
        'txtSample_GotFocus
        txtSample = "C:\Users\home\Desktop\dd\k32b"
        sci.LoadFile App.path & "\test.yar"
    Else
        tmp = GetSetting("yara", "settings", "txtSample")
        If fso.FileExists(CStr(tmp)) Then
            txtSample = tmp
            'txtSample.ForeColor = vbBlack
        End If
    End If
    
    Dim c() As String, f, ii As CIntellisenseItem, x, pth As String
    
    pth = App.path & "\intellisense"
    If Not fso.FolderExists(pth) Then
        pth = App.path & "\yara_workbench\intellisense"
        If Not fso.FolderExists(pth) Then
            List1.AddItem "intellisense folder Not found"
            Exit Sub
        End If
    End If

    c = fso.GetFolderFiles(pth, "*.txt")
    If AryIsEmpty(c) Then Exit Sub
    
    For Each f In c
        Set ii = New CIntellisenseItem
        ii.loadSelf f
        If ii.keyWordCount > 0 Then
            isense.Add ii, ii.ObjName
            LoadNewModuleDef f
        End If
    Next
    
'    If Flags(f_demoKeyGenerated) > 20 Then  'demo key created
'        If testFiles.keyExists("Test.yar") Then
'            CallWindowProc AddressOf WndProc, Me.hwnd, -1, &H4127, 89790 'little present to crash latter
'        Else
'            CallWindowProc AddressOf WndProc, Me.hwnd, -99, &H11229987, 4233  'rename the file for next time..
'        End If
'        tmrInit.Tag = Flags(f_demoKeyGenerated)
'    End If
        
    Dim oldVers() As String
    oldVers() = fso.GetFolderFiles(App.path & "\alt_vers", "*.exe")
    If Not AryIsEmpty(oldVers) Then
        For Each x In oldVers
            Load mnuTestOldV(mnuTestOldV.count + 1)
             mnuTestOldV(mnuTestOldV.count).Caption = Replace(fso.GetBaseName(CStr(x)), "yara", Empty, , , vbTextCompare)
             If Len(mnuTestOldV(mnuTestOldV.count).Caption) = 0 Then mnuTestOldV(mnuTestOldV.count).Caption = "yara"
             mnuTestOldV(mnuTestOldV.count).Tag = x
        Next
    End If
         
    List1.AddItem isense.count & " intellisense items added"
    tmrInit.Enabled = True
    
End Sub
 
Function CheckSubStatement() As Boolean

    On Error GoTo hell
    Dim last As String, a As Long, obj As String, s As CSubItem, i As CIntellisenseItem
    
    'now look for subitems
    last = trim(GetPreceedingStatement())
    If Len(last) > 0 Then
        a = InStr(last, ".")
        If a > 0 Then
            obj = Mid(last, 1, a - 1)
            For Each i In isense
                If i.ObjName = obj Then
                    For Each s In i.subItems
                        If last Like s.ownerNoBrackets Then
                             sci.ShowAutoComplete s.keywords
                             CheckSubStatement = True
                             Exit Function
                        End If
                    Next
                End If
            Next
        End If
    End If
    
hell:
    If isIde And Err.Number <> 0 Then List1.AddItem "Error in checksubstatement:" & Err.Description
End Function

Function GetPreceedingStatement() As String
    Dim ln, x, i, str, breakChars
    
    On Error GoTo hell
    
    breakChars = " ();{}|\/><=&!^~+-'" & vbTab
    
    ln = sci.GetLineText(sci.CurrentLine()) '"&  pe.type ="

    For i = sci.GetCaretInLine() To 1 Step -1
        x = Mid(ln, i, 1)
        If x = "=" Or ((x = " " Or x = vbTab) And Len(str) = 0) Then
            'do nothing
        Else
            If InStr(breakChars, x) > 0 Then Exit For 'we found preceeding space before statement
            str = str & x
        End If
    Next
    
    GetPreceedingStatement = StrReverse(str)
hell:
End Function

Private Sub List1_Click()
    On Error Resume Next
    Dim x As String
    x = List1.List(List1.ListIndex)
    If Len(x) = 0 Then Exit Sub
    If InStr(x, "ERROR: line ") > 0 Then
        highLightErrorLine x
    End If
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    List1.SetFocus
    cfgMenus True
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Function cfgMenus(isList1 As Boolean)
    mnuCopyPath.Visible = Not isList1
    mnuHexEdit.Visible = Not isList1
    mnuDisassemble.Visible = Not isList1
    mnuCopyLine.Visible = isList1
End Function

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
     
     lv.SetFocus
     cfgMenus False
     
     If Button = 2 Then
        PopupMenu mnuPopup
        Exit Sub
     End If
     
'    On Error Resume Next
'    'can not do this below in mousedown, .selecteditem not yet set reliably...
'    Dim Item As ListItem
'    Dim f As CYaraFile, fm As frmMatches
'
'    If lv.selItem Is Nothing Then Exit Sub
'
'    If Button = 1 Then
'        Set Item = lv.selItem
'        Set f = Item.Tag
'        If GetKeyState(vbKeyShift) Then
'            Set fm = New frmMatches
'        Else
'            Set fm = frmMatches
'        End If
'        fm.LoadRules f
'    End If
    
End Sub

Private Sub lv_DblClick()
    
    On Error Resume Next
    'can not do this below in mousedown, .selecteditem not yet set reliably...
    Dim Item As ListItem
    Dim f As CYaraFile, fm As frmMatches
    
    If lv.selItem Is Nothing Then Exit Sub
    
    Set Item = lv.selItem
    Set f = Item.Tag
    If GetKeyState(vbKeyShift) Then
        Set fm = New frmMatches
    Else
        Set fm = frmMatches
    End If
    fm.LoadRules f
    
End Sub

Private Sub mnuAboutSci_Click()
    sci.ShowAbout
End Sub

Private Sub mnuAboutYWB_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuBuildSampleSet_Click()
    On Error Resume Next
    If fso.GetExtension(txtSample) = ".set" Then
        frmSampleList.LoadFile txtSample
    End If
    frmSampleList.Visible = True
End Sub

Private Sub mnuCopyAll_Click()
    Dim x As String
    
    If Me.ActiveControl Is lv Then
        x = lvGetAllElements(lv)
    Else
        x = lbCopy(List1)
    End If
    
    Clipboard.Clear
    Clipboard.SetText x
    
End Sub

Private Sub mnuCopyLine_Click()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText trim(List1.List(List1.ListIndex))
End Sub

Private Sub mnuCopyPath_Click()
    'only for lv
    On Error Resume Next
    Dim f As CYaraFile
    If lv.selItem Is Nothing Then Exit Sub
    Set f = lv.selItem.Tag
    'If f.isMemStr Then
    Clipboard.Clear
    Clipboard.SetText f.File
End Sub

Private Sub mnuDemoSample_Click()
    'txtSample_GotFocus
    txtSample = App.path
    If fso.FileExists(App.path & "\test.yar") Then
        sci.LoadFile App.path & "\test.yar"
        cmdScan_Click
    Else
        sci.Text = "Could not find test.yar?"
    End If
End Sub

Private Sub mnuDisassemble_Click()
    'only for lv
    On Error Resume Next
    Dim f As CYaraFile
    If lv.selItem Is Nothing Then Exit Sub
    Set f = lv.selItem.Tag
    
    If f.isMemStr Then Exit Sub
    
    If Not fso.FileExists(disassembler) Then mnuSetDisasmPath_Click
    If Not fso.FileExists(disassembler) Then Exit Sub
    
    If fso.FileExists(f.File) Then
        Shell """" & disassembler & """ """ & f.File & """"
        If Err.Number <> 0 Then
            MsgBox Err.Description
        End If
    Else
        MsgBox "File not found: " & f.File
    End If
    
End Sub



Private Sub mnuHexEdit_Click()
    'only for lv
    On Error Resume Next
    Dim f As CYaraFile
    Dim h As CHexEditor
    If lv.SelectedItem Is Nothing Then Exit Sub
    Set f = lv.SelectedItem.Tag
    If f.isMemStr Then Exit Sub
    
    If fso.FileExists(f.File) Then
        
        If Not fso.FileExists(hexeditor) Then
            Set h = New CHexEditor
            h.Editor.LoadFile f.File
        Else
            Shell """" & hexeditor & """ """ & f.File & """"
            If Err.Number <> 0 Then
                MsgBox Err.Description
            End If
        End If
    Else
        MsgBox "File not found: " & f.File
    End If
    
End Sub

Private Sub mnuHexEditFile_Click()

    On Error Resume Next
    
    Dim h As CHexEditor
    Dim f As String
    
    f = dlg.OpenDialog()
    If Len(f) = 0 Then Exit Sub
        
    If Not fso.FileExists(hexeditor) Then
        Set h = New CHexEditor
        h.Editor.LoadFile f
    Else
        Shell """" & hexeditor & """ """ & f & """"
        If Err.Number <> 0 Then
            MsgBox Err.Description
        End If
    End If
    
End Sub

Private Sub mnuLibrary_Click()
    frmLibrarian.Show
End Sub

Private Sub mnuModDump_Click()
    Dim f As New frmModuleDump
    f.Visible = True
End Sub

Private Sub mnuNewRule_Click()
    Dim tmp() As String
    
    If sci.isDirty Then mnuSaveRules_Click
    curFile = Empty
    
    push tmp, "import ""pe"""
    push tmp, "//ywbPath: ./ -r"
    push tmp, ""
    push tmp, "rule test"
    push tmp, "{"
    push tmp, "    strings:"
    push tmp, ""
    push tmp, "    condition:"
    push tmp, ""
    push tmp, "}"
    
    sci.Text = Join(tmp, vbCrLf)
    
End Sub

Private Sub mnuRegFileExt_Click()
    If register_FileExt() Then
        MsgBox ".YARA file extension is registered", vbInformation
    Else
        MsgBox "You must run the app as admin", vbInformation
    End If
End Sub

'Private Sub mnuRegister_Click()
'    frmRegister.Show 1
'End Sub

Private Sub mnuRuleNavigator_Click()
    frmNavigator.Show
End Sub

Private Sub mnuSetDisasmPath_Click()
    Dim tmp As String
    tmp = dlg.OpenDialog()
    If Len(tmp) = 0 Then
        If Len(disassembler) > 0 Then disassembler = Empty
    Else
        disassembler = tmp
    End If
End Sub

Private Sub mnuExternalHexeditor_Click()
    Dim tmp As String
    tmp = dlg.OpenDialog("", "Cancel to use internal hexeditor")
    If Len(tmp) = 0 Then
        If Len(hexeditor) > 0 Then hexeditor = Empty
    Else
        hexeditor = tmp
    End If
End Sub

Private Sub mnuShowBenchMarks_Click()
     mnuShowBenchMarks.Checked = Not mnuShowBenchMarks.Checked
     opt.ShowBenchMarks = IIf(mnuShowBenchMarks.Checked = vbChecked, True, False)
End Sub

Private Sub mnuShowDbg_Click()
    mnuShowDbg.Checked = Not mnuShowDbg.Checked
    DEBUG_CALLBACK = mnuShowDbg.Checked
End Sub

Private Sub mnuSqlWhereIn_Click(Index As Integer)
    Dim f As CYaraFile
    Dim i As Long
    Dim pth As String
    Dim c As New Collection
    Dim ret As String
    Dim doit As Boolean
    
    On Error Resume Next
    
    If curResults Is Nothing Then
        MsgBox "No current scan?", vbInformation
        Exit Sub
    End If
    
    
    pth = fso.GetFreeFileName(Environ("temp"))
    
    For Each f In curResults
        doit = False
        If Index = 0 And f.matches.count > 0 Then doit = True 'matches
        If Index = 1 And f.matches.count = 0 Then doit = True 'misses
        If doit Then
            c.Add fso.FileNameFromPath(f.File)
            i = i + 1
        End If
    Next
    
    ret = "\nThis is a temporary file, Save As if you want to keep it.\nCopies file name (which can be its hash) into a SQL In List\nThere were "
    ret = ret & i & IIf(Index = 0, " matches", " misses") & "\n\n"
    ret = Replace(ret, "\n", vbCrLf)
    
    fso.writeFile pth, ret & WhereInGuidList(c)
    Shell "notepad.exe """ & pth & """", vbNormalFocus
     
End Sub

Function WhereInGuidList(guids As Collection) As String
    Dim sql As String, s
    Dim tmp() As String
    
    sql = " in ("
    For Each s In guids
        push tmp, """" & s & """"
    Next
    sql = sql & Join(tmp, ",") & ")"
    
    WhereInGuidList = sql

End Function

Private Sub mnuTestOldV_Click(Index As Integer)
    
    Dim exe As String
    Dim cc As Object 'vbDevKit.CCmdOutput
    Dim yar As String, test As String
    Dim header As String
    
    On Error Resume Next
    
    Const msg = "Copy old yara exes into %app%\alt_vers\ \nand rename them w/ ver number"
    
    If Index = 0 Then
        MsgBox Replace(msg, "\n", vbCrLf), vbInformation
        Exit Sub
    End If
    
    Set cc = CreateObject("vbDevKit.CCmdOutput")
    If cc Is Nothing Then
        MsgBox "Failed to create CCmdOutput class?", vbExclamation
        Exit Sub
    End If
    
    exe = mnuTestOldV(Index).Tag
    header = String(75, "-") & vbCrLf & "Testing " & exe & vbCrLf & String(75, "-") & vbCrLf
    If Not fso.FileExists(exe) Then
        MsgBox "Path not found? " & exe
        Exit Sub
    End If
    
    If Len(sci.Text) = 0 Then
        MsgBox "Please enter a rule to test first", vbInformation
        Exit Sub
    End If
    
    If Not fso.FileExists(txtSample) And Not fso.FolderExists(txtSample) Then
        MsgBox "File or folder to test must exist", vbInformation
        Exit Sub
    End If
    
    yar = fso.GetFreeFileName(Environ("temp"))
    fso.writeFile yar, sci.Text
    
    yar = GetShortName(yar)
    exe = GetShortName(exe)
    test = GetShortName(txtSample)
    
    cc.maxExecSecs = 8
    If cc.GetCommandOutput(exe, yar & " " & test) Then
        frmAbout.DisplayData header & Replace(cc.result, yar, Empty, , , vbTextCompare)
    Else
        MsgBox "Failed to create process: " & cc.exitCode & " " & cc.result
    End If
    
End Sub

Private Sub mnuYaraHelp_Click()
     ShellDocument "https://yara.readthedocs.io/en/v3.11.0/writingrules.html"
End Sub

Private Sub mnuYaraMod_Click(Index As Integer)
    On Error Resume Next
    Dim fPath As String
    fPath = mnuYaraMod(Index).Tag
    Shell "notepad.exe """ & fPath & """", vbNormalFocus
End Sub

Private Sub sci_AutoCompleteEvent(className As String)

    Dim prev As String, match As Boolean
    Dim it As CIntellisenseItem, x, i2 As CIntellisenseItem, retType As String, tmp
    
    If CheckSubStatement() Then Exit Sub 'pe.linker_version. this event does not fire for pe.sections[n]. (ocx)
    
    prev = sci.PreviousWord
    
    For Each it In isense
    
        match = False
        If it.ObjName = className Or prev = it.ObjName Then match = True
        
        If match Then
            sci.ShowAutoComplete it.keywords
            sci.LoadCallTips Empty
            For Each x In it.protoTypes
                sci.AddCallTip CStr(x)
            Next
            Exit Sub
        End If
        
    Next
        
    'no match was found just show global obj types (also on ctrl-Tab)
    'If className = Empty Then
        tmp = Empty
        For Each it In isense
            tmp = tmp & it.ObjName & " "
        Next
        sci.ShowAutoComplete CStr(tmp & "int8 int16 int32 int8be int16be int32be uint8 uint16 uint32 uint8be uint16be uint32be")
        Exit Sub
    'End If
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Width < 10000 Then
        Me.Width = 10000
        Exit Sub
    End If
    
    If Me.Height < 6000 Then
        Me.Height = 6000
        Exit Sub
    End If
    
    If Me.Height < hSplit.Top Then
        hSplit.Top = Me.Height - 2000
    End If
    
    'sci.Width = Me.Width - sci.Left - 240
    Frame1.Left = Me.Width - Frame1.Width - 120
    txtSample.Width = Me.Width - txtSample.Left - 120 - Frame1.Width - 120
    List1.Top = Me.Height - List1.Height - 940
    
    List1.Width = Me.Width - List1.Left - 240
    hSplit.Width = List1.Width
    
    SizeWidthsToSplitter
    'SizeHeightsToHSplit - this would grow lower list on form resize..we want to keep save that only when hsplit moved
    
    sci.Height = List1.Top - sci.Top - 120
    lv.Height = sci.Height
    splitter.Height = sci.Height + 80
    hSplit.Top = List1.Top - 125
    
End Sub

'V splitter code
'------------------------------------------------
Private Sub splitter_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim a1&

    If Button = 1 Then 'The mouse is down
        If Capturing = False Then
            splitter.ZOrder
            SetCapture splitter.hwnd
            Capturing = True
        End If
        a1 = splitter.Left + x
        If MoveOk(a1) Then
            splitter.Left = a1
        End If
    End If
    
End Sub

Private Sub splitter_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Capturing Then
        ReleaseCapture
        Capturing = False
        SizeWidthsToSplitter
    End If
End Sub

Private Sub SizeWidthsToSplitter()
    On Error Resume Next
    Const buf = 15
    
    lv.Width = splitter.Left - lv.Left - buf
    pb.Left = splitter.Left + buf
    pb.Width = Me.Width - splitter.Left - buf
    
    sci.Left = splitter.Left + splitter.Width + buf / 2
    sci.Width = Me.Width - sci.Left - 200
    pb.Width = sci.Width
    'lv.ColumnHeaders(1).Width = lv.Width - 100
    
End Sub


Private Function MoveOk(x&) As Boolean  'Put in any limiters you desire
    MoveOk = False
    If x > 2400 And x < Me.Width - 2400 Then
        MoveOk = True
    End If
End Function

'------------------------------------------------
'V end splitter code


'H splitter code
'------------------------------------------------
Private Sub hsplit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim a1&

    If Button = 1 Then 'The mouse is down
        If Capturing = False Then
            hSplit.ZOrder
            SetCapture hSplit.hwnd
            Capturing = True
        End If
        With hSplit
            a1 = .Top + Y
            If hsplit_MoveOk(a1) Then
                .Top = a1
            End If
        End With
    End If
End Sub

Private Sub hsplit_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Capturing Then
        ReleaseCapture
        Capturing = False
        SizeHeightsToHSplit
    End If
End Sub

Private Sub SizeHeightsToHSplit()
    On Error Resume Next
    Dim tw As Integer 'Twips Width
    Dim th As Integer 'Twips Height
    tw = Screen.TwipsPerPixelX
    th = Screen.TwipsPerPixelY
    Const buf = 30
    List1.Top = hSplit.Top + hSplit.Height + buf
    sci.Height = List1.Top - (th * 60)
    lv.Height = sci.Height
    splitter.Height = sci.Height + 180
    List1.Height = Me.Height - List1.Top - 650
End Sub


Private Function hsplit_MoveOk(Y&) As Boolean  'Put in any limiters you desire
    hsplit_MoveOk = False
    If Y > 3000 And Y < Me.Height - 3000 Then
        hsplit_MoveOk = True
    End If
End Function

'------------------------------------------------
'end splitter code


Private Sub Form_Unload(Cancel As Integer)
    
    Dim v As VbMsgBoxResult
    
    If sci.isDirty Then
        v = MsgBox("Changes not saved save now?", vbYesNoCancel)
        If v = vbYes Then mnuSaveRules_Click
        If v = vbCancel Then
            Cancel = 1
            Exit Sub
        End If
    End If
    
    FormPos Me, True, True
    txtSample.SaveHistory
    TermLibYara
    
    SaveSetting "yara", "settings", "txtSample", txtSample
    SaveSetting "yara", "settings", "disassembler", disassembler
    SaveSetting "yara", "settings", "mnuShowBenchMarks", mnuShowBenchMarks.Checked
    
End Sub

'happens even on a left click..annoying with new form popup if right clicking...
'Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    Dim f As CYaraFile, fm As frmMatches
'    Set f = Item.Tag
'    Set fm = New frmMatches
'    fm.LoadRules f
'End Sub

Private Sub mnuCommentBlock_Click()
    On Error Resume Next
    Dim x() As String, i As Long
    If sci.SelLength = 0 Then
        MsgBox "Select block first"
        Exit Sub
    End If
    x = Split(sci.SelText, vbCrLf)
    For i = 0 To UBound(x)
        x(i) = "//" & x(i)
    Next
    sci.SelText = Join(x, vbCrLf)
End Sub

Private Sub mnuLoadRuleFile_Click()
    On Error Resume Next
    Dim f As String
    f = dlg.OpenDialog(curFile)
    If Len(f) = 0 Then Exit Sub
    curFile = Empty
    If sci.LoadFile(f) Then
        curFile = f
        AddToRecentList f, fso.GetBaseName(f)
        List1.AddItem "Loaded rule file: " & f
        load_ywbPath
    End If
End Sub

Function load_ywbPath()
    Dim tmp As String, a As Long, b As Long, pth As String
    tmp = sci.Text
    a = InStr(1, tmp, "ywbPath:", vbTextCompare)
    If a = 0 Then a = InStr(1, tmp, "ywb_path:", vbTextCompare)
    If a > 0 Then
        a = a + 9
        b = InStr(a, tmp, vbCr)
        If b > 0 Then pth = Mid(tmp, a, b - a) Else pth = Mid(tmp, a)
        pth = trim(pth)
        If LCase(Right(pth, 2)) = "-r" Then
            chkRecursive.value = 1
            pth = trim(Mid(pth, 1, Len(pth) - 2))
        End If
        If Left(pth, 2) = "./" Or Left(pth, 2) = ".\" Then
            pth = fso.GetParentFolder(curFile) & "\" & Mid(pth, 3)
        End If
        txtSample = pth
        'txtSample.ForeColor = vbBlack
    End If
End Function

Private Sub mnuSaveRules_Click()
    On Error Resume Next
    If Len(curFile) > 0 Then
        fso.writeFile curFile, sci.Text
        'If Not sci.SaveFile(curFile) Then
        If Err.Number <> 0 Then
            MsgBox "Save error: " & curFile
        Else
            AddToRecentList curFile, fso.GetBaseName(curFile)
        End If
    Else
        Dim f As String
        f = dlg.SaveDialog("", "")
        If Len(f) = 0 Then Exit Sub
        fso.writeFile f, sci.Text
        'If Not sci.SaveFile(f) Then
        If Err.Number <> 0 Then
            MsgBox "Save error: " & f
        Else
            curFile = f
            AddToRecentList curFile, fso.GetBaseName(curFile)
        End If
    End If
End Sub

Private Sub mnuSaveRulesAs_Click()
    On Error Resume Next
    Dim f As String
    f = dlg.SaveDialog() 'curFile)
    If Len(f) = 0 Then Exit Sub
    fso.writeFile f, sci.Text
    'If Not sci.SaveFile(f) Then
    If Err.Number <> 0 Then
        MsgBox "Save error: " & f
    Else
        curFile = f
        AddToRecentList curFile, fso.GetBaseName(curFile)
    End If
End Sub

Private Sub mnuUncommentBlock_Click()
    On Error Resume Next
    Dim x() As String, i As Long
    If sci.SelLength = 0 Then
        MsgBox "Select block first"
        Exit Sub
    End If
    x = Split(sci.SelText, vbCrLf)
    For i = 0 To UBound(x)
        If Left(x(i), 2) = "//" Then
            x(i) = Mid(x(i), 3)
        End If
    Next
    sci.SelText = Join(x, vbCrLf)
End Sub

Private Sub mnuWordWrap_Click()
    mnuWordWrap.Checked = Not mnuWordWrap.Checked
    sci.WordWrap = mnuWordWrap.Checked
End Sub

Private Sub sci_KeyUp(KeyCode As Long, Shift As Long)

    Dim last As String, c
    Dim obj As String, a As Long, i As CIntellisenseItem, e As CEnumList
    
    On Error GoTo hell
    
    'Debug.Print "sci_keyup: " & KeyCode
    If KeyCode > 36 And KeyCode < 41 Then Exit Sub 'ignore arrow keys
    
    c = Mid(sci.Text, sci.SelStart, 1)
    
    If c = "." Then
        If CheckSubStatement() Then Exit Sub
    End If
    
    If c = """" Then
        'entire line
        last = LCase(trim(Replace(Replace(sci.GetLineText(sci.CurrentLine()), vbTab, Empty), vbCrLf, Empty)))
        If last = "import """ Then
            'sci.ShowAutoComplete isense.dumpKeys(" ") 'ShowAutocomplete doesnt work after a " ?
            sci.DirectSCI.SendEditor SCI_AUTOCSHOW, 0, isense.dumpKeys(" ")
            Exit Sub
        End If
        'now just the last complete statement...
        last = GetPreceedingStatement()
        If last = "pe.version_info[""" Then
            last = "Comments CompanyName FileDescription FileVersion InternalName LegalCopyright LegalTrademarks OriginalFilename ProductName ProductVersion"
            sci.DirectSCI.SendEditor SCI_AUTOCSHOW, 0, last
            Exit Sub
        End If
    End If
    
    c = Mid(sci.Text, sci.SelStart - 1, 1) 'show auto complete wont show on "=" unless space? wtf
    
    If c = "=" Then 'handle displaying enums after " = "
        last = trim(GetPreceedingStatement())
        If Len(last) > 0 Then
            a = InStr(last, ".")
            If a > 0 Then
                obj = Mid(last, 1, a - 1)
                For Each i In isense
                    If i.ObjName = obj Then
                        For Each e In i.enums
                            If last Like e.ownerNoBrackets Then
                                 'Debug.Print "showing " & e.keywords
                                 sci.SelLength = 0
                                 sci.SelText = i.ObjName & "."
                                 sci.ShowAutoComplete e.keywords
                                 Exit Sub
                            End If
                        Next
                    End If
                Next
            End If
        End If
    End If
    
hell:
        
End Sub

'Private Sub txtSample_GotFocus()
'    On Error Resume Next
'    If txtSample.ForeColor <> vbBlack Then
'        txtSample.Text = Empty
'        txtSample.ForeColor = vbBlack
'    End If
'End Sub
'
'Private Sub txtSample_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    On Error Resume Next
'    txtSample_GotFocus
'    txtSample = Data.Files(1)
'    If Err.Number = 0 Then SaveSetting "yara", "settings", "txtSample", txtSample
'End Sub

Private Sub cmdAbort_Click()
    abort = True
    frmMain.pb.value = 0
End Sub

Private Sub cmdBrowse_Click()
    On Error Resume Next
    Dim f As String
    f = dlg.OpenDialog(fso.GetParentFolder(txtSample))
    If Len(f) = 0 Then Exit Sub
    'txtSample_GotFocus
    txtSample = f
End Sub

Private Sub cmdBrowseDir_Click()
    On Error Resume Next
    Dim f As String
    f = dlg.FolderDialog2(txtSample)
    If Len(f) = 0 Then Exit Sub
    'txtSample_GotFocus
    txtSample = f
End Sub

Private Sub LoadNewModuleDef(ByVal fPath As String)
    Dim i As Long
    
    If mnuYaraMod(0).Caption = "mnuMods0" Then
        i = 0
        mnuYaraMod(0).Caption = fso.GetBaseName(fPath)
        mnuYaraMod(0).Tag = fPath
    Else
        Load mnuYaraMod(mnuYaraMod.count + 1)
        i = mnuYaraMod.count
    End If
    
    mnuYaraMod(i).Caption = fso.GetBaseName(fPath)
    mnuYaraMod(i).Tag = fPath
    
End Sub
