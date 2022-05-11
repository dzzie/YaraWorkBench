VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      Height          =   375
      Left            =   9660
      TabIndex        =   9
      Top             =   540
      Width           =   1455
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Enabled         =   0   'False
      Height          =   315
      Left            =   9660
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtSample 
      Height          =   315
      Left            =   960
      OLEDropMode     =   1  'Manual
      TabIndex        =   7
      Top             =   60
      Width           =   8475
   End
   Begin VB.TextBox txtRuleFile 
      Height          =   315
      Left            =   2520
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   5
      Top             =   600
      Width           =   6915
   End
   Begin VB.TextBox txtRule 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   420
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Text            =   "Form2.frx":0000
      Top             =   960
      Width           =   10815
   End
   Begin VB.OptionButton optString 
      Caption         =   "String"
      Height          =   255
      Left            =   540
      TabIndex        =   3
      Top             =   600
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optFile 
      Caption         =   "File"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   540
      Width           =   735
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
      Height          =   3660
      Left            =   360
      TabIndex        =   0
      Top             =   3840
      Width           =   10875
   End
   Begin VB.Label Label1 
      Caption         =   "Sample"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   795
   End
   Begin VB.Label Rule 
      Caption         =   "Rule"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'void __stdcall SetCallBack(void* lpfnMsgHandler) {
Private Declare Sub SetCallBack Lib "yhelp.dll" (ByVal msgProc As Long)

'bool __stdcall testFile(char* rule, char* testString, wchar_t* wFName) {
Private Declare Function testFile Lib "yhelp.dll" (ByVal ruleTextOrFilePath As String, ByVal testDataStringOr As String, ByVal filePathIfExists As Long) As Boolean

'int __stdcall yr_initialize(void);
Private Declare Function yr_initialize Lib "libyara" () As Long

'int __stdcall yr_finalize(void);
Private Declare Function yr_finalize Lib "libyara" () As Long

Private Sub cmdAbort_Click()
    abort = True
End Sub

Private Sub cmdScan_Click()
    Dim r As String
    Dim rv As Boolean
    Dim t As String
    Dim filePath As String
    
    abort = False
    List1.Clear
    
    If FolderExists(txtSample) Then
        MsgBox "Files only for now"
        Exit Sub
    End If
    
    If Not FileExists(txtSample) Then
        List1.AddItem "File does not exist"
        Exit Sub
    End If
        
    r = txtRuleFile
    If optString.Value Then
        r = Trim(txtRule)
        If Len(r) = 0 Then
            List1.AddItem "Rule text can not be blank"
            Exit Sub
        End If
    Else
        If Not FileExists(r) Then
            List1.AddItem "Rule file not found"
            Exit Sub
        Else
            If FileLen(r) = 0 Then
                List1.AddItem "Rule file is empty"
                Exit Sub
            End If
        End If
    End If
    
    List1.AddItem "Starting scan " & Now
    SaveSetting "yara", "settings", "txtRule", Trim(txtRule)
    SaveSetting "yara", "settings", "txtRuleFile", Trim(txtRuleFile)

    StartBenchMark
    'if len(strMemScanTest) > 0 then
    '    List1.AddItem "Scanning memBuf sz:" & len(strMemScanTest)
    '    rv = testFile(r, strMemScanTest, 0)
    'Else
        List1.AddItem "Scanning file: " & txtSample.Text
        rv = testFile(r, vbNullString, StrPtr(txtSample.Text))  'MUST use .text for strptr
    'end if
    t = EndBenchMark
    
    If rv Then
        List1.AddItem "Signature FOUND"
    Else
        List1.AddItem "Signature NOT found"
    End If
    
    List1.AddItem "Scan complete " & t
    
End Sub

Private Sub Form_Load()
    Dim tmp As String
    Dim rv As Long
    
    'since source lives in a sub directory IDE requires this
    hYaraLib = LoadLibrary("libYara.dll")
    If hYaraLib = 0 Then hYaraLib = LoadLibrary(App.path & "\..\libYara.dll")
        
    If hYaraLib = 0 Then
        List1.AddItem "Could not load libYara.dll? in app.path or one up?"
        Exit Sub
    End If
    
    hHelperLib = LoadLibrary("yhelp.dll")
    If hHelperLib = 0 Then hHelperLib = LoadLibrary(App.path & "\..\yhelp.dll")

    If hHelperLib = 0 Then
        List1.AddItem "Could not load yhelp.dll? in app.path or one up?"
        Exit Sub
    End If

    rv = yr_initialize
    
    If rv <> 0 Then
        
        List1.AddItem "Error initilizing...Disabling scan"
        Exit Sub
    Else
        List1.AddItem "libYara Initilized ok"
    End If
    
    cmdScan.Enabled = True
    SetCallBack AddressOf vb_stdout
    
    txtRuleFile = GetSetting("yara", "settings", "txtRuleFile")
    tmp = GetSetting("yara", "settings", "txtSample")
    
    If FileExists(txtSample) Then
        txtSample = tmp
    Else
        txtSample = App.path & "\test.exe"
        If Not FileExists(txtSample) Then txtSample = App.path & "\..\test.exe"
    End If
    
    tmp = GetSetting("yara", "settings", "txtRule")
    If Len(tmp) > 0 Then txtRule = tmp
        
    tmp = GetSetting("yara", "settings", "txtRuleFile")
    If Len(tmp) > 0 Then txtRuleFile = tmp
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'so the IDE doesnt hang onto any dll refs on its own (in case we need to recompile..an hitting STOP in ide negates this protection..)
    If hHelperLib <> 0 Then FreeLibrary hHelperLib
    
    If hYaraLib <> 0 Then
        FreeLibrary hYaraLib
        yr_finalize
    End If
    
    SaveSetting "yara", "settings", "txtRule", Trim(txtRule)
    SaveSetting "yara", "settings", "txtRuleFile", Trim(txtRuleFile)
    
End Sub

Private Sub txtRuleFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    txtRuleFile = Data.Files(1)
    If Err.Number = 0 Then SaveSetting "yara", "settings", "txtRuleFile", txtRuleFile
End Sub

Private Sub txtSample_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    txtSample = Data.Files(1)
    If Err.Number = 0 Then SaveSetting "yara", "settings", "txtSample", txtSample
End Sub
