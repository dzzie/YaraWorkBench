VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CYaraScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public fileFilter As String
Public recursive As Boolean
Public errors As New CollectionEx
Public matchCount As Long

Public Event info(msgType As cb_type, msg As String)

Private Sub Class_Initialize()
    fileFilter = "*"
End Sub

Function scan(ByVal FileFolderOrMemStr As String, ByVal ruleFileOrStr As String, Optional dumpModules As Long = 0) As Collection
     
     Dim r As String
     Dim ff() As String
     Dim x
     Dim c As New Collection
     Dim yf As New CYaraFile
     Dim modDumpFile As String
     Dim fs As clsFileStream
     Dim i As Long
     
     abort = False
     matchCount = 0
     frmMain.pb.value = 0
     Set scan = c       'set our return value always valid even if empty

     
     If fso.FileExists(ruleFileOrStr) Then
        r = fso.ReadFile(ruleFileOrStr)
     Else
        r = ruleFileOrStr
     End If
    
     r = trim(r)
     If Len(r) = 0 Then
         RaiseEvent info(cb_error, "Rule text can not be blank")
         Exit Function
     End If
     
     'couple exit functions below but always set on start so should be ok... valid values: 0/1
     InitModuleDump dumpModules 'works through a call back adding to modCallbacks.ActiveYaraFile
     
     If fso.FileExists(FileFolderOrMemStr) And fso.GetExtension(FileFolderOrMemStr) = ".set" Then
        
        RaiseEvent info(cb_info, "Loading Sample set: " & FileFolderOrMemStr)
        Set fs = New clsFileStream
        fs.fOpen FileFolderOrMemStr, otreading
        
        Do While Not fs.EndOfFile
            If abort Then Exit Do
            If opt.ShowBenchMarks Then RaiseEvent info(cb_info, "  " & Rpad(fso.FileNameFromPath(CStr(x)), 25))
            Set yf = New CYaraFile
            Set yf.parent = Me
            Set modCallBacks.ActiveYaraFile = yf
            yf.File = fs.ReadLine()
            StartBenchMark
            yf.found = testFile(r, vbNullString, yf.File)
            If yf.found Then matchCount = matchCount + 1
            yf.ElapsedTime = EndBenchMark
            If opt.ShowBenchMarks Then RaiseEvent info(cb_update, yf.ElapsedTime)
            c.Add yf
            'i = i + 1
            'frmMain.pb.value = i todo: we dont know max...
        Loop
        
     ElseIf fso.FileExists(FileFolderOrMemStr) Then
     
        RaiseEvent info(cb_info, "Scanning file: " & FileFolderOrMemStr)
    
        Set yf.parent = Me
        yf.File = FileFolderOrMemStr
        Set modCallBacks.ActiveYaraFile = yf
        StartBenchMark
        'yf.found = testFile(r, vbNullString, StrPtr(FileFolderOrMemStr)) 'MUST use .text for strptr
        yf.found = testFile(r, vbNullString, FileFolderOrMemStr)   'MUST use .text for strptr
        If yf.found Then matchCount = matchCount + 1
        yf.ElapsedTime = EndBenchMark
        c.Add yf
        
     ElseIf fso.FolderExists(FileFolderOrMemStr) Then
     
        RaiseEvent info(cb_info, "Scanning folder: " & FileFolderOrMemStr)
        ff = fso.GetFolderFiles(FileFolderOrMemStr, fileFilter, , recursive)
        
        If AryIsEmpty(ff) Then
            RaiseEvent info(cb_error, "No files match filter  " & fileFilter & " recursive: " & recursive)
            Exit Function
        End If
        
        frmMain.pb.Max = UBound(ff) + 1
        RaiseEvent info(cb_info, "Files found: " & UBound(ff))
        
        For Each x In ff
            If abort Then Exit Function
            If opt.ShowBenchMarks Then RaiseEvent info(cb_info, "  " & Rpad(fso.FileNameFromPath(CStr(x)), 25))
            Set yf = New CYaraFile
            Set yf.parent = Me
            Set modCallBacks.ActiveYaraFile = yf
            yf.File = x
            StartBenchMark
            yf.found = testFile(r, vbNullString, yf.File)   'MUST use .text for strptr
            If yf.found Then matchCount = matchCount + 1
            yf.ElapsedTime = EndBenchMark
            If opt.ShowBenchMarks Then RaiseEvent info(cb_update, yf.ElapsedTime)
            c.Add yf
            i = i + 1
            frmMain.pb.value = i
            DoEvents
        Next
        
    Else
        
        RaiseEvent info(cb_info, "Scanning 0x" & Hex(Len(FileFolderOrMemStr)) & " byte MemStr")

        Set modCallBacks.ActiveYaraFile = yf
        Set yf.parent = Me
        yf.File = FileFolderOrMemStr
        yf.isMemStr = True
        StartBenchMark
        yf.found = testFile(r, FileFolderOrMemStr, 0)
        If yf.found Then matchCount = matchCount + 1
        yf.ElapsedTime = EndBenchMark
        c.Add yf
        
    End If
    
    frmMain.pb.value = 0
    InitModuleDump 0 'couple exit functions above but always set on start so should be ok
                     'as long as scans are only started through this func, added overhead is all if not..
    
End Function

 

