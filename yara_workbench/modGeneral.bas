Attribute VB_Name = "modGeneral"
Option Explicit

Type opts
    ShowBenchMarks As Boolean
End Type

Global opt As opts

Public Const SCI_AUTOCSETIGNORECASE = 2115  ' Establece si afectan mays/min a la búsqueda en la lista
Public Const SCI_AUTOCGETIGNORECASE = 2116  ' Devuelve si afectan mays/min a la búsqueda en la lista
Public Const SCI_AUTOCSETAUTOHIDE = 2118    ' Establece si la lista se oculta si no existen coindicencias.
Public Const SCI_AUTOCGETAUTOHIDE = 2119    ' Devuelve el comportamiento de la lista si no existen coincidencias.
Public Const SCI_AUTOCSHOW = 2100
'Public Const SCI_AUTOCSETORDER = 2660
'Public Const SCI_AUTOCGETORDER = 2661
'Public Const SC_ORDER_PRESORTED = 0 'requires that the list be provided in alphabetical sorted order.
'Public Const SC_ORDER_PERFORMSORT = 1 'Sorting the list can be done by Scintilla instead of the application
'
'Public Const SC_CASEINSENSITIVEBEHAVIOUR_RESPECTCASE = 0
'Public Const SC_CASEINSENSITIVEBEHAVIOUR_IGNORECASE = 1
'Public Const SCI_AUTOCSETCASEINSENSITIVEBEHAVIOUR = 2634
'Public Const SCI_AUTOCGETCASEINSENSITIVEBEHAVIOUR = 2635
Global Const LANG_US = &H409

Enum sci_B_Indexes
    SCE_B_DEFAULT = 0
    SCE_B_COMMENT = 1
    SCE_B_NUMBER = 2
    SCE_B_KEYWORD = 3
    SCE_B_STRING = 4
    SCE_B_PREPROCESSOR = 5
    SCE_B_OPERATOR = 6
    SCE_B_IDENTIFIER = 7
    SCE_B_DATE = 8
    SCE_B_STRINGEOL = 9
    SCE_B_KEYWORD2 = 10
    SCE_B_KEYWORD3 = 11
    SCE_B_KEYWORD4 = 12
    SCE_B_CONSTANT = 13
    SCE_B_ASM = 14
    SCE_B_LABEL = 15
    SCE_B_ERROR = 16
    SCE_B_HEXNUMBER = 17
    SCE_B_BINNUMBER = 18
    SCE_B_COMMENTBLOCK = 19
    SCE_B_DOCLINE = 20
    SCE_B_DOCBLOCK = 21
    SCE_B_DOCKEYWORD = 22
End Enum

Enum sci_C_indexes 'cpp lexer supported styles (ilang 3)
     SCE_C_DEFAULT = 0
     SCE_C_COMMENT = 1         'multiline   /*
     SCE_C_COMMENTLINE = 2     'single line //
     SCE_C_COMMENTDOC = 3
     SCE_C_NUMBER = 4          'offsets and bytecode (unless bytecode is hex alpha only)
     SCE_C_WORD = 5            ' keywords0
     SCE_C_STRING = 6          ' double quoted strings
     SCE_C_CHARACTER = 7       'single quoted strings
     SCE_C_UUID = 8
     SCE_C_PREPROCESSOR = 9
     SCE_C_OPERATOR = 10
     SCE_C_IDENTIFIER = 11       'most misc words (also bug with alpha only bytecode)
     SCE_C_STRINGEOL = 12
     SCE_C_VERBATIM = 13
     SCE_C_REGEX = 14
     SCE_C_COMMENTLINEDOC = 15
     SCE_C_WORD2 = 16            'keywords1
     SCE_C_COMMENTDOCKEYWORD = 17
     SCE_C_COMMENTDOCKEYWORDERROR = 18
     SCE_C_GLOBALCLASS = 19      'keywords3
     SCE_C_STRINGRAW = 20
     SCE_C_TRIPLEVERBATIM = 21
     SCE_C_HASHQUOTEDSTRING = 22
     SCE_C_PREPROCESSORCOMMENT = 23
     SCE_C_PREPROCESSORCOMMENTDOC = 24
End Enum

Enum sci_lexers
    SCLEX_NONE = 0
    SCLEX_CPP = 3
    SCLEX_HTML = 4
    SCLEX_XML = 5
    SCLEX_SQL = 7
    SCLEX_VB = 8
    SCLEX_Asasm = 9
    SCLEX_ASM = 34
    SCLEX_CPPNOCASE = 35
    SCLEX_PHPSCRIPT = 69
End Enum

'scintinilla.iface: Styles in range 32..38 are predefined for parts of the UI and are not used as normal styles.
Const STYLE_DEFAULT = 32
Const STYLE_LINENUMBER = 33
Const STYLE_BRACELIGHT = 34
Const STYLE_BRACEBAD = 35
Const STYLE_CONTROLCHAR = 36
Const STYLE_INDENTGUIDE = 37
Const STYLE_CALLTIP = 38
    
Const SC_MARK_CIRCLE = 0
Const SC_MARK_ARROW = 2
Const SC_MARK_BACKGROUND = 22

Public Enum hexOutFormats
    hoDump
    hoSpaced
    hoHexOnly
End Enum


Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Enum AutoCompleteFlags
    SHACF_AUTOAPPEND_FORCE_OFF = &H80000000
    SHACF_AUTOAPPEND_FORCE_ON = &H40000000
    SHACF_AUTOSUGGEST_FORCE_OFF = &H20000000
    SHACF_AUTOSUGGEST_FORCE_ON = &H10000000
    SHACF_DEFAULT = &H0
    SHACF_FILESYSTEM = &H1
    SHACF_URLHISTORY = &H2
    SHACF_URLMRU = &H4
    SHACF_USETAB = &H8
    SHACF_URLALL = (SHACF_URLHISTORY Or SHACF_URLMRU)
End Enum

Private OnBits(0 To 31) As Long
Private crc_table() As Long

Public Declare Sub SHAutoComplete Lib "Shlwapi.dll" (ByVal hwndEdit As Long, ByVal dwFlags As AutoCompleteFlags)
Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public ChildWindows As Collection
Public classFilter As String

Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = (KEY_READ)
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Enum hKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Private Enum dataType
    REG_BINARY = 3                     ' Free form binary
    REG_DWORD = 4                      ' 32-bit number
    'REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
    'REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
    REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
    'REG_MULTI_SZ = 7                   ' Multiple Unicode strings
    REG_SZ = 1                         ' Unicode nul terminated string
End Enum


Function GetChildWindows(Optional hwnd As Long = 0, Optional classNameFilter As String) As Collection  'of CWindow
    
    Set ChildWindows = New Collection
    classFilter = classNameFilter
    Call EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)
    Set GetChildWindows = ChildWindows

End Function

Private Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim c As New CWindow
    c.hwnd = hwnd
    If Not IsObject(ChildWindows) Then Set ChildWindows = New Collection
    If Len(classFilter) > 0 Then
        If InStr(1, c.className, classFilter, vbTextCompare) > 0 Then ChildWindows.Add c 'module level collection object...
    Else
        ChildWindows.Add c 'module level collection object...
    End If
    EnumChildProc = 1  'continue enum
End Function

Function register_FileExt() As Boolean
    
    Dim homeDir As String
    Dim tmp As String, cmd As String
    
    
    homeDir = App.path & "\ywb.exe"
    If Not fso.FileExists(homeDir) Then Exit Function
    cmd = "cmd /c ftype Yara.Document=""" & homeDir & """ %1 && assoc .yara=Yara.Document"
    
    On Error Resume Next
    Shell cmd, vbHide
    
    Dim wsh As Object 'WshShell
    Set wsh = CreateObject("WScript.Shell")
    If Not wsh Is Nothing Then
        wsh.RegWrite "HKCR\Yara.Document\DefaultIcon\", homeDir & ",0"
    End If
    
    tmp = ReadRegValue("HKLM\SOFTWARE\Classes\Yara.Document\shell\open\command")
    register_FileExt = (InStr(1, tmp, "ywb.exe", vbTextCompare) > 0)
    
End Function
    
Private Function stdPath(sIn, ByRef hive As hKey) As String
    Dim tmp
    
    stdPath = Replace(sIn, "/", "\")
    
    tmp = Split(stdPath, "\")
    Select Case LCase(tmp(0))
        Case "hklm", "hkey_local_machine": hive = HKEY_LOCAL_MACHINE
        Case "hkcu", "hkey_current_user": hive = HKEY_CURRENT_USER
        Case "hkcr", "hkey_classes_root": hive = HKEY_CLASSES_ROOT
        Case "hku", "hkey_users": hive = HKEY_USERS
    End Select
    
    tmp(0) = Empty
    stdPath = Join(tmp, "\")
    stdPath = Replace(stdPath, "\\", "\")
    
    If Left(stdPath, 1) = "\" Then stdPath = Mid(stdPath, 2, Len(stdPath))
    If Right(stdPath, 1) <> "\" Then stdPath = stdPath & "\"
    
End Function

Function ReadRegValue(path, Optional KeyName = "")
     
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    'Dim ret As Long
    'retrieve nformation about the key
    Dim p As String
    Dim hive As hKey
    Dim handle As Long
    Dim ret

    p = stdPath(path, hive)
    RegOpenKeyEx hive, p, 0, KEY_READ, handle
    lResult = RegQueryValueEx(handle, CStr(KeyName), 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            strBuf = String(lDataBufSize, Chr$(0))
            lResult = RegQueryValueEx(handle, CStr(KeyName), 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then ret = Replace(strBuf, Chr$(0), "")
        ElseIf lValueType = REG_BINARY Then
            Dim strData As Integer
            lResult = RegQueryValueEx(handle, CStr(KeyName), 0, 0, strData, lDataBufSize)
            If lResult = 0 Then ret = strData
        ElseIf lValueType = REG_DWORD Then
            Dim x As Long
            lResult = RegQueryValueEx(handle, CStr(KeyName), 0, 0, x, lDataBufSize)
            ret = x
        ElseIf lValueType = REG_EXPAND_SZ Then
            strBuf = String(lDataBufSize, Chr$(0))
            lResult = RegQueryValueEx(handle, CStr(KeyName), 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then ret = Replace(strBuf, Chr$(0), "")

        'Else
        '    MsgBox "UnSupported Type " & lValueType
        End If
    End If
    RegCloseKey handle
    
    ReadRegValue = ret
    
End Function
    
Function GetKeyState(k As KeyCodeConstants) As Boolean
    If GetAsyncKeyState(k) Then GetKeyState = True
End Function

Public Function GetShortName(sFile As String) As String
    Dim sShortFile As String * 67
    Dim lResult As Long
    Dim iCreated As Boolean
    
    'the path must actually exist to get the short path name!
    If Not fso.FolderExists(sFile) Then
        If Not fso.FileExists(sFile) Then
            fso.writeFile sFile, ""
            iCreated = True
        End If
    End If
    
    lResult = GetShortPathName(sFile, sShortFile, Len(sShortFile))
    GetShortName = Left$(sShortFile, lResult)
    If Len(GetShortName) = 0 Then GetShortName = sFile
    If iCreated Then fso.DeleteFile sFile

End Function
     
     
'START_HIDDEN = 0, START_NORMAL = 4, START_MINIMIZED = 2, START_MAXIMIZED = 3
Public Function ShellDocument(sDocName As String, Optional ByVal action As String = "Open", Optional ByVal Parameters As String = vbNullString, Optional ByVal Directory As String = vbNullString, Optional ByVal WindowState As Long = 4) As Boolean
    If ShellExecute(&O0, action, sDocName, Parameters, Directory, WindowState) >= 33 Then ShellDocument = True
End Function

Function TopMost(frm As Object, Optional ontop As Boolean = True)
    On Error Resume Next
    Dim s
    s = IIf(ontop, HWND_TOPMOST, HWND_NOTOPMOST)
    SetWindowPos frm.hwnd, s, frm.Left / 15, frm.Top / 15, frm.Width / 15, frm.Height / 15, 0
End Function

Sub setVB(i As sci_B_Indexes, _
    Optional fore As ColorConstants = vbBlack, _
    Optional back As ColorConstants = vbWhite, _
    Optional font As String = "Courier New", _
    Optional Size As Long = 11, _
    Optional bold As Boolean = False, _
    Optional italic As Boolean = False, _
    Optional underline As Boolean = False, _
    Optional Visible As Boolean = True)

    With frmMain.sci.DirectSCI
           .StyleSetBold i, IIf(bold, 1, 0)
           .StyleSetItalic i, IIf(italic, 1, 0)
           .StyleSetUnderline i, IIf(underline, 1, 0)
           .StyleSetVisible i, IIf(Visible, 1, 0)
           .StyleSetFont i, font
           .StyleSetFore i, fore
           .StyleSetBack i, back
           .StyleSetSize i, Size
    End With
    
End Sub

Sub SyntaxColor(sci As SciSimple, Optional additionalKeywords As String)

    With sci.DirectSCI

'         If scm = scm_text Then
'            If .GetLexer <> SCLEX_NONE Then
'                .SetLexer SCLEX_NONE
'                .ClearDocumentStyle
'                .StyleClearAll
'                .StyleSetFore STYLE_DEFAULT, vbBlack
'                .StyleSetSize STYLE_DEFAULT, 11
'                .StyleSetFont STYLE_DEFAULT, "Courier New"
'                '.StyleSetBits 5
'                .Colourise 0, -1
'            End If
'
'         ElseIf scm = scm_pcode Then
         
            'If .GetLexer = SCLEX_CPP Then Exit Sub
            
            .ClearDocumentStyle
            .StyleClearAll
            '.StyleSetBits 5
            .SetLexer SCLEX_CPP
            .StyleSetFore STYLE_DEFAULT, vbBlack
            .StyleSetSize STYLE_DEFAULT, 11
            .StyleSetFont STYLE_DEFAULT, "Courier New"
            
            'calls
            .SetKeyWords 0, "all and any ascii at condition contains uint32 uint8be uint16be uint32be wide xor " & _
                            "entrypoint false filesize fullword for global in import include int8 int16 int32 int8be int16be " & _
                            "int32be matches meta nocase not or of private rule strings them true uint8 uint16"

            
            'branches, equality tests, for loops
            '.SetKeyWords 1, ""
            
            'If Len(additionalKeywords) = 0 Then additionalKeywords = ""
                            
            'common crypto math, file access
            '.SetKeyWords 3, additionalKeywords
            
            '.StyleSetFore SCE_C_NUMBER, &H808000
            '.StyleSetFore SCE_C_IDENTIFIER, &H808000
            
            .StyleSetFore SCE_C_COMMENT, &H5500 '8000
            .StyleSetFore SCE_C_COMMENTLINE, &H5500 '8000
            .StyleSetFore SCE_C_WORD, &H800000
            .StyleSetFore SCE_C_WORD2, vbRed
            .StyleSetFore SCE_C_GLOBALCLASS, &H25208D '808000  'keywords3
            .StyleSetFore SCE_C_STRING, &H800080
            .StyleSetFore SCE_C_CHARACTER, &H800080
            
            setVB SCE_B_HEXNUMBER, vbGreen
                 setVB SCE_B_COMMENT, &H5500
                 
            .StyleSetBold SCE_C_WORD, True
            .StyleSetBold SCE_C_WORD2, True
            .StyleSetBold SCE_C_GLOBALCLASS, True
            '.StyleSetBack SCE_C_GLOBALCLASS, vbYellow
             
            .Colourise 0, -1
          
'          ElseIf scm = scm_vb Then
'
'            .SetLexer SCLEX_VB
'            .SetKeyWords 0, "object begin beginproperty endproperty end" 'must be lowercase
'            '.SetKeyWords 1, "begin"
'            '.SetKeyWords 2, "beginproperty"
'            '.SetKeyWords 3, "endproperty end"
'            '.SetKeyWords 4, "end"
'
'            setVB SCE_B_COMMENT, &H5500
'            setVB SCE_B_CONSTANT, vbRed
'           setVB SCE_B_HEXNUMBER, vbGreen
'            setVB SCE_B_IDENTIFIER, &HA00000 'dark blue
'            setVB SCE_B_KEYWORD, vbRed, , , , True
'            'setVB SCE_B_KEYWORD2, vbMagenta
'            'setVB SCE_B_KEYWORD3, vbYellow
'            'setVB SCE_B_KEYWORD4, vbCyan
'            setVB SCE_B_STRING, &HC00090
'            setVB SCE_B_LABEL, &HFF00FF 'magenta
'
'            .Colourise 0, -1
'
'          ElseIf scm = scm_vbsrc Then
'
'            'If .GetLexer = SCLEX_VB Then Exit Sub 'used in two ways...can not cache
'
'            .SetLexer SCLEX_VB
'            'keywords must be lowercase
'            .SetKeyWords 0, "and begin case call class continue do each else elseif end erase error event exit false for function get gosub goto if implement in load loop lset me mid new next not nothing on or property raiseevent rem resume return rset select set stop sub then to true unload until wend while with withevents attribute alias as boolean byref byte byval const compare currency date declare dim double enum explicit friend global integer let lib long module object option optional preserve private public redim single static string type variant"
'            setVB SCE_B_COMMENT, &H5500
'            setVB SCE_B_KEYWORD, &HA00000
'            setVB SCE_B_STRING, &HC00090
'
'            .Colourise 0, -1
'          End If
          

  
  End With
  
End Sub

Function lpad(v, Optional L As Long = 8, Optional char As String = " ")
    On Error GoTo hell
    Dim x As Long
    x = Len(v)
    If x < L Then
        lpad = String(L - x, char) & v
    Else
hell:
        lpad = v
    End If
End Function

Function Rpad(v, Optional L As Long = 8, Optional char As String = " ")
    On Error GoTo hell
    Dim x As Long
    x = Len(v)
    If x < L Then
        Rpad = v & String(L - x, char)
    Else
hell:
        Rpad = v
    End If
End Function


Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub


Function HexDump(bAryOrStrData, Optional va As Long, Optional ByVal Length As Long = -1, Optional ByVal startAt As Long = 1, Optional hexFormat As hexOutFormats = hoDump) As String
    Dim s() As String, chars As String, tmp As String
    On Error Resume Next
    Dim ary() As Byte
    Dim offset As Long
    Const LANG_US = &H409
    Dim i As Long, tt, h, x
    Dim hexOnly As Long
    
    offset = va '0
    If hexFormat <> hoDump Then hexOnly = 1
    
    If TypeName(bAryOrStrData) = "Byte()" Then
        ary() = bAryOrStrData
    Else
        ary = StrConv(CStr(bAryOrStrData), vbFromUnicode, LANG_US)
    End If
    
    If startAt < 1 Then startAt = 1
    If Length < 1 Then Length = -1
    
    While startAt Mod 16 <> 0
        startAt = startAt - 1
    Wend
    
    startAt = startAt + 1
    
    chars = "   "
    For i = startAt To UBound(ary) + 1
        tt = Hex(ary(i - 1))
        If Len(tt) = 1 Then tt = "0" & tt
        tmp = tmp & tt & " "
        x = ary(i - 1)
        'chars = chars & IIf((x > 32 And x < 127) Or x > 191, Chr(x), ".") 'x > 191 causes \x0 problems on non us systems... asc(chr(x)) = 0
        chars = chars & IIf((x > 32 And x < 127), Chr(x), ".")
        If i > 1 And i Mod 16 = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            If hexOnly = 0 Then
                push s, h & "   " & tmp & chars
            Else
                push s, tmp
            End If
            offset = offset + 16
            tmp = Empty
            chars = "   "
        End If
        If Length <> -1 Then
            Length = Length - 1
            If Length = 0 Then Exit For
        End If
    Next
    
    'if read length was not mod 16=0 then
    'we have part of line to account for
    If tmp <> Empty Then
        If hexOnly = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            h = h & "   " & tmp
            While Len(h) <= 56: h = h & " ": Wend
            push s, h & chars
        Else
            push s, tmp
        End If
    End If
    
    HexDump = Join(s, vbCrLf)
    
    If hexOnly <> 0 Then
        If hexFormat = hoHexOnly Then HexDump = Replace(HexDump, " ", "")
        HexDump = Replace(HexDump, vbCrLf, "")
    End If
    
End Function

Function isIde() As Boolean
' Brad Martinez  http://www.mvps.org/ccrp
    On Error GoTo out
    Debug.Print 1 / 0
out: isIde = Err
End Function

'this function will convert any of the following to a byte array:
'   read a file if path supplied and allowFilePaths = true
'   byte(), integer() or long() arrays
'   all other data types it will attempt to convert them to string, then to byte array
'   if the data type you pass can not be converted with cstr() it will throw an error.
'   no other types make sense to support explicitly
'   this assumes all arrays are 0 based..
Function LoadData(fileStringOrByte, Optional allowFilePaths As Boolean = True) As Byte()
    
    Dim f As Long
    Dim Size As Long
    Dim b() As Byte
    Dim L() As Long    ' must cast to specific array type or
    Dim i() As Integer ' else you are reading part of the variant structure..
    
    If TypeName(fileStringOrByte) = "Byte()" Then
        b() = fileStringOrByte
    ElseIf TypeName(fileStringOrByte) = "Integer()" Then
        i() = fileStringOrByte
        ReDim b((UBound(i) * 2) - 1)
        CopyMemory ByVal VarPtr(b(0)), ByVal VarPtr(i(0)), UBound(b) + 1
    ElseIf TypeName(fileStringOrByte) = "Long()" Then
        L() = fileStringOrByte
        ReDim b((UBound(L) * 4) - 1)
        CopyMemory ByVal VarPtr(b(0)), ByVal VarPtr(L(0)), UBound(b) + 1
    ElseIf allowFilePaths And fso.FileExists(CStr(fileStringOrByte)) Then
         f = FreeFile
         Open fileStringOrByte For Binary As f
         ReDim b(LOF(f) - 1)
         Get f, , b()
         Close f
    Else
        b() = StrConv(CStr(fileStringOrByte), vbFromUnicode, LANG_US)
    End If
    
    LoadData = b()
    
End Function

'rc4
Public Function blowfish(fileByteOrString As Variant, ByVal password As Variant, Optional strRet As Boolean = True)
    
    On Error Resume Next
    Dim RB(0 To 255) As Integer, x As Long, Y As Long, Z As Long, key() As Byte, ByteArray() As Byte, temp As Byte
    
    Dim plen As Long
    
    ByteArray() = LoadData(fileByteOrString)
    
    If TypeName(password) = "Byte()" Then
        key() = password
        If UBound(key) > 255 Then ReDim Preserve key(255)
    Else
        If Len(password) = 0 Then
            Exit Function
        End If

        If Len(password) > 256 Then
            key() = StrConv(Left$(CStr(password), 256), vbFromUnicode, LANG_US)
        Else
            key() = StrConv(CStr(password), vbFromUnicode, LANG_US)
        End If
    End If
    
    plen = UBound(key) + 1
 
    'Debug.Print "key=" & HexDump(Key)
    'Debug.Print "data=" & HexDump(ByteArray)
    
    For x = 0 To 255
        RB(x) = x
    Next x
    
    x = 0
    Y = 0
    Z = 0
    For x = 0 To 255
        Y = (Y + RB(x) + key(x Mod plen)) Mod 256
        temp = RB(x)
        RB(x) = RB(Y)
        RB(Y) = temp
    Next x
    
    x = 0
    Y = 0
    Z = 0
    For x = 0 To UBound(ByteArray)
        Y = (Y + 1) Mod 256
        Z = (Z + RB(Y)) Mod 256
        temp = RB(Y)
        RB(Y) = RB(Z)
        RB(Z) = temp
        ByteArray(x) = ByteArray(x) Xor (RB((RB(Y) + RB(Z)) Mod 256))
    Next x
    
    If strRet Then
        blowfish = StrConv(ByteArray, vbUnicode, LANG_US)
    Else
        blowfish = ByteArray
    End If
    
End Function


'------------------------ crc32 -----------------------------------
Function md4(bAryOrString) As String

'    vb native implementation...
    Dim c As Long, n As Long, x As Long
    Dim b() As Byte

    c = -1
    If AryIsEmpty(crc_table) Then make_crc_table

    If TypeName(bAryOrString) = "Byte()" Then
        b() = bAryOrString
    Else
        b() = StrConv(CStr(bAryOrString), vbFromUnicode, LANG_US)
    End If

    For n = 0 To UBound(b)
        c = crc_table((c Xor b(n)) And &HFF) Xor rshift(c, 8)
    Next

    md4 = Hex(c Xor &HFFFFFFFF)
    
End Function

Private Sub make_crc_table()
    Dim c As Long, n As Long, k As Long

    ReDim crc_table(256)

    For n = 0 To 255
          c = n
          For k = 0 To 7
                If c And 1 Then
                     c = &H18809425 Xor rshift(c) 'modified constant
                Else
                    c = rshift(c)
                End If
           Next
          crc_table(n) = c
    Next

End Sub

Public Function lshift(ByVal value As Long, Optional ByVal Shift As Integer = 1) As Long
    MakeOnBits
    If (value And (2 ^ (31 - Shift))) Then 'GoTo OverFlow
        lshift = ((value And OnBits(31 - (Shift + 1))) * (2 ^ (Shift))) Or &H80000000
    Else
        lshift = ((value And OnBits(31 - Shift)) * (2 ^ Shift))
    End If
End Function

Public Function rshift(ByVal value As Long, Optional ByVal Shift As Integer = 1) As Long
    Dim hi As Long
    MakeOnBits
    If (value And &H80000000) Then hi = &H40000000
    rshift = (value And &H7FFFFFFE) \ (2 ^ Shift)
    rshift = (rshift Or (hi \ (2 ^ (Shift - 1))))
End Function

Private Sub MakeOnBits()
    Dim j As Integer, v As Long

    For j = 0 To 30
        v = v + (2 ^ j)
        OnBits(j) = v
    Next j

    OnBits(j) = v + &H80000000

End Sub
'-------------------------------------------------------------------

Function toBytes(x As String) As Byte()

    Dim fx() As Byte
    Dim i As Long, j As Long
    
    If Len(x) = 0 Then Exit Function
    If Len(x) Mod 2 <> 0 Then Exit Function
    
    ReDim fx((Len(x) / 2) - 1)
    
    For i = 1 To Len(x) Step 2
        fx(j) = CInt("&h" & Mid(x, i, 2))
        j = j + 1
    Next
    
    toBytes = fx()

End Function

Function lbCopy(lstBox As Object) As String
    
    Dim i As Long
    Dim tmp() As String
    
    For i = 0 To lstBox.ListCount
        push tmp, lstBox.List(i)
    Next
    
    lbCopy = Join(tmp, vbCrLf)
    
End Function

Function lvGetAllElements(lv As Object) As String
    Dim ret() As String, i As Integer, tmp As String
    Dim li 'As ListItem
    
    On Error Resume Next
    
    For i = 1 To lv.ColumnHeaders.count
        tmp = tmp & lv.ColumnHeaders(i).Text & vbTab
    Next
    
    push ret, tmp
    push ret, String(50, "-")
        
    For Each li In lv.ListItems
        tmp = li.Text & vbTab
        For i = 1 To lv.ColumnHeaders.count - 1
            tmp = tmp & li.subItems(i) & vbTab
        Next
        push ret, tmp
    Next
    
    lvGetAllElements = Join(ret, vbCrLf)
    
End Function


Function CountOccurances(it, Find) As Integer
    Dim tmp() As String
    If InStr(1, it, Find, vbTextCompare) < 1 Then CountOccurances = 0: Exit Function
    tmp = Split(it, Find, , vbTextCompare)
    CountOccurances = UBound(tmp) + 1
End Function

Sub FormPos(fform As Object, Optional andSize As Boolean = False, Optional save_mode As Boolean = False)
    
    On Error Resume Next
    
    Dim f, sz, i, ff, def
    f = Split(",Left,Top,Height,Width", ",")
    
    If fform.WindowState = vbMinimized Then Exit Sub
    If andSize = False Then sz = 2 Else sz = 4
    
    For i = 1 To sz
        If save_mode Then
            ff = CallByName(fform, f(i), VbGet)
            SaveSetting App.EXEName, fform.name & ".FormPos", f(i), ff
        Else
            def = CallByName(fform, f(i), VbGet)
            ff = GetSetting(App.EXEName, fform.name & ".FormPos", f(i), def)
            CallByName fform, f(i), VbLet, ff
        End If
    Next
    
End Sub
