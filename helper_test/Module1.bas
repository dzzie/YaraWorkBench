Attribute VB_Name = "Module1"
Option Explicit

Enum cb_type
    cb_output = 0
    cb_info = 1
    cb_match = 2
    cb_update = 3
    cb_error = 4
    cb_matchInfo = 5
End Enum

Enum cb_msgs
    CALLBACK_MSG_RULE_MATCHING = 1
    CALLBACK_MSG_RULE_NOT_MATCHING = 2
    CALLBACK_MSG_SCAN_FINISHED = 3
    CALLBACK_MSG_IMPORT_MODULE = 4
    CALLBACK_MSG_MODULE_IMPORTED = 5
End Enum

Enum vt_errs
    ERROR_SUCCESS = 0
    ERROR_INSUFFICIENT_MEMORY = 1
    ERROR_COULD_NOT_ATTACH_TO_PROCESS = 2
    ERROR_COULD_NOT_OPEN_FILE = 3
    ERROR_COULD_NOT_MAP_FILE = 4
    ERROR_INVALID_FILE = 6
    ERROR_CORRUPT_FILE = 7
    ERROR_UNSUPPORTED_FILE_VERSION = 8
    ERROR_INVALID_REGULAR_EXPRESSION = 9
    ERROR_INVALID_HEX_STRING = 10
    ERROR_SYNTAX_ERROR = 11
    ERROR_LOOP_NESTING_LIMIT_EXCEEDED = 12
    ERROR_DUPLICATED_LOOP_IDENTIFIER = 13
    ERROR_DUPLICATED_IDENTIFIER = 14
    ERROR_DUPLICATED_TAG_IDENTIFIER = 15
    ERROR_DUPLICATED_META_IDENTIFIER = 16
    ERROR_DUPLICATED_STRING_IDENTIFIER = 17
    ERROR_UNREFERENCED_STRING = 18
    ERROR_UNDEFINED_STRING = 19
    ERROR_UNDEFINED_IDENTIFIER = 20
    ERROR_MISPLACED_ANONYMOUS_STRING = 21
    ERROR_INCLUDES_CIRCULAR_REFERENCE = 22
    ERROR_INCLUDE_DEPTH_EXCEEDED = 23
    ERROR_WRONG_TYPE = 24
    ERROR_EXEC_STACK_OVERFLOW = 25
    ERROR_SCAN_TIMEOUT = 26
    ERROR_TOO_MANY_SCAN_THREADS = 27
    ERROR_CALLBACK_ERROR = 28
    ERROR_INVALID_ARGUMENT = 29
    ERROR_TOO_MANY_MATCHES = 30
    ERROR_INTERNAL_FATAL_ERROR = 31
    ERROR_NESTED_FOR_OF_LOOP = 32
    ERROR_INVALID_FIELD_NAME = 33
    ERROR_UNKNOWN_MODULE = 34
    ERROR_NOT_A_STRUCTURE = 35
    ERROR_NOT_INDEXABLE = 36
    ERROR_NOT_A_FUNCTION = 37
    ERROR_INVALID_FORMAT = 38
    ERROR_TOO_MANY_ARGUMENTS = 39
    ERROR_WRONG_ARGUMENTS = 40
    ERROR_WRONG_RETURN_TYPE = 41
    ERROR_DUPLICATED_STRUCTURE_MEMBER = 42
    ERROR_EMPTY_STRING = 43
    ERROR_DIVISION_BY_ZERO = 44
    ERROR_REGULAR_EXPRESSION_TOO_LARGE = 45
    ERROR_TOO_MANY_RE_FIBERS = 46
    ERROR_COULD_NOT_READ_PROCESS_MEMORY = 47
    ERROR_INVALID_EXTERNAL_VARIABLE_TYPE = 48
    ERROR_REGULAR_EXPRESSION_TOO_COMPLEX = 49
    ERROR_INVALID_MODULE_NAME = 50
    ERROR_TOO_MANY_STRINGS = 51
    ERROR_INTEGER_OVERFLOW = 52
    ERROR_CALLBACK_REQUIRED = 53
    ERROR_INVALID_OPERAND = 54
    ERROR_COULD_NOT_READ_FILE = 55
    ERROR_DUPLICATED_EXTERNAL_VARIABLE = 56
    ERROR_INVALID_MODULE_DATA = 57
    ERROR_WRITING_FILE = 58
End Enum

Global abort As Boolean
Private startTime As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, Source As Any, ByVal length As Long)
Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Public hHelperLib As Long
Public hYaraLib As Long

'callback functions
'------------------------------
Public Function vb_stdout(ByVal t As cb_type, ByVal lpMsg As Long) As Long
    Dim msg As String
    Dim m As String
    Dim cm As cb_msgs
    
    If t = cb_update Then
        cm = lpMsg
        If abort Then
            Form1.List1.AddItem "User Aborting..."
            vb_stdout = -1
        End If
        Exit Function
    End If
    
    If lpMsg = 0 Then Exit Function
    
    msg = StringFromPointer(lpMsg)
    
    Select Case t
        Case cb_output: m = "output: " & msg
        Case cb_info: m = "info: " & msg
        Case cb_match: m = "--> MATCH: " & msg
        Case cb_error: m = "error: " & msg
        Case cb_matchInfo: m = msg
        Case Else: m = "unk: " & t & " " & msg
    End Select
    
    If Len(m) > 0 Then Form1.List1.AddItem m
    
    'if we return -1 then scaner callback function will abort
    
End Function



Function FolderExists(path As String) As Boolean
  On Error GoTo hell
  Dim tmp As String
  tmp = path & "\"
  If Len(tmp) = 1 Then Exit Function
  If Dir(tmp, vbDirectory) <> "" Then FolderExists = True
  Exit Function
hell:
    FolderExists = False
End Function

Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function


Function StringFromPointer(buf As Long) As String
    Dim sz As Long
    Dim tmp As String
    Dim b() As Byte
    
    If buf = 0 Then Exit Function
       
    sz = lstrlen(buf)
    If sz = 0 Then Exit Function
    
    ReDim b(sz)
    CopyMemory b(0), ByVal buf, sz
    tmp = StrConv(b, vbUnicode)
    If Right(tmp, 1) = Chr(0) Then tmp = Left(tmp, Len(tmp) - 1)
    
    StringFromPointer = tmp
 
End Function

Sub StartBenchMark(Optional ByRef t As Long = -111)
    If t = -111 Then
        startTime = GetTickCount()
    Else
        t = GetTickCount()
    End If
End Sub

Function EndBenchMark(Optional ByRef t As Long = -111) As String
    Dim endTime As Long, loadTime As Long
    endTime = GetTickCount()
    If t = -111 Then
        loadTime = endTime - startTime
    Else
        loadTime = endTime - t
    End If
    EndBenchMark = loadTime / 1000 & " seconds"
End Function

Function vt_errs2Str(e) As String
   Dim m As String
   Select Case e
       Case 0: m = "ERROR_SUCCESS"
       Case 1: m = "ERROR_INSUFFICIENT_MEMORY"
       Case 2: m = "ERROR_COULD_NOT_ATTACH_TO_PROCESS"
       Case 3: m = "ERROR_COULD_NOT_OPEN_FILE"
       Case 4: m = "ERROR_COULD_NOT_MAP_FILE"
       Case 6: m = "ERROR_INVALID_FILE"
       Case 7: m = "ERROR_CORRUPT_FILE"
       Case 8: m = "ERROR_UNSUPPORTED_FILE_VERSION"
       Case 9: m = "ERROR_INVALID_REGULAR_EXPRESSION"
       Case 10: m = "ERROR_INVALID_HEX_STRING"
       Case 11: m = "ERROR_SYNTAX_ERROR"
       Case 12: m = "ERROR_LOOP_NESTING_LIMIT_EXCEEDED"
       Case 13: m = "ERROR_DUPLICATED_LOOP_IDENTIFIER"
       Case 14: m = "ERROR_DUPLICATED_IDENTIFIER"
       Case 15: m = "ERROR_DUPLICATED_TAG_IDENTIFIER"
       Case 16: m = "ERROR_DUPLICATED_META_IDENTIFIER"
       Case 17: m = "ERROR_DUPLICATED_STRING_IDENTIFIER"
       Case 18: m = "ERROR_UNREFERENCED_STRING"
       Case 19: m = "ERROR_UNDEFINED_STRING"
       Case 20: m = "ERROR_UNDEFINED_IDENTIFIER"
       Case 21: m = "ERROR_MISPLACED_ANONYMOUS_STRING"
       Case 22: m = "ERROR_INCLUDES_CIRCULAR_REFERENCE"
       Case 23: m = "ERROR_INCLUDE_DEPTH_EXCEEDED"
       Case 24: m = "ERROR_WRONG_TYPE"
       Case 25: m = "ERROR_EXEC_STACK_OVERFLOW"
       Case 26: m = "ERROR_SCAN_TIMEOUT"
       Case 27: m = "ERROR_TOO_MANY_SCAN_THREADS"
       Case 28: m = "ERROR_CALLBACK_ERROR"
       Case 29: m = "ERROR_INVALID_ARGUMENT"
       Case 30: m = "ERROR_TOO_MANY_MATCHES"
       Case 31: m = "ERROR_INTERNAL_FATAL_ERROR"
       Case 32: m = "ERROR_NESTED_FOR_OF_LOOP"
       Case 33: m = "ERROR_INVALID_FIELD_NAME"
       Case 34: m = "ERROR_UNKNOWN_MODULE"
       Case 35: m = "ERROR_NOT_A_STRUCTURE"
       Case 36: m = "ERROR_NOT_INDEXABLE"
       Case 37: m = "ERROR_NOT_A_FUNCTION"
       Case 38: m = "ERROR_INVALID_FORMAT"
       Case 39: m = "ERROR_TOO_MANY_ARGUMENTS"
       Case 40: m = "ERROR_WRONG_ARGUMENTS"
       Case 41: m = "ERROR_WRONG_RETURN_TYPE"
       Case 42: m = "ERROR_DUPLICATED_STRUCTURE_MEMBER"
       Case 43: m = "ERROR_EMPTY_STRING"
       Case 44: m = "ERROR_DIVISION_BY_ZERO"
       Case 45: m = "ERROR_REGULAR_EXPRESSION_TOO_LARGE"
       Case 46: m = "ERROR_TOO_MANY_RE_FIBERS"
       Case 47: m = "ERROR_COULD_NOT_READ_PROCESS_MEMORY"
       Case 48: m = "ERROR_INVALID_EXTERNAL_VARIABLE_TYPE"
       Case 49: m = "ERROR_REGULAR_EXPRESSION_TOO_COMPLEX"
       Case 50: m = "ERROR_INVALID_MODULE_NAME"
       Case 51: m = "ERROR_TOO_MANY_STRINGS"
       Case 52: m = "ERROR_INTEGER_OVERFLOW"
       Case 53: m = "ERROR_CALLBACK_REQUIRED"
       Case 54: m = "ERROR_INVALID_OPERAND"
       Case 55: m = "ERROR_COULD_NOT_READ_FILE"
       Case 56: m = "ERROR_DUPLICATED_EXTERNAL_VARIABLE"
       Case 57: m = "ERROR_INVALID_MODULE_DATA"
       Case 58: m = "ERROR_WRITING_FILE"
   End Select
   vt_errs2Str = m
End Function
