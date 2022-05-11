VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIntellisenseItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ObjName As String
Public keywords As String
Public protoTypes As New CollectionEx  'any function call with args or ()
Private comments() As String           'parallel to protos by index
Public keyWordCount As Long
Public path As String
Public enums As New CollectionEx
Public subItems As New CollectionEx

Friend Function loadSelf(ByVal pth As String) As Long
    Dim tmp() As String, x, alias() As String, cmt
    Dim a As Long, keyWord As String, hasProto As Boolean
    Dim e As CEnumList, s As CSubItem, s2 As CSubItem, lineNo As Long
    Dim lastTopLevelKeyword As String, lastSecondLevelKeyword As String
    
    If Not fso.FileExists(pth) Then Exit Function
    
    path = pth
    ObjName = fso.GetBaseName(pth)
    tmp() = Split(fso.ReadFile(pth), vbCrLf)
    
    For Each x In tmp
    
        lineNo = lineNo + 1
        'Debug.Print "line: " & lineNo & " raw: " & x
        
        hasProto = False
        cmt = Empty
        x = trim(Replace(x, vbTab, Empty))
        
        If Left(x, 1) = "{" Then
            If s Is Nothing Then
                Set s = New CSubItem
                s.owner = Me.ObjName & "." & lastTopLevelKeyword & "."
                GoTo nextOne
            Else
                Set s2 = New CSubItem
                s2.owner = s.owner & lastSecondLevelKeyword & "."
                GoTo nextOne
            End If
        End If
        
        If Left(x, 1) = "}" Then
            If Not s2 Is Nothing Then
                subItems.Add s2, s2.owner
                Set s2 = Nothing
            Else
                If Not s Is Nothing Then
                    subItems.Add s, s.owner
                    Set s = Nothing
                End If
            End If
            GoTo nextOne
        End If
        
        If Left(x, 5) = "#enum" Then
            If Not e Is Nothing Then enums.Add e, e.owner
            Set e = New CEnumList
            e.owner = trim(Mid(x, 6)) '#enum pe.machine assumes all enums at end of modules one after the other..
        End If
        
        If Left(x, 1) = "#" Then GoTo nextOne
        If Len(x) = 0 Then GoTo nextOne
        
        a = InStr(x, ";")
        
        If a > 0 Then
            cmt = Mid(x, a + 1)
            x = trim(Mid(x, 1, a - 1))
        End If
        
        a = InStr(x, "(")
        
        If a > 0 And InStr(x, "()") < a Then
            hasProto = True
            keyWord = trim(Mid(x, 1, a - 1))
            If protoTypes.keyExists(keyWord) Then
                protoTypes(keyWord) = protoTypes(keyWord) & vbCrLf & x 'were going multiline with overloads
                GoTo nextOne
            Else
                protoTypes.Add x, keyWord
                push comments, cmt 'only comments first func if overloaded
            End If
            x = keyWord
        End If
        
        
        If Not s2 Is Nothing Then
            s2.keywords = s2.keywords & " " & x
        ElseIf Not s Is Nothing Then
            lastSecondLevelKeyword = x
            s.keywords = s.keywords & " " & x
        Else
        
            If Not hasProto And protoTypes.keyExists(x) Then GoTo nextOne
            
            'enums get added to both the specific enum and the master list
            If Not e Is Nothing Then e.keywords = e.keywords & " " & x
    
            lastTopLevelKeyword = x
            keywords = keywords & " " & x
            keyWordCount = keyWordCount + 1
        End If
        
nextOne:
    Next
           
    If Not e Is Nothing Then
        enums.Add e, e.owner
    End If
    
    While InStr(keywords, "  ") > 0
        keywords = Replace(keywords, "  ", " ")
    Wend
    
End Function

Function protoTypeFor(keyWord) As String
    Dim x, tmp() As String, Y, a As Long
    For Each x In protoTypes
        a = InStr(x, "(")
        If a > 0 Then x = Mid(x, 1, a - 1)
        If LCase(x) = LCase(keyWord) Then
            protoTypeFor = trim(x)
            Exit Function
        End If
    Next
End Function
'
'Function funcRetType(keyWord) As String
'    Dim proto As String, a As Long, b As Long, wasAry As Boolean, tmp As String
'    proto = protoTypeFor(keyWord)
'    If Len(proto) = 0 Then Exit Function
'
'    If Right(proto, 2) = "()" Then
'        proto = Mid(proto, 1, Len(proto) - 2) 'if as string(), we string the ()
'        wasAry = True 'or func call with no args..but then there will be no " As "
'    End If
'
'    a = InStrRev(proto, " As ", , vbTextCompare)
'    b = InStrRev(proto, ")")
'    If a > 0 And a > b Then 'not a function argument
'        tmp = Mid(proto, a + 4)
'        If Left(tmp, 4) = "New " Then tmp = Mid(tmp, 5)
'        funcRetType = tmp
'    End If
'
'End Function
'
'Function hasExactKeyWord(keyWord) As Boolean
'    Dim x
'    For Each x In Split(keyWords, " ")
'        If x = keyWord Then
'            hasExactKeyWord = True
'            Exit Function
'        End If
'    Next
'End Function


