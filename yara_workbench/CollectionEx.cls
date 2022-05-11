VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollectionEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Wrapper for VB6 collection object with some extended capabilities
' add now allows keys to be numeric only and is case sensitive
' if you use a numeric only key, make sure to use forceAsKey =1 when accessing by key
' the CCollectionExtender class allows you to do some more handy stuff not traditionally possible
'  author: elroy
'  link:   http://www.vbforums.com/showthread.php?834579-Wrapper-for-VB6-Collections

Option Explicit

Dim c As New Collection

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)

Private Type dump
    key As String
    data As Variant
End Type

Public OptionForceAsKey As Boolean

Public Sub Attach(existing As Collection)
    Set c = existing
End Sub

Public Sub AddAutoIndex(vData As Variant, Optional ByVal key As String, Optional Before As Variant, Optional After As Variant)
    
    Dim i As Long, temp As String

    If Len(key) > 0 Then
        If keyExists(key) Then
            Do
                i = i + 1
                temp = key & "_" & i
            Loop While keyExists(temp)
            key = temp
        End If
    End If
    
    Add vData, key, Before, After
    
End Sub

Public Sub Add(vData As Variant, Optional ByVal key As String, Optional Before As Variant, Optional After As Variant)
    ' If not specified, Before and/or After come is as IsError() = True, which the collection.Add method can handle.
    ' It still raises an error if both are specified, and that behavior is not changed.
    ' Also errors when trying to add duplicate keys, as expected.
    If Len(key) Then
        c.Add vData, Base16Encode(key), Before, After
    Else
        c.Add vData, , Before, After
    End If
End Sub

Property Let Item(ByVal IndexOrKey As Variant, Optional ByVal forceAsKey As Byte = 0, val As Variant)

'this works as default as well since set in get
    Dim i As Integer
    Dim k As String
    Dim tmp As Variant
    
    i = VarType(IndexOrKey)
    
    If OptionForceAsKey Then forceAsKey = 1
    If forceAsKey <> 0 Then i = 8
    
    Select Case i
        Case 2 To 7 ' Integer, Long, Single, Double, Currency, & Date.  (A bit weird to pass in a date, but who knows.)
            
            k = Me.keyForIndex(CLng(IndexOrKey))
            c.remove CLng(IndexOrKey)
            If Len(k) > 0 Then
                c.Add val, k
            Else
                c.Add val
            End If
            
        Case 8
        
            If keyExists(CStr(IndexOrKey)) Then
                k = Base16Encode(CStr(IndexOrKey))
                Me.remove IndexOrKey, forceAsKey
            End If
            
            Add val, IndexOrKey
            
    End Select
    

End Property

Property Get Item(ByVal IndexOrKey As Variant, Optional ByVal forceAsKey As Byte = 0) As Variant
Attribute Item.VB_UserMemId = 0
    'Attribute Item.VB_UserMemId = 0 ' This must be placed in with Notepad.  Makes this method the default.
    ' Had to leave this one public so the "Default" method could be used.
    Dim i As Integer
    Dim k As String
    
    i = VarType(IndexOrKey)
    
    If OptionForceAsKey Then forceAsKey = 1
    If forceAsKey <> 0 Then i = 8
    
    Select Case i
        Case 2 To 7 ' Integer, Long, Single, Double, Currency, & Date.  (A bit weird to pass in a date, but who knows.)
       
            If IsObject(c.Item(IndexOrKey)) Then
                Set Item = c.Item(IndexOrKey)
            Else
                Item = c.Item(IndexOrKey)
            End If
    
        Case 8
            k = Base16Encode(CStr(IndexOrKey))
            
            If IsObject(c.Item(k)) Then
                Set Item = c.Item(k)
            Else
                Item = c.Item(k)
            End If
    
    End Select
    
End Property

'property so we can mouse over it in ide to see value..
Property Get count()
    count = c.count
End Property

Public Sub remove(ByVal IndexOrKey As Variant, Optional ByVal forceAsKey As Byte = 0)
    Dim i As Integer
    
    i = VarType(IndexOrKey)
    
    If OptionForceAsKey Then forceAsKey = 1
    If forceAsKey <> 0 Then i = 8
    
    Select Case i
        Case 2 To 7 ' Integer, Long, Single, Double, Currency, & Date.  (A bit weird to pass in a date, but who knows.)
            c.remove CLng(IndexOrKey)
        Case 8
            c.remove Base16Encode(CStr(IndexOrKey))
    End Select
    
    ' It does nothing if bad IndexOrKey type (such as Null, an object or other nonsense).
End Sub

Sub CreateKey(ByVal k As String)
    If Me.keyExists(k) Then Exit Sub
    Add Empty, k
End Sub

Function dumpKeys(Optional divider As String = vbCrLf) As String
    Dim i As Long
    Dim tmp() As String
    For i = 1 To c.count
        push tmp, Me.keyForIndex(i)
    Next
    dumpKeys = Join(tmp, divider)
End Function

Property Get isEmpty() As Boolean
    isEmpty = (c.count = 0)
End Property

Public Sub Clear()
    Set c = New Collection
End Sub

Public Function Clone() As CollectionEx
    
    Dim cc As New CollectionEx
    Dim i As Long
    
    For i = 1 To c.count
        'If IsObject(c(i)) Then Err.Raise 256, , "CollectionEx.Clone: Can not clone a collection with object references"
        cc.Add c(i), keyForIndex(i)
    Next
    
    Set Clone = cc
End Function

Public Function uniqueKey(ByVal suggested As String) As String
    Dim i As Long, tmp As String
    
    If Len(suggested) = 0 Then suggested = Base16Encode(Now)
    
    If keyExists(suggested) Then
        Do
            i = i + 1
            tmp = suggested & "_" & i
        Loop While keyExists(tmp)
        suggested = tmp
    End If
    
    uniqueKey = suggested
    
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'Attribute NewEnum.VB_UserMemId = -4 ' This must be placed in with Notepad.
    'Attribute NewEnum.VB_MemberFlags = "40" ' This is also a good idea, as it hides this method.
    '
    ' This allows use of the "For Each" syntax.
    ' Just enumerate with a variant for the item, using this class as the series.
    '
   Set NewEnum = c.[_NewEnum]
End Function

Public Function keyExists(ByVal key As String) As Boolean
    On Error GoTo DoesntExist
    IsObject c(Base16Encode(key))
    keyExists = True
DoesntExist:
End Function

Public Function toString(Optional ByVal delimiter As String = vbCrLf, Optional ByVal includeKeys As Boolean = False) As String
    Dim tmp() As Variant
    Dim k, i As Long
    
    If c.count = 0 Then GoTo done
    
    ReDim tmp(c.count - 1)
    
    For i = 1 To c.count
        k = Empty
        If includeKeys Then k = keyForIndex(i)
        If Len(k) > 0 Then k = k & " = "
        If IsObject(c(i)) Then
            tmp(i - 1) = k & "[ " & TypeName(c(i)) & " object ]"
        Else
            tmp(i - 1) = k & c(i)
        End If
    Next
    
done:
    toString = Join(tmp, delimiter)
End Function

Public Function toDumpFile(ByVal fPath As String) As Boolean
        
    On Error GoTo hell
    
    If FileExists(fPath) Then Kill fPath
    If FileExists(fPath) Then GoTo hell
    
    Dim tmp() As dump, d As dump
    Dim v, i As Long, f As Long
    
    If c.count = 0 Then GoTo hell
    
    ReDim tmp(c.count - 1)
    
    For i = 1 To c.count
        tmp(i - 1).key = keyForIndex(i)
        If IsObject(c(i)) Then
            tmp(i - 1).data = "[ invalid can not save " & TypeName(c(i)) & " object ]"
        Else
            tmp(i - 1).data = c(i)
        End If
    Next
    
    f = FreeFile
    Open fPath For Binary As f
    Put f, , CLng(UBound(tmp))
    Put f, , tmp()
    Close f
    
    toDumpFile = True
    
hell:
End Function

Public Function toTextFile(ByVal fPath As String, Optional header As String) As Boolean
        
    On Error GoTo hell
    
    If FileExists(fPath) Then Kill fPath
    If FileExists(fPath) Then GoTo hell
    
    Dim tmp() As String
    Dim v, i As Long, f As Long
    
    If c.count = 0 Then GoTo hell
    
    ReDim tmp(c.count - 1)
    
    For i = 1 To c.count
        If IsObject(c(i)) Then
            tmp(i - 1) = "[ invalid can not save " & TypeName(c(i)) & " object ]"
        Else
            tmp(i - 1) = Replace(c(i), vbCrLf, "\n")
        End If
        If i Mod 100 = 0 Then DoEvents
    Next
    
    If Len(header) > 0 Then
        header = vbCrLf & header & vbCrLf & String(100, "-") & vbCrLf & vbCrLf
    End If
    
    WriteFile fPath, header & Join(tmp, vbCrLf)
    
    toTextFile = True
    
hell:
End Function

Public Function fromTextFile(ByVal fPath As String, Optional append As Boolean, Optional trim As Boolean) As Long
    
    On Error GoTo hell
    
    If Not FileExists(fPath) Then GoTo hell
    
    Dim tmp() As String, dat As String
    Dim v, i As Long, f As Long, cnt As Long
    
    If c.count > 0 And append = False Then Set c = New Collection
    
    dat = ReadFile(fPath)
    If InStr(dat, Chr(0)) > 0 Then
        MsgBox "This looks like binary data and not a text file?"
        Exit Function
    End If
    
    tmp() = Split(dat, vbCrLf)
    
    For i = 0 To UBound(tmp)
        If trim Then tmp(i) = VBA.trim(tmp(i))
        c.Add tmp(i)
        If i Mod 100 = 0 Then DoEvents
    Next

hell:
    fromTextFile = c.count
End Function

Public Function fromDumpFile(ByVal fPath As String, Optional append As Boolean, Optional trimIt As Boolean) As Long
    
    On Error GoTo hell
    
    If Not FileExists(fPath) Then GoTo hell
    
    Dim tmp() As dump ', d As dump
    Dim v, i As Long, f As Long, cnt As Long
    Dim d As String
    
    If c.count > 0 And append = False Then Set c = New Collection
    
    f = FreeFile
    Open fPath For Binary As f
    Get f, , cnt
    ReDim tmp(cnt)
    Get f, , tmp()
    Close f
    
    For i = 0 To UBound(tmp)
        d = tmp(i).data
        If trimIt Then d = trim(d)
        If Len(tmp(i).key) = 0 Then
            c.Add d
        Else
            c.Add d, Base16Encode(uniqueKey(tmp(i).key))
        End If
    Next

hell:
    fromDumpFile = c.count
End Function

Public Function keyForValue(val, ByRef key) As Boolean
    Dim i As Long
    
    For i = 1 To c.count
            
        If IsObject(c(i)) Then
            If IsObject(val) Then
                If ObjPtr(c(i)) = ObjPtr(val) Then
                    key = keyForIndex(i)
                    keyForValue = True
                    Exit Function
                End If
            End If
        Else
            If c(i) = val Then
                key = keyForIndex(i)
                keyForValue = True
                Exit Function
            End If
        End If
            
    Next
    
End Function

Public Function keyForIndex(index As Long) As String
    ' Get a key based on its index value.  Must be in range, or error.
    Dim i     As Long
    Dim Ptr   As Long
    Dim sKey  As String
    '
    If index < 1 Or index > c.count Then
        Err.Raise 9
        Exit Function
    End If
    '
    If index <= c.count / 2 Then                                ' Start from front.
        CopyMemory Ptr, ByVal ObjPtr(c) + &H18, 4               ' First item pointer of collection header.
        For i = 2 To index
            CopyMemory Ptr, ByVal Ptr + &H18, 4                 ' Next item pointer of collection item.
        Next i
    Else                                                        ' Start from end and go back.
        CopyMemory Ptr, ByVal ObjPtr(c) + &H1C, 4               ' Last item pointer of collection header.
        For i = c.count - 1 To index Step -1
            CopyMemory Ptr, ByVal Ptr + &H14, 4                 ' Previous item pointer of collection item.
        Next i
    End If
    '
    i = StrPtr(sKey)                                            ' Save string pointer because we're going to borrow the string.
    CopyMemory ByVal VarPtr(sKey), ByVal Ptr + &H10, 4          ' Key string of collection item.
    keyForIndex = Base16Decode(sKey)                                ' Move key into property's return.
    CopyMemory ByVal VarPtr(sKey), i, 4                         ' Put string pointer back to keep memory straight.
End Function

Public Function indexForKey(ByVal key As String) As Long
    Dim i As Long
    indexForKey = 0
    For i = 1 To c.count
        If keyForIndex(i) = key Then
            indexForKey = i
            Exit Function
        End If
    Next
End Function

Public Function changeKey(ByVal OldKey As String, ByVal NewKey As String) As Boolean
    Dim Ptr   As Long
    Dim tKey  As String
    Dim iHold  As Long
    Dim index As Long
    Dim vData As Variant
    
    On Error GoTo hell
    
    index = indexForKey(OldKey)
   
    If IsObject(c.Item(index)) Then
        Set vData = c.Item(index)
    Else
       vData = c.Item(index)
    End If
    
    c.remove index
    
    If c.count = 0 Then
        c.Add vData, Base16Encode(NewKey)             ' Add first item.
    ElseIf index > c.count Then
        c.Add vData, Base16Encode(NewKey), , c.count  ' Add to end.
    Else
        c.Add vData, Base16Encode(NewKey), index      ' Add prior to one specified, so it takes on that index number.
    End If
    
    changeKey = True
    Exit Function
hell:
End Function

Public Function changeKeyByIndex(ByVal index As Long, ByVal NewKey As String) As Boolean
    Dim curKey As String
    On Error GoTo hell
    curKey = keyForIndex(index)
    changeKeyByIndex = changeKey(curKey, NewKey)
    Exit Function
hell:
End Function

Public Function changeIndex(OldIndex As Long, newIndex As Long) As Boolean
    Dim vData As Variant
    Dim sKey As String
    Dim i As Long
    
    On Error GoTo hell
    
    If OldIndex = newIndex Then GoTo ok
    If c.count = 1 Then GoTo ok
    
    If IsObject(c.Item(OldIndex)) Then
        Set vData = c.Item(OldIndex)
    Else
        vData = c.Item(OldIndex)
    End If

    sKey = Base16Encode(keyForIndex(OldIndex))
    c.remove OldIndex
    
    If c.count = 0 Then
        c.Add vData, sKey             ' Add first item.
    ElseIf newIndex > c.count Then
        c.Add vData, sKey, , c.count  ' Add to end.
    Else
        c.Add vData, sKey, newIndex   ' Add prior to one specified, so it takes on that index number.
    End If
    
ok:
    changeIndex = True
    Exit Function
hell:
End Function

Public Function changeIndexByKey(ByVal key As String, ByVal newIndex As Long) As Boolean
    Dim curIndex As Long
    On Error GoTo hell
    curIndex = indexForKey(key)
    changeIndexByKey = changeIndex(curIndex, newIndex)
    Exit Function
hell:
End Function

Public Function Keys() As String()
    Dim tmp() As String, i As Long
    
    If c.count = 0 Then Exit Function
    
    push tmp(), ""
    For i = 1 To c.count
        push tmp(), keyForIndex(i)
    Next
    
    Keys = tmp()
    
End Function

Public Function toArray() As Variant()
    Dim tmp() As Variant
    Dim v, i As Long
    Dim o As Object
    
    If c.count = 0 Then GoTo done
    
    ReDim tmp(c.count - 1)
    
    For i = 1 To c.count
        If IsObject(c(i)) Then
            Set tmp(i - 1) = c(i)
        Else
            tmp(i - 1) = c(i)
        End If
    Next
    
done:
    toArray = tmp()
End Function

Public Sub fromArray(ary As Variant, Optional append As Boolean = True, Optional uniqueOnly As Boolean, Optional trim As Boolean)
    Dim i As Long
    'On Error GoTo hell
    
    If AryIsEmpty(ary) Then Exit Sub
    If Not append And c.count > 0 Then Set c = New Collection
    
    For i = LBound(ary) To UBound(ary)
        If trim Then ary(i) = VBA.trim(ary(i))
        If uniqueOnly Then
            If Not keyExists(ary(i)) Then
                Add ary(i), ary(i)
            End If
        Else
            c.Add ary(i)
        End If
    Next

Exit Sub
hell:
    Stop
End Sub


Public Function findMatches(c2 As CollectionEx, Optional byKey As Boolean = False) As CollectionEx
    
    Dim ret As New CollectionEx
    Dim i As Long
    Dim k As String
    Dim v, v2
    Dim longestCol As CollectionEx
    Dim shortCol As CollectionEx
    Dim a As String, b As String
    
    If c.count > c2.count Then
        Set longestCol = Me
        Set shortCol = c2
    Else
        Set longestCol = c2
        Set shortCol = Me
    End If
    
    On Error Resume Next
    Set findMatches = ret
    
    If byKey Then
        For i = 1 To longestCol.count
            k = longestCol.keyForIndex(i)
            If shortCol.keyExists(k) Then
                ret.Add longestCol(i), k
            End If
        Next
    Else
        For i = 1 To longestCol.count
            v = longestCol(i)
            v = trim(v)
            For Each v2 In shortCol
               If LCase(v) = LCase(v2) Then
                   ret.Add v, longestCol.keyForIndex(i)
                   Exit For
               End If
            Next
        Next
    End If
        
End Function

Public Function diff(c2 As CollectionEx, Optional byKey As Boolean = False) As CollectionEx
    
    Dim ret As New CollectionEx
    Dim i As Long
    Dim k As String
    Dim v, v2
    Dim found As Boolean
    Dim longestCol As CollectionEx
    Dim shortCol As CollectionEx
    
    On Error Resume Next
    Set diff = ret
    
    If c.count > c2.count Then
        Set longestCol = Me
        Set shortCol = c2
    Else
        Set longestCol = c2
        Set shortCol = Me
    End If
    
    If byKey Then
        For i = 1 To longestCol.count
            k = longestCol.keyForIndex(i)
            If Not shortCol.keyExists(k) Then
                ret.Add longestCol(i), k
            End If
            If i Mod 100 = 0 Then DoEvents
        Next
    Else
        For i = 1 To longestCol.count
            v = longestCol(i)
            found = False
            For Each v2 In shortCol
                If v = v2 Then
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                ret.Add v, longestCol.keyForIndex(i)
            End If
            If i Mod 100 = 0 Then DoEvents
        Next
    End If
        
End Function




Private Function Base16Encode(s As String) As String
    Dim i As Long
    Base16Encode = Space$(Len(s) * 4)
    For i = 0 To Len(s) - 1
        Mid$(Base16Encode, i * 4 + 1, 4) = Right$("0000" & Hex$(AscW(Mid$(s, i + 1, 1))), 4)
    Next i
End Function

Private Function Base16Decode(s As String) As String
    Dim i As Long
    Base16Decode = Space$(Len(s) \ 4)
    For i = 0 To Len(s) - 1 Step 4
        Mid$(Base16Decode, i \ 4 + 1, 1) = ChrW$(val("&h" & Mid$(s, i + 1, 4)))
    Next i
End Function

Private Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Function Sort(Optional descending As Boolean = True) As Boolean
    
    Dim tmp() As Variant
    Dim Keys() As String
    
    Dim v, i As Long
    Dim o As Object
    
    If c.count = 0 Then Exit Function
    
    ReDim tmp(c.count - 1)
    ReDim Keys(c.count - 1)
    
    For i = 1 To c.count
        If IsObject(c(i)) Then
            Exit Function
        Else
            Keys(i - 1) = Me.keyForIndex(i)
            tmp(i - 1) = c(i)
        End If
    Next
    
    If Not SortArray(descending, tmp, Keys) Then Exit Function
    
    Set c = New Collection
    
    For i = 0 To UBound(tmp)
        Add tmp(i), Keys(i)
    Next
    
    Sort = True

End Function

Private Function SortArray(descending As Boolean, ByRef TheArray As Variant, ByRef Keys As Variant) As Boolean
    Dim sorted As Boolean
    Dim temp, temp_key, x As Long
    Dim doit As Boolean
    
    On Error GoTo hell
    
    sorted = False
    Do While Not sorted
        sorted = True
        For x = 0 To UBound(TheArray) - 1
        
            doit = False
            
            If descending Then
                If TheArray(x) < TheArray(x + 1) Then doit = True
            Else
                If TheArray(x) > TheArray(x + 1) Then doit = True
            End If
            
            If doit Then
                temp = TheArray(x + 1)
                temp_key = Keys(x + 1)
                TheArray(x + 1) = TheArray(x)
                Keys(x + 1) = Keys(x)
                TheArray(x) = temp
                Keys(x) = temp_key
                sorted = False
            End If
            
        Next x
    Loop
    
    SortArray = True
hell:
    
End Function

Private Function ReadFile(filename) As Variant
  Dim f As Long
  Dim temp As Variant
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Private Function WriteFile(ByVal path As String, it) As Boolean
    Dim f As Long
    Dim b() As Byte
    On Error Resume Next
    
    If FileExists(path) Then Kill path
    
    If TypeName(it) = "Byte()" Then
        b() = it
    Else
        b() = StrConv(CStr(it), vbFromUnicode, &H409)
    End If
    
    f = FreeFile
    Open path For Binary As #f
    Put #f, , it
    Close f
    
    WriteFile = (Err.Number = 0)
    
End Function
