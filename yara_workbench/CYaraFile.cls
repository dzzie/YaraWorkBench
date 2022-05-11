VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CYaraFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public found As Boolean
Public matches As New Collection 'of CYaraMatch
Public File As String
Public ElapsedTime As String
Public isMemStr As Boolean
Public parent As CYaraScan

Private di As CStringBuilder

Property Get DumpInfo() As String
    If di Is Nothing Then Exit Property
    DumpInfo = di.ToString()
End Property

'we load the class on demand because only used for frmModuleDump and has heap allocs in class initilize
Public Function AddDumpInfo(info As String)

    If di Is Nothing Then
        Set di = New CStringBuilder
    End If
    
    di.Append info
    
End Function


Function LastMatch() As CYaraMatch
     Set LastMatch = ActiveYaraFile.matches(ActiveYaraFile.matches.count)
End Function

Function TotalMatches() As Long
    Dim m As CYaraMatch
    For Each m In matches
        TotalMatches = TotalMatches + m.results.count
    Next
End Function

Function MatchNames(Optional divider As String = ", ")
    Dim m As CYaraMatch, tmp As String, t() As String
    If matches.count = 0 Then Exit Function
    For Each m In matches
        tmp = m.Name
        If Left(tmp, 6) = "match " Then tmp = Mid(tmp, 7)
        If Left(tmp, 8) = "default." Then tmp = Mid(tmp, 9)
        push t, tmp
    Next
    MatchNames = Join(t, divider)
End Function
