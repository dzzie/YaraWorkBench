VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMatchResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public offset As Long
Public leng As Long
Public id As String

Function dump() As String
    Dim x() As String
    push x, Rpad("Var:") & id
    push x, Rpad("Leng:") & leng
    push x, Rpad("FOff:") & Hex(offset)
    dump = Join(x, vbCrLf)
End Function

Friend Function loadSelf(Msg)
    Dim tmp() As String
    tmp = Split(Replace(Msg, vbTab, ""), ":")
    offset = CLng(Replace(tmp(0), "0x", "&h"))
    leng = CLng("&h" & tmp(1))
    id = tmp(2)
End Function

