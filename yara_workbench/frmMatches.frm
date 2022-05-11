VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMatches 
   Caption         =   "Matches"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   12315
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   11820
      Top             =   180
   End
   Begin MSComctlLib.ListView lv 
      Height          =   3915
      Left            =   60
      TabIndex        =   1
      Top             =   4380
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   6906
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "L"
         Object.Width           =   617
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "foff"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Var"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   3060
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   9075
   End
   Begin MSComctlLib.ListView lvRule 
      Height          =   4155
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   7329
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Rule"
         Object.Width           =   7056
      EndProperty
   End
End
Attribute VB_Name = "frmMatches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim my_file As CYaraFile
Dim pe As New CPEEditor
Dim cap As CapstoneAx.CDisassembler

Private Function initCapstone(mode As cs_mode, ByRef errMsg As String) As Boolean
    Dim x() As String
    
    errMsg = Empty
    If Not cap Is Nothing Then Set cap = Nothing
    
    If hCapstone = 0 Then
        push x, "Could not find capstone.dll"
        GoTo exitNow
    End If
    
    Set cap = New CDisassembler
    
    If Not cap.init(CS_ARCH_X86, mode, True) Then
        push x, "Failed to init engine: " & cap.errMsg
        GoTo exitNow
    End If
      
    push x, "Capstone loaded @ 0x" & Hex(cap.hLib)
    push x, "hEngine: 0x" & Hex(cap.hCapstone)
    push x, "Version: " & cap.version
    initCapstone = True
    
exitNow:
    errMsg = Join(x, vbCrLf)
    
End Function

Function LoadRules(f As CYaraFile)
    
    On Error Resume Next
    
    Dim x() As String
    Dim li As ListItem
    Dim m As CYaraMatch
    Dim tmp As String
    
    text1 = Empty
    lvRule.ListItems.Clear
    lv.ListItems.Clear

    Set my_file = f
    
    If Not f.isMemStr Then
        If pe.LoadFile(f.File) Then
            text1 = "PE File Loaded ok: " & vbCrLf & f.File
        Else
            text1 = "Failed to load pe file: " & pe.errMessage & vbCrLf & f.File
        End If
    End If
    
    For Each m In f.matches
        tmp = m.name
        If Left(tmp, 6) = "match " Then tmp = Mid(tmp, 7)
        If Left(tmp, 8) = "default." Then tmp = Mid(tmp, 9)
        Set li = lvRule.ListItems.Add(, , tmp)
        Set li.Tag = m
    Next
    
    If lvRule.ListItems.count > 0 Then
        lvRule_ItemClick lvRule.ListItems(1)
        If lv.ListItems.count > 0 Then
            lv_ItemClick lv.ListItems(1)
        End If
    End If
    
    Me.Visible = True
    TopMost Me
    Timer1.Enabled = True
    
End Function

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    FormPos Me, True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    text1.Width = Me.Width - text1.Left - 240
    text1.Height = Me.Height - text1.Top - 520
    lv.Height = Me.Height - lv.Top - 520
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormPos Me, True, True
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)

    On Error GoTo hell
    
    Dim r As CMatchResult
    Dim ul As ULong64
    Dim sect As String
    Dim x() As String
    Dim errMsg As String
    Dim mode As cs_mode
    Dim ret As Collection
    Dim ci As CapstoneAx.CInstruction
    
    Dim buf() As Byte
    Dim hDump As String
    Dim f As Long
    
    Set r = Item.Tag
    
    If fso.FileExists(my_file.File) Then
        f = FreeFile
        ReDim buf((16 * 4) - 1) 'line lines of hexdump
        Open my_file.File For Binary Access Read As f
        Get f, r.offset + 1, buf()
        Close f
        hDump = HexDump(buf, r.offset)
    End If
    
    push x, Rpad("File:") & my_file.File
    push x, Rpad("Scan:") & my_file.ElapsedTime
    push x, Rpad("Rule:") & lvRule.SelectedItem.Text
    
    push x, r.dump
    
    If pe.isLoaded Then
        Set ul = pe.OffsetToVA(r.offset, sect)
        ul.use0x = True
        ul.useTick = True
        push x, Rpad("VA:") & ul.ToString()
        push x, Rpad("Date:") & pe.CompiledDate
        push x, Rpad(".Net:") & pe.isDotNet
        push x, Rpad("Bit:") & IIf(pe.is32Bit, "32", "64") & vbCrLf
    End If
    
    push x, hDump & vbCrLf
    
    If pe.isLoaded Then
    
        If Not AryIsEmpty(buf) Then
            If pe.is32Bit Then mode = CS_MODE_32 Else mode = CS_MODE_64
            If initCapstone(mode, errMsg) Then
                Set ret = cap.disasm(&H1000, buf)
                For Each ci In ret
                    push x, Mid(ci.Text, Len("00001000    ") + 1)
                Next
            End If
        End If
            
    End If
           
    text1 = Join(x, vbCrLf)
    
    Exit Sub
hell: MsgBox Err.Description
    
End Sub

Private Sub lvRule_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim m As CYaraMatch, r As CMatchResult, li As ListItem
    lv.ListItems.Clear
    Set m = Item.Tag
    For Each r In m.results
        Set li = lv.ListItems.Add(, , r.leng)
        li.subItems(1) = Hex(r.offset)
        li.subItems(2) = r.id
        Set li.Tag = r
    Next
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    TopMost Me, False
End Sub
