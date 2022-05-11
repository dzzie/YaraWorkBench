VERSION 5.00
Object = "{2668C1EA-1D34-42E2-B89F-6B92F3FF627B}#5.0#0"; "scivb2.ocx"
Begin VB.Form frmNavigator 
   Caption         =   "Signature Navigator / Extractor"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   13080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCopyList 
      Caption         =   "keep running copy list"
      Height          =   195
      Left            =   5100
      TabIndex        =   5
      Top             =   120
      Value           =   1  'Checked
      Width           =   1995
   End
   Begin YaraWorkBench.ucFilterList lv 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   12726
   End
   Begin sci2.SciSimple sci 
      Height          =   7215
      Left            =   3900
      TabIndex        =   1
      Top             =   360
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   12726
   End
   Begin VB.Label mnuClearClip 
      Caption         =   "Clear Clipboard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3780
      TabIndex        =   4
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "(Dbl Click to add items to clipboard)"
      Height          =   195
      Left            =   900
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblRefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuCopySelected 
         Caption         =   "Copy Selected"
      End
      Begin VB.Menu mnuFindDups 
         Caption         =   "Find Duplicates"
      End
   End
End
Attribute VB_Name = "frmNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim selItem As ListItem
Dim itemsCopied As Long

Private Sub Form_Load()
    
    On Error Resume Next
    
    itemsCopied = 0
    Clipboard.Clear
    lv.Clear
    lv.SetColumnHeaders "Start,Lines,Name*", "10,10,7300"
    lv.SetFont "Courier", 12
    lv.MultiSelect = True
    lv.AllowDelete = True
    mnuPopup.visible = False
    sci.WordWrap = False
    SyntaxColor sci
    Me.Icon = frmMain.Icon
    
    Dim tmp() As String, x, y
    Dim curRule() As String, curName As String, curStart As Long
    Dim li As ListItem
    Dim i As Long
    
    tmp = Split(frmMain.sci.Text, vbCrLf)
    For Each x In tmp
        i = i + 1
        y = LCase(trim(Replace(x, vbTab, Empty)))
        If Left(y, 6) = "import" Then
            DoEvents
        ElseIf Left(y, 5) = "rule " Or Left(y, 12) = "private rule" Then
            If Not AryIsEmpty(curRule) Then
                Set li = lv.AddItem(curStart, UBound(curRule) + 1, curName)
                li.Tag = Join(curRule, vbCrLf)
            End If
            curStart = i
            Erase curRule
            push curRule, x
            If Left(y, 12) = "private rule" Then
                curName = trim(Mid(x, InStr(1, x, "private rule", vbTextCompare) + 12))
            Else
                curName = trim(Mid(x, InStr(1, x, "rule", vbTextCompare) + 4))
            End If
            curName = Replace(curName, "{", Empty)
        Else
            If Len(curName) > 0 Then push curRule, x
        End If
    Next
            
    If Not AryIsEmpty(curRule) Then
        Set li = lv.AddItem(curStart, UBound(curRule) + 1, curName)
        li.Tag = Join(curRule, vbCrLf)
    End If
    
    Me.Caption = lv.ListItems.count & " Rules loaded"
    If Len(frmMain.curFile) > 0 Then Me.Caption = Me.Caption & "  - File: " & frmMain.curFile
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    sci.Width = Me.Width - sci.Left - 400
    sci.Height = Me.Height - sci.Top - 600
    lv.Height = sci.Height
End Sub

Private Sub lblRefresh_Click()
    Form_Load
End Sub

Private Sub lv_DblClick()
    If selItem Is Nothing Then Exit Sub
    Dim tmp As String
    If chkCopyList.value = 1 Then
        tmp = Clipboard.GetText
    Else
        itemsCopied = 0
    End If
    Clipboard.Clear
    Clipboard.SetText tmp & selItem.Tag & vbCrLf
    itemsCopied = itemsCopied + 1
    Me.Caption = itemsCopied & " items copied"
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    Set selItem = Item
    sci.Text = Item.Tag
    frmMain.sci.GotoLineCentered CLng(Item.Text)
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuPopup
    End If
End Sub

Private Sub mnuClearClip_Click()
    itemsCopied = 0
    Clipboard.Clear
    Me.Caption = "Clipboard Cleared"
End Sub

Private Sub mnuCopySelected_Click()
    If lv.SelCount = 0 Then Exit Sub
    Dim li As ListItem
    Dim r() As String
    
    If chkCopyList.value = 1 Then
        If itemsCopied > 0 Then
            push r, Clipboard.GetText
        End If
    Else
        itemsCopied = 0
    End If
    
    
    For Each li In lv.selItems
        push r, li.Tag & vbCrLf
    Next
    
    Clipboard.Clear
    Clipboard.SetText Join(r, vbCrLf)
    itemsCopied = itemsCopied + UBound(r) + 1
    Me.Caption = "Copied " & itemsCopied & " rules"
    
End Sub

Private Sub mnuFindDups_Click()

    Dim li As ListItem, n As String
    Dim c As New CollectionEx, dups As New CollectionEx
    
    For Each li In lv.ListItems
        li.Selected = False
        n = trim(li.subItems(2))
        If c.keyExists(n) Then
            If dups.keyExists(n) Then
                dups(n, 1) = dups(n, 1) + 1
            Else
                dups.Add 1, n
            End If
            li.Selected = True
        Else
            c.Add 1, n
        End If
    Next
    
    If dups.count > 0 Then
        sci.Text = dups.count & " Duplicates: " & vbCrLf & dups.ToString(, True)
    Else
        MsgBox "No Duplicates", vbInformation
    End If
        
    
    
End Sub
