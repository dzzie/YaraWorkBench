VERSION 5.00
Object = "{2668C1EA-1D34-42E2-B89F-6B92F3FF627B}#5.0#0"; "scivb2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLibrarian 
   Caption         =   "Library Manager"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   12165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpenDir 
      Caption         =   "Open Folder"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSaveNew 
      Caption         =   "Save New"
      Height          =   375
      Left            =   9180
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton mnuCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   10800
      TabIndex        =   2
      Top             =   120
      Width           =   1155
   End
   Begin sci2.SciSimple sci 
      Height          =   5835
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10292
   End
   Begin MSComctlLib.ListView lv 
      Height          =   6195
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   10927
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
         Text            =   "File"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmLibrarian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim selLi As ListItem
Dim homeDir As String

Private Sub cmdOpenDir_Click()
    On Error Resume Next
    Shell "explorer """ & homeDir & """", vbNormalFocus
End Sub

Private Sub cmdRemove_Click()

    If selLi Is Nothing Then
        MsgBox "No item selected"
        Exit Sub
    End If
        
    If MsgBox("Are you sure you want to delete " & selLi.Text & "?", vbYesNo + vbInformation) = vbNo Then Exit Sub
    
    sci.Text = Empty
    fso.DeleteFile CStr(selLi.Tag)
    lv.ListItems.Remove selLi.Index
    Set selLi = Nothing
    
End Sub

 

Private Sub cmdSave_Click()
    If selLi Is Nothing Then
        cmdSaveNew_Click
    Else
        If Not sci.SaveFile(CStr(selLi.Tag)) Then
            MsgBox "Save file failed"
        End If
    End If
End Sub

Private Sub cmdSaveNew_Click()
    On Error Resume Next
    Dim f As String, li As ListItem, exists As Boolean
    f = dlg.SaveDialog("", homeDir)  'todo: extract rule name for def file name...
    If Len(f) = 0 Then Exit Sub
    If sci.SaveFile(f) Then
        For Each li In lv.ListItems
            If li.Tag = f Then
                exists = True
                Exit Sub
            End If
        Next
        If Not exists Then
            Set li = lv.ListItems.Add(, , fso.GetBaseName(f))
            li.Tag = f
        End If
    Else
        MsgBox "Save new failed"
    End If
End Sub

Private Sub Form_Load()
    Dim li As ListItem
    Dim pth As String
    Dim x() As String
    Dim xx
    
    Me.Icon = frmMain.Icon
    SyntaxColor sci
    pth = App.path & "\library"
    lv.ColumnHeaders(1).Width = lv.Width - 30
    
    If Not fso.FolderExists(pth) Then
        Me.Caption = "Library folder not found: " & pth
        Exit Sub
    End If
    
    homeDir = pth
    Me.Caption = "Library folder: " & pth
    x = fso.GetFolderFiles(pth)
    If AryIsEmpty(x) Then Exit Sub
    
    For Each xx In x
        Set li = lv.ListItems.Add(, , fso.GetBaseName(CStr(xx)))
        li.Tag = xx
    Next
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    sci.Width = Me.Width - sci.Left - 240
    sci.Height = Me.Height - sci.Top - 700
    lv.Height = Me.Height - lv.Top - 700
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    Dim f As String
    Set selLi = Item
    f = Item.Tag
    If fso.FileExists(f) Then sci.LoadFile f
End Sub

Private Sub mnuCopy_Click()
    Clipboard.Clear
    Clipboard.SetText sci.Text
End Sub
