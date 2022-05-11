VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSampleList 
   Caption         =   "Build Sample List - Right click on listview for menu or drag and drop files/folders"
   ClientHeight    =   5295
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   3120
      TabIndex        =   9
      Top             =   4800
      Width           =   435
   End
   Begin VB.CommandButton cmdBrowseDir 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3660
      TabIndex        =   8
      Top             =   4800
      Width           =   495
   End
   Begin VB.CheckBox chkRecursive 
      Caption         =   "R"
      Height          =   255
      Left            =   4260
      TabIndex        =   7
      ToolTipText     =   "Recursive Scan (Folders)"
      Top             =   4800
      Width           =   435
   End
   Begin VB.TextBox txtCount 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1380
      TabIndex        =   6
      Top             =   4740
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "?"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   4860
      Width           =   435
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   4920
      TabIndex        =   3
      Top             =   4740
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load "
      Height          =   435
      Left            =   6480
      TabIndex        =   2
      Top             =   4740
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   435
      Left            =   7980
      TabIndex        =   1
      Top             =   4740
      Width           =   1215
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   8070
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
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
      OLEDropMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Count"
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   4800
      Width           =   495
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuLoadFile 
         Caption         =   "Load File"
      End
      Begin VB.Menu mnuLoadDir 
         Caption         =   "Load Directory"
      End
      Begin VB.Menu mnuRemoveSel 
         Caption         =   "Remove Selected"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "frmSampleList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'note set files contain file paths only so we dont have to rescan directory contents...?

Dim loadedFile As String
Dim fileCount As Long

Private Sub UpdateCount()
    fileCount = lv.ListItems.count
    txtCount = fileCount
End Sub

Private Sub cmdBrowse_Click()
    mnuLoadFile_Click
End Sub

Private Sub cmdBrowseDir_Click()
    mnuLoadDir_Click
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLoad_Click()
    Dim x, tmp
    x = dlg.OpenDialog()
    If Len(x) = 0 Then Exit Sub
    lv.ListItems.Clear
    LoadFile x
    UpdateCount
End Sub

Sub LoadFile(path)
    On Error Resume Next
    Dim x, tmp() As String
    If fso.FileExists(CStr(path)) Then
        loadedFile = path
        x = fso.ReadFile(loadedFile)
        tmp = Split(x, vbCrLf)
        For Each x In tmp
            lv.ListItems.Add , , x
        Next
    End If
    UpdateCount
End Sub


Private Sub cmdSave_Click()
    Dim x, tmp() As String, li As ListItem, ext As String
    x = dlg.SaveDialog(fso.FileNameFromPath(loadedFile), fso.GetParentFolder(loadedFile))
    If Len(x) = 0 Then Exit Sub
    ext = fso.GetExtension(x)
    If ext <> ".set" Then
        If ext = ".bin" Then 'default not user specified
            x = fso.ChangeExt(CStr(x), ".set")
        Else
            x = x & ".set"
        End If
    End If
    For Each li In lv.ListItems
        push tmp, li.Text
    Next
    fso.WriteFile CStr(x), Join(tmp, vbCrLf)
    frmMain.txtSample = x
End Sub

Private Sub Command1_Click()

    MsgBox "This form allows you to build a sample set file." & vbCrLf & _
            "" & vbCrLf & _
            "This can be individual files from various directories, or multiple different directories to all be scanned at the same time" & vbCrLf & _
            "" & vbCrLf & _
            "Using a sample set file can also increase scan time on huge sample sets because all of the files in a directory do not have to be enumerated first" & vbCrLf & _
            "" & vbCrLf & _
            "You can drag and drop files and folders on the listview to add them or use the right click menu.", vbInformation
            
End Sub

Private Sub Form_Load()
    mnuPopup.visible = False
    lv.ColumnHeaders(1).Width = lv.Width
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub lv_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, xx As Single, y As Single)
    On Error Resume Next
    Dim x
    For Each x In Data.Files
        lv.ListItems.Add , , x
    Next
    UpdateCount
End Sub

Private Sub mnuClear_Click()
    lv.ListItems.Clear
    UpdateCount
End Sub

Private Sub mnuLoadDir_Click()
    Dim x, tmp() As String
    x = dlg.FolderDialog2()
    If Len(x) > 0 Then
        tmp = fso.GetFolderFiles(CStr(x), , , IIf(chkRecursive.value = 1, True, False))
        For Each x In tmp
            lv.ListItems.Add , , x
        Next
        UpdateCount
    End If
End Sub

Private Sub mnuLoadFile_Click()
    Dim x
    x = dlg.OpenDialog()
    If Len(x) > 0 Then lv.ListItems.Add , , x
    UpdateCount
End Sub


Private Sub mnuRemoveSel_Click()
    Dim i
    For i = lv.ListItems.count To 1 Step -1
        If lv.ListItems(i).selected Then
            lv.ListItems.Remove (i)
        End If
    Next
    UpdateCount
End Sub
