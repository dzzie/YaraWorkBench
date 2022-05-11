VERSION 5.00
Object = "{2668C1EA-1D34-42E2-B89F-6B92F3FF627B}#5.0#0"; "scivb2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmModuleDump 
   Caption         =   "Module Dump Info    Find = CTRL+F"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   12285
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7380
      TabIndex        =   8
      Top             =   60
      Width           =   495
   End
   Begin MSComctlLib.ListView lv 
      Height          =   6915
      Left            =   60
      TabIndex        =   7
      Top             =   540
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   12197
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
   Begin VB.CheckBox chkWrap 
      Caption         =   "Wrap"
      Height          =   315
      Left            =   9540
      TabIndex        =   6
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   6900
      TabIndex        =   0
      Top             =   60
      Width           =   435
   End
   Begin sci2.SciSimple sci 
      Height          =   6915
      Left            =   2520
      TabIndex        =   5
      Top             =   540
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   12197
   End
   Begin VB.CommandButton cmdDump 
      Caption         =   "Dump"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10500
      TabIndex        =   4
      Top             =   60
      Width           =   1455
   End
   Begin VB.ComboBox cbo 
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
      Left            =   7980
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   60
      Width           =   1455
   End
   Begin VB.TextBox txtSample 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   660
      TabIndex        =   2
      Text            =   "Drag & Drop, browse, or type"
      Top             =   60
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "Sample:"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "frmModuleDump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkWrap_Click()
    Static ignore As Boolean
    If ignore Then Exit Sub
    ignore = True
    If chkWrap.value = vbChecked Then
        chkWrap.value = vbChecked
        sci.WordWrap = True
    Else
        chkWrap.value = vbUnchecked
        sci.WordWrap = False
    End If
    ignore = False
End Sub

Private Sub cmdBrowse_Click()
    Dim f As String
    f = dlg.OpenDialog()
    If Len(f) = 0 Then Exit Sub
    txtSample_GotFocus
    txtSample = f
End Sub

Private Sub cmdBrowseDir_Click()
    On Error Resume Next
    Dim f As String
    f = dlg.FolderDialog2()
    If Len(f) = 0 Then Exit Sub
    txtSample_GotFocus
    txtSample = f
End Sub

Private Sub cmdDump_Click()
    Dim scan As New CYaraScan
    Dim c As Collection, f As CYaraFile
    Dim x, li As ListItem
    
    lv.ListItems.Clear
    sci.Text = Empty
    If Len(txtSample) = 0 Or txtSample.ForeColor <> vbBlack Then
        MsgBox "Enter sample file or folder to scan."
    End If
    
    Set c = scan.scan(txtSample, "import """ & cbo.Text & """" & vbCrLf, True)
    
    If scan.errors.count > 0 Then
        sci.Text = "Errors: " & scan.errors.ToString()
    End If
    
    For Each f In c
        Set li = lv.ListItems.Add(, , fso.FileNameFromPath(f.File))
        Set li.Tag = f
    Next
    
    On Error Resume Next
    lv_ItemClick lv.ListItems(1)
    
End Sub

Private Sub Form_Load()
    Dim i As CIntellisenseItem, ii As Long, found As Boolean
    Dim yf As CYaraFile
    
    sci.WordWrap = False
    Me.Icon = frmMain.Icon
    SHAutoComplete txtSample.hwnd, SHACF_FILESYSTEM
    lv.ColumnHeaders(1).Width = lv.Width * 2
    
    For Each i In frmMain.isense
        cbo.AddItem i.ObjName
        If i.ObjName = "pe" Then found = True
        If Not found Then ii = ii + 1
    Next
    
    cbo.ListIndex = ii
    
    If fso.FileExists(frmMain.txtSample) Then
        txtSample.ForeColor = vbBlack
        txtSample = frmMain.txtSample
    ElseIf Not frmMain.lv.SelectedItem Is Nothing Then
        Set yf = frmMain.lv.SelectedItem.Tag
        txtSample.ForeColor = vbBlack
        txtSample = yf.File
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    sci.Width = Me.Width - sci.Left - 240
    sci.Height = Me.Height - sci.Top - 700
    lv.Height = sci.Height
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim f As CYaraFile
    On Error Resume Next
    Set f = Item.Tag
    sci.Text = f.DumpInfo
End Sub

Private Sub txtSample_GotFocus()
    On Error Resume Next
    If txtSample.ForeColor <> vbBlack Then
        txtSample.Text = Empty
        txtSample.ForeColor = vbBlack
    End If
End Sub

Private Sub txtSample_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    txtSample_GotFocus
    txtSample = Data.Files(1)
End Sub
