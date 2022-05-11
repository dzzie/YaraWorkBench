VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About Yara Workbench"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   10755
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5235
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   60
      Width           =   9315
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub DisplayData(x)
    text1.Text = x
    Me.Visible = True
End Sub

Private Sub Form_Load()
    Dim tmp As String, pth As String, x(), t As String
    
    Me.Icon = frmMain.Icon
    
    push x, ""
    push x, String(80, "-")
    push x, "   Yara Workbench v" & App.Major & "." & App.Minor & "." & App.Revision
    push x, "   Author:     David Zimmer <dzzie@yahoo.com>"
    push x, "   Site:       http://sandsprite.com"
    push x, "   Copyright:  2019 All Rights Reserved"
    push x, "   Yara Version: " & YaraVersion()
    push x, String(80, "-")
    push x, vbCrLf
    
    t = Join(x, vbCrLf)
    pth = App.path & "\Credits.txt"
    
    If fso.FileExists(pth) Then
        t = t & fso.ReadFile(pth)
    Else
        t = t & "Credits file not found: " & pth
    End If
    
    text1 = t
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    text1.Width = Me.Width - 240
    text1.Height = Me.Height - 300
End Sub
