VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3570
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function dbg(x, ParamArray y())
    Dim z, i
    z = x & " = "
    For i = 0 To UBound(y)
        z = z & Hex(y(i)) & " "
    Next
    List1.AddItem z
End Function

Private Sub Form_Load()
    
    Dim cc As Long, ret As Long
    Dim realval As Long
    Dim warnings As Long
    
    Dim yara As String
    
    yara = "import ""pe"" rule test { condition: pe.imports(""KERNEL32.dll"", ""DeleteCriticalSection"")}"
    
    realval = VarPtr(cc)
    
    ret = yr_initialize()
    dbg "yr_initialize", ret, cc, realval
    
    ret = yr_compiler_create(cc)
    dbg "yr_compiler_create", ret, cc, realval
    
    Call yr_compiler_set_callback(realval, AddressOf callback, 0)
     
    If yr_compiler_add_string(cc, StrPtr(yara), 0) <> 0 Then
        'result = compiler->last_error;
        'goto _exit;
    End If
  
    
      
    ret = yr_finalize()
    
    
    
End Sub
