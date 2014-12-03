VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmFilePreview 
   Caption         =   "File Preview"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Tag             =   "1003"
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8040
      TabIndex        =   1
      Tag             =   "1058"
      Top             =   6120
      Width           =   855
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   10398
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   50000
      TextRTF         =   $"frmFilePreview.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbltoolClose 
      Caption         =   "Label1"
      Height          =   135
      Left            =   5520
      TabIndex        =   2
      Tag             =   "1014"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmFilePreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    If Dir(App.path & "\Manual_GBuild.chm") <> "" Then
        App.HelpFile = App.path & "\Manual_GBuild.chm"
        Call HtmlHelp(Me.hWnd, App.HelpFile, HH_DISPLAY_TOC, 0)
    Else
        MsgBox "Cannot find" & " " & App.path & "\Manual_GBuild.chm"
    End If
 End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 15
        cmdOK_Click
    End Select

End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Width = 9450
    Me.Height = 7200

   ' If gsglXFactor <> 1 Or gsglYFactor <> 1 Then
        Call SetDeviceIndependentWindow(Me, 1)
   ' End If
 
    
    'LoadResStrings Me
    RichTextBox1.Left = 5
    RichTextBox1.Width = Me.Width - 100
    RichTextBox1.Top = 120
    RichTextBox1.Height = Me.Height - _
        3100
        
    RichTextBox1.FileName = PreviewFile
    Me.cmdOK.Top = RichTextBox1.Top + RichTextBox1.Height + _
        800
    With cmdOK
        .Left = Me.Width - (1 + 1 / 4) * .Width
    End With
    cmdOK.ToolTipText = lbltoolClose.Caption
    Unload frmDocument
End Sub

Private Sub RichTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    If Dir(App.path & "\Manual_GBuild.chm") <> "" Then
        App.HelpFile = App.path & "\Manual_GBuild.chm"
        Call HtmlHelp(Me.hWnd, App.HelpFile, HH_DISPLAY_TOC, 0)
    Else
        MsgBox "Cannot find" & " " & App.path & "\Manual_GBuild.chm"
    End If
 End If

End Sub
