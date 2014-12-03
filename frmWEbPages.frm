VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmWEbPages 
   Caption         =   "Web Page"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
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
      Height          =   255
      Left            =   7680
      TabIndex        =   0
      Tag             =   "1058"
      Top             =   7680
      Width           =   735
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7575
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10815
      ExtentX         =   19076
      ExtentY         =   13361
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmWEbPages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()
    Unload Me
    frmSplash.Show
End Sub


Private Sub Form_Load()
    On Error GoTo ErrorWeb
    Me.Top = 0
    Me.Left = 0
    
    Me.Height = 6400
    Me.Width = 9450
   LoadResStrings Me
    
  
        Call SetDeviceIndependentWindow(Me, 1)
  
    With WebBrowser1
         .Top = 0
         .Left = 0
         .Height = Height - 1000 * gsglYFactor
         .Width = Width - 200 * gsglYFactor
        .Navigate WebPage
        cmdSave.Top = .Top + .Height + 100 * gsglYFactor
    End With
    Exit Sub
ErrorWeb:
    MsgBox Err.Description, vbExclamation
End Sub

