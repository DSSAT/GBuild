VERSION 5.00
Begin VB.Form frmDocument 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11175
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "DSSAT v 4.6.0.0"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   3600
      Width           =   2385
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000C&
      Caption         =   "GBuild"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   5160
      TabIndex        =   0
      Top             =   2040
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   3420
      Left            =   840
      Picture         =   "frmDocument2.frx":0000
      Top             =   840
      Width           =   4050
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub Form_Load()
    
    'If gsglXFactor <> 1 Or gsglYFactor <> 1 Then
    '    Call SetDeviceIndependentWindow(Me, 1)
    'End If
    
    ' VSH
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Label2.Caption = "DSSAT v " & fso.GetFileVersion("C:\DSSAT46\dssat46.exe")
    
    
    Me.Left = (MainWidth - Me.Width) / 2
    Me.Top = (MainHeight - Me.Height) / 2 - 1000
'Public MainHeight As Single
'Public MainWidth As Single

End Sub

