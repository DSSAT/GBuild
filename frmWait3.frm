VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWait3 
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4860
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3480
      Picture         =   "frmWait3.frx":0000
      ScaleHeight     =   1095
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar prgLoad 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Reloading Data... Please wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Tag             =   "1147"
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "frmWait3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
 Dim me_Height
    me_Height = Me.Height
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 1000
    LoadResStrings Me
   ' If gsglXFactor <> 1 Or gsglYFactor <> 1 Then
        Call SetDeviceIndependentWindow(Me, 1)
   ' End If
    Picture2.Left = Me.Label1.Left + Label1.Width + 100
    Me.Height = me_Height
  '  prgLoad.Value = 10
End Sub
