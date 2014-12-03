VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWait 
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4650
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar prgLoad 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3360
      Picture         =   "frmWait.frx":0000
      ScaleHeight     =   1095
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Creating the graph. Please wait..."
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
      TabIndex        =   0
      Tag             =   "1137"
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "frmWait"
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
  '  If gsglXFactor <> 1 Or gsglYFactor <> 1 Then
        Call SetDeviceIndependentWindow(Me, 1)
  '  End If
    Picture2.Left = Me.Label1.Left + Label1.Width + 100
    Me.Height = me_Height
   ' Me.prgLoad.Value = 10
End Sub

