VERSION 5.00
Begin VB.Form frmTecnSupp 
   Caption         =   "Technical Support"
   ClientHeight    =   9060
   ClientLeft      =   2580
   ClientTop       =   1815
   ClientWidth     =   11025
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   11025
   Tag             =   "3239"
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   3360
      TabIndex        =   6
      Tag             =   "1156"
      Top             =   240
      Width           =   5775
      Begin VB.Label Label3 
         Caption         =   "Oxana Uryasev"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Tag             =   "1150"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Agricultural and Biological Engineering Department University of Florida"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   17
         Tag             =   "1043"
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "James W. Jones"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Tag             =   "1149"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Agricultural and Biological Engineering Department University of Florida"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   15
         Tag             =   "1043"
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "Cheryl Porter"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   14
         Tag             =   "1152"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Agricultural and Biological Engineering Department University of Florida"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3240
         TabIndex        =   13
         Tag             =   "1043"
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label lblJones 
         Caption         =   "jimj@ufl.edu"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   480
         TabIndex        =   12
         Top             =   495
         Width           =   2055
      End
      Begin VB.Label lblGerrit 
         Caption         =   "gerrit.hoogenboom@wsu.edu"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3360
         TabIndex        =   11
         Top             =   495
         Width           =   2055
      End
      Begin VB.Label lblCheryl 
         Caption         =   "cporter@ufl.edu"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3360
         TabIndex        =   10
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label lblOxana 
         Caption         =   "oxana@ufl.edu"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   480
         TabIndex        =   9
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Gerrit Hoogenboom"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   8
         Tag             =   "1151"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "AgWeatherNet Washington State University"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         TabIndex        =   7
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      Picture         =   "frmTecnSupp.frx":0000
      ScaleHeight     =   2025
      ScaleWidth      =   2745
      TabIndex        =   5
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Left            =   8280
      TabIndex        =   0
      Tag             =   "1058"
      Top             =   5400
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2535
      Left            =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label14 
      Caption         =   $"frmTecnSupp.frx":11EF2
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   4
      Tag             =   "1153"
      Top             =   4200
      Width           =   6015
   End
   Begin VB.Label Label10 
      Caption         =   "Pictures are credited to: "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Tag             =   "1155"
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label13 
      Caption         =   "Oxana Uryasev"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label12 
      Caption         =   "Programed by:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Tag             =   "1154"
      Top             =   3840
      Width           =   2055
   End
End
Attribute VB_Name = "frmTecnSupp"
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

Private Sub Form_Load()
 Dim NewPath As String
     Me.Width = 9450
    Me.Height = 6400

    ' If gsglXFactor <> 1 Or gsglYFactor <> 1 Then
        Call SetDeviceIndependentWindow(Me, 1)
   ' End If
     Top = 0
     Left = 0

 If Right(App.path, 1) = "\" Then
    NewPath = Mid$(App.path, 1, Len(App.path) - 1)
 Else
    NewPath = App.path
 End If
    
    LoadResStrings Me
      
       ' lblOxana.MouseIcon = LoadPicture(NewPath & "\H_arrow.cur")
       ' lblOxana.MousePointer = 99
       ' lblGerrit.MouseIcon = LoadPicture(NewPath & "\H_arrow.cur")
       ' lblGerrit.MousePointer = 99
       ' lblJones.MouseIcon = LoadPicture(NewPath & "\H_arrow.cur")
       'lblJones.MousePointer = 99
        'lblCheryl.MouseIcon = LoadPicture(NewPath & "\H_arrow.cur")
       ' lblCheryl.MousePointer = 99
       
    With Shape1
        .Top = Picture1.Top - 120
        .Left = Picture1.Left - 120
        .Width = Picture1.Width + 240
        .Height = Picture1.Height + 240
    End With
   cmdOK.Top = Me.Height - (3 + 1 / 4) * cmdOK.Height
   cmdOK.Left = Me.Width - (1 + 1 / 4) * cmdOK.Width
       ' Unload frmEMail


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 15
        cmdOK_Click
    End Select

End Sub



Private Sub lblCheryl_Click()
   ' prmEMAIL = lblCheryl.Caption
    ' Screen.MousePointer = vbHourglass
   ' frmEMail.Show
End Sub

Private Sub lblGerrit_Click()
   ' prmEMAIL = lblGerrit.Caption
   '  Screen.MousePointer = vbHourglass
   ' frmEMail.Show
End Sub

Private Sub lblJones_Click()
  '  prmEMAIL = lblJones.Caption
  '   Screen.MousePointer = vbHourglass
  '  frmEMail.Show
End Sub


Private Sub lblOxana_Click()
   ' prmEMAIL = lblOxana.Caption
   '  Screen.MousePointer = vbHourglass
   ' frmEMail.Show
End Sub
