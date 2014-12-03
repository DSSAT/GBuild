VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11190
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   8280
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   825
      ScaleWidth      =   2505
      TabIndex        =   13
      Top             =   4920
      Width           =   2535
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   8280
      Picture         =   "frmSplash.frx":7272
      ScaleHeight     =   825
      ScaleWidth      =   2505
      TabIndex        =   12
      Top             =   3960
      Width           =   2535
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   8280
      Picture         =   "frmSplash.frx":EA54
      ScaleHeight     =   825
      ScaleWidth      =   2505
      TabIndex        =   11
      Top             =   3000
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   8280
      Picture         =   "frmSplash.frx":161A6
      ScaleHeight     =   825
      ScaleWidth      =   2505
      TabIndex        =   10
      Top             =   2040
      Width           =   2535
   End
   Begin VB.PictureBox Picture22 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   240
      Picture         =   "frmSplash.frx":1D6FC
      ScaleHeight     =   1935
      ScaleWidth      =   2175
      TabIndex        =   8
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   0
      Tag             =   "1058"
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Department of Biological and Agricultural Engineering The University of Georgia"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   2280
      TabIndex        =   9
      Top             =   3225
      Width           =   5895
   End
   Begin VB.Label Label16 
      Caption         =   "Developed by:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Tag             =   "1156"
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Version: 4.6.0"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4762
      TabIndex        =   6
      Tag             =   "1045"
      Top             =   840
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Graphically display simulated and experimental data."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   5
      Tag             =   "1042"
      Top             =   1320
      Width           =   7680
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      Caption         =   "GBuild"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4830
      TabIndex        =   4
      Tag             =   "1048"
      Top             =   240
      Width           =   1530
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "                                 AgWeatherNet                                      Washington State University"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   4200
      Width           =   5655
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Agricultural and Biological Engineering Department University of Florida"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Tag             =   "1043"
      Top             =   2280
      Width           =   5655
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "DSSAT Foundation"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   5040
      Width           =   5655
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Height          =   6615
      Left            =   0
      Top             =   0
      Width           =   11185
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim k As Integer
Dim Readline
Dim Filex() As String


Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdTechnSupp_Click()
    frmTecnSupp.Show
End Sub

Private Sub Form_Activate()
Dim PauseTime
Dim Start
 If frmSplashShow = False Then
        PauseTime = 2  ' Set duration.
        Start = Timer   ' Set start time.
        Do While Timer < Start + PauseTime
            DoEvents   ' Yield to other processes.
        Loop
    Unload Me
Else
   ' Me.cmdTechnSupp.Visible = True
    Me.cmdOK.Visible = True
End If
    Dim XbuildPath As String
    If Right(App.path, 1) = "\" Then
        XbuildPath = App.path
    Else
        XbuildPath = App.path & "\"
    End If
    
    
        If Dir(XbuildPath & "WebSite.txt") <> "" Then
            Picture1.MouseIcon = LoadPicture(XbuildPath & "H_arrow.cur")
            Picture1.MousePointer = 99
           ' cmdEgyptWeb.MouseIcon = LoadPicture(XbuildPath & "H_arrow.cur")
           ' cmdEgyptWeb.MousePointer = 99
          '  Picture2.MouseIcon = LoadPicture(XbuildPath & "H_arrow.cur")
          '  Picture2.MousePointer = 99
            Picture3.MouseIcon = LoadPicture(XbuildPath & "H_arrow.cur")
            Picture3.MousePointer = 99
            Picture4.MouseIcon = LoadPicture(XbuildPath & "H_arrow.cur")
            Picture4.MousePointer = 99
           ' cmdICASA.MouseIcon = LoadPicture(XbuildPath & "H_arrow.cur")
           ' cmdICASA.MousePointer = 99
            'cmbUniversityWeb.MouseIcon = LoadPicture(XbuildPath & "H_arrow.cur")
           ' cmbUniversityWeb.MousePointer = 99
            cmdOK.Visible = True
          '  cmdTechnSupp.Visible = True
            
            On Error GoTo ErrorWeb
            
            Open XbuildPath & "WebSite.txt" For Input As #15
           
            
            k = 0
            Do While Not EOF(15)    'loop till the end of the file
                Line Input #15, Readline
                ReDim Preserve Filex(k + 1)
                k = k + 1
                Filex(k) = Readline  'puts it into memory
            Loop
            
            Close 15
        Else
           ' MsgBox lblCannotFindFile.Caption & " " & XbuildPath & "WebSite.txt." & _
            Chr(10) & lblNoWebPages.Caption, vbExclamation
            'cmdEgyptWeb.Enabled = False
            'cmdICASA.Enabled = False
           ' cmbUniversityWeb.Enabled = False
        End If

    
    'End If
    Exit Sub
ErrorWeb:
    MsgBox Err.Description & XbuildPath & "WebSite.txt."
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
    LoadResStrings Me
    
    ' VSH
    lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
    ' lblVersion.Caption = "Version " & "4.6.0.0"
    
    lblProductName.Caption = App.Title
End Sub

Private Sub cmbUniversityWeb_Click()

    If k > 0 Then
        If Mid$(Filex(1), 1, 4) = "http" Then
            WebPage = Filex(1)
           Shell ("c:\program files\internet explorer\iexplore.exe " & WebPage)
           ' frmWEbPages.Show
        Else
            'MsgBox lblWebAddress.Caption & " " & _
            'XbuildPath & " " & lblCheckFile.Caption, vbExclamation
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdICASA_Click()
    
    If k > 1 Then
        If Mid$(Filex(2), 1, 4) = "http" Then
            WebPage = Filex(2)
           Shell ("c:\program files\internet explorer\iexplore.exe " & WebPage)
           ' frmWEbPages.Show
        Else
            'MsgBox lblWebAddress.Caption & " " &
            'XbuildPath & " " & lblCheckFile.Caption, vbExclamation
            Exit Sub
        End If
    End If
End Sub

Private Sub Picture1_Click()
    If k > 2 Then
        If Mid$(Filex(3), 1, 4) = "http" Then
            WebPage = Filex(1)
           Shell ("c:\program files\internet explorer\iexplore.exe " & WebPage)
           ' frmWEbPages.Show
        Else
            'MsgBox lblWebAddress.Caption & " " & _
            'XbuildPath & " " & lblCheckFile.Caption, vbExclamation
            Exit Sub
        End If
    End If
End Sub

Private Sub Picture2_Click()
    If k > 2 Then
        If Mid$(Filex(3), 1, 4) = "http" Then
            WebPage = Filex(2)
            'frmWEbPages.Show
            Shell ("c:\program files\internet explorer\iexplore.exe " & WebPage)
        Else
            'MsgBox lblWebAddress.Caption & " " & _
            'XbuildPath & " " & lblCheckFile.Caption, vbExclamation
            Exit Sub
        End If
    End If
End Sub

Private Sub Picture3_Click()
    If k > 2 Then
        If Mid$(Filex(3), 1, 4) = "http" Then
            WebPage = Filex(3)
            Shell ("c:\program files\internet explorer\iexplore.exe " & WebPage)
            'frmWEbPages.Show
        Else
            'MsgBox lblWebAddress.Caption & " " & _
            'XbuildPath & " " & lblCheckFile.Caption, vbExclamation
            Exit Sub
        End If
    End If
End Sub

Private Sub Picture4_Click()
    If k > 2 Then
        If Mid$(Filex(3), 1, 4) = "http" Then
            WebPage = Filex(4)
           ' frmWEbPages.Show
            Shell ("c:\program files\internet explorer\iexplore.exe " & WebPage)
        Else
            'MsgBox lblWebAddress.Caption & " " & _
            'XbuildPath & " " & lblCheckFile.Caption, vbExclamation
            Exit Sub
        End If
    End If
End Sub

Private Sub Picture5_Click()
 If k > 2 Then
        If Mid$(Filex(3), 1, 4) = "http" Then
            WebPage = Filex(5)
            'frmWEbPages.Show
            Shell ("c:\program files\internet explorer\iexplore.exe " & WebPage)
        Else
            'MsgBox lblWebAddress.Caption & " " & _
            'XbuildPath & " " & lblCheckFile.Caption, vbExclamation
            Exit Sub
        End If
    End If
End Sub
