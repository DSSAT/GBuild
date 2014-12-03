VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions1 
   Caption         =   "Select option"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTab 
      Height          =   3615
      Index           =   1
      Left            =   3960
      TabIndex        =   2
      Top             =   1320
      Width           =   5415
      Begin VB.Frame Frame5 
         Caption         =   "Simulated Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2520
         TabIndex        =   23
         Tag             =   "1089"
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         Begin VB.OptionButton OptSimDontPlot 
            Caption         =   "Don't plot"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   240
            TabIndex        =   25
            Tag             =   "1081"
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton OptSimPlot 
            Caption         =   "Plot"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Tag             =   "1080"
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Statistic"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2880
         TabIndex        =   21
         Tag             =   "1082"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1575
         Begin VB.OptionButton OptStatShow 
            Caption         =   "Show"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Tag             =   "1088"
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Experimental Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   600
         TabIndex        =   18
         Top             =   480
         Width           =   1455
         Begin VB.OptionButton OptPlot 
            Caption         =   "Plot"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Tag             =   "1080"
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton OptDontPlot 
            Caption         =   "Don't plot"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Tag             =   "1081"
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.CheckBox ChkExpShow 
         Caption         =   "Save on Start/Load program"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Tag             =   "1079"
         Top             =   1920
         Width           =   3615
      End
   End
   Begin VB.Frame fraTab 
      Height          =   3615
      Index           =   3
      Left            =   1080
      TabIndex        =   9
      Top             =   1560
      Width           =   5415
      Begin VB.Frame Frame7 
         Caption         =   "Gridlines"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   480
         TabIndex        =   39
         Tag             =   "1141"
         Top             =   1320
         Width           =   2055
         Begin VB.OptionButton OptGridShow 
            Caption         =   "Show"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Tag             =   "1142"
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton OptGridDontShow 
            Caption         =   "Don't show"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Tag             =   "1143"
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Marker Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2760
         TabIndex        =   36
         Tag             =   "1138"
         Top             =   360
         Width           =   2055
         Begin VB.TextBox txtMarkerSize 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   42
            Top             =   600
            Width           =   495
         End
         Begin VB.OptionButton OptLarge 
            Caption         =   "Custom"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Tag             =   "1140"
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton OptSmall 
            Caption         =   "Default (1)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Tag             =   "1139"
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Simulated Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   480
         TabIndex        =   26
         Tag             =   "1120"
         Top             =   360
         Width           =   2055
         Begin VB.OptionButton optShowLine 
            Caption         =   "Show Line"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Tag             =   "1118"
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optDontShowLine 
            Caption         =   "Don't show line"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Tag             =   "1119"
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.CheckBox ChkEditChart 
         Caption         =   "Save on Start/Load program"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Tag             =   "1079"
         Top             =   2400
         Width           =   3615
      End
   End
   Begin VB.Frame fraTab 
      Height          =   3615
      Index           =   0
      Left            =   4680
      TabIndex        =   1
      Top             =   1080
      Width           =   5415
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   2760
         TabIndex        =   15
         Top             =   360
         Width           =   2295
         Begin VB.OptionButton OptDaysOfYear 
            Caption         =   "Day of Year"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   43
            Tag             =   "1144"
            Top             =   960
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.OptionButton optDate 
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Tag             =   "1075"
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton OptDaysAfterPlant 
            Caption         =   "Days after planting"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Tag             =   "1076"
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1455
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2295
         Begin VB.CheckBox chkSimExp 
            Caption         =   "Simulated Data vs.Experimental Data"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   27
            Tag             =   "1113"
            Top             =   840
            Width           =   1935
         End
         Begin VB.OptionButton OptTime 
            Caption         =   "Time Series plot"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Tag             =   "1095"
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton OptGenetic 
            Caption         =   "Scatter Plot"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Tag             =   "1145"
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.CheckBox ChkXAxis 
         Caption         =   "Save on Start/Load program"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Tag             =   "1079"
         Top             =   2040
         Width           =   2535
      End
   End
   Begin VB.Frame fraTab 
      Height          =   3615
      Index           =   2
      Left            =   3000
      TabIndex        =   3
      Top             =   720
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "Browse..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4440
         TabIndex        =   35
         Tag             =   "1109"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtFilePath 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   34
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CommandButton cmbBrowse 
         Caption         =   "Browse..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4440
         TabIndex        =   31
         Tag             =   "1109"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtFolder 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   30
         Top             =   480
         Width           =   2535
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4080
         Top             =   3720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CheckBox ChkFile 
         Caption         =   "Save on Start/Load program"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Tag             =   "1079"
         Top             =   1800
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "File Path"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Database Path"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "Defaults"
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
      Left            =   9720
      TabIndex        =   11
      Tag             =   "1087"
      Top             =   5880
      Width           =   855
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
      Height          =   315
      Left            =   10680
      TabIndex        =   7
      Tag             =   "1058"
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   8760
      TabIndex        =   6
      Tag             =   "1059"
      Top             =   5880
      Width           =   855
   End
   Begin MSComctlLib.TabStrip tabRTF 
      Height          =   6015
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   10610
      MultiSelect     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Graph Type"
            Object.Tag             =   "1085"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Data Series"
            Object.Tag             =   "1086"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Data Files"
            Object.Tag             =   "1105"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit Chart"
            Object.Tag             =   "1117"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H8000000A&
      Height          =   5775
      Left            =   0
      Top             =   0
      Width           =   9375
   End
End
Attribute VB_Name = "frmOptions1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim default_prmShowLine As Integer
Dim default_ShowExperimentalData As Integer
Dim default_ShowStatistic As Integer
Dim default_ApplicationDrive As Integer
Dim default_DataBasePath As String
Dim default_ShowX_Axis As Integer
Dim default_ExpData_vs_Simulated As Integer
Dim default_ShowRealDates As Integer
Dim default_Marker_Small As Integer
Dim default_Show_Grid As Integer
Dim SetDefaults As Boolean

Private Sub chkSimExp_Click()
    If chkSimExp.Value = 1 Then
        optShowLine.Value = False
        optDontShowLine.Value = True
        OptPlot.Value = True
        OptDontPlot.Value = False
        OptTime.Value = False
        OptGenetic.Value = True
    End If
End Sub

Private Sub cmbBrowse_Click()
    Screen.MousePointer = vbHourglass
    With CommonDialog1
        .CancelError = True
        On Error GoTo ProcExit
        .Filter = "All Files (*.*)|*.*"
        .FileName = ""
        .Flags = cdlOFNHideReadOnly
        .ShowOpen
        Screen.MousePointer = vbDefault
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        If InStrRev(.FileName, "\") > 1 Then
            txtFolder.Text = Mid(.FileName, 1, InStrRev(.FileName, "\") - 1)
        End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ProcExit:
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDefaults_Click()
    
    prmShowLine = default_prmShowLine
    ShowExperimentalData = default_ShowExperimentalData
   ' ShowStatistic = default_ShowStatistic
    default_ApplicationDrive = 1
   ' DataBasePath = default_DataBasePath
   ' PartOfFullPath = default_PartOfFullPath
    ShowX_Axis = default_ShowX_Axis
    Marker_Small = default_Marker_Small
    Show_Grid = default_Show_Grid
    ExpData_vs_Simulated = default_ExpData_vs_Simulated
    ShowRealDates = default_ShowRealDates
   ' txtFolder.Text = default_DataBasePath
   
   ' OptDataPath.Value = True
    SetDefaults = True
    
    'cmdOK_Click
End Sub

Private Sub cmdOK_Click()
        
    If IsNumeric(txtMarkerSize.Text) = False Then
        OptSmall.Value = True
        Me.OptLarge.Value = False
        txtMarkerSize.Text = ""
    Else
        OptSmall.Value = False
        OptLarge.Value = True
        Marker_Size = Val(txtMarkerSize.Text)
    End If
    
    If FileType = "T-file" Then
        optShowLine.Value = False
        optDontShowLine.Value = True
        OptPlot.Value = True
        OptDontPlot.Value = False
       ' optDate.Value = True
        OptDaysAfterPlant.Value = False
       ' OptTime.Value = True
    End If
    
    If OptPlot.Value = True Then
        ShowExperimentalData = 1
    Else
        ShowExperimentalData = 0
        ExpData_vs_Simulated = 0
    End If
   ' If OptStatShow.Value = True Then
        ShowStatistic = 1
   ' Else
      '  ShowStatistic = 0
   ' End If
    If optDate.Value = True Then
        ShowRealDates = 1
    ElseIf OptDaysAfterPlant.Value = True Then
        ShowRealDates = 0
    ElseIf OptDaysOfYear.Value = True Then
        ShowRealDates = 2
    End If
    If OptTime.Value = True Then
        ShowX_Axis = 1
        ExpData_vs_Simulated = 0
    Else
        ShowX_Axis = 0
    End If
    
    If chkSimExp.Value = 1 Then
        ExpData_vs_Simulated = 1
        ShowX_Axis = 0
        ShowExperimentalData = 1
    Else
        ExpData_vs_Simulated = 0
    End If
    
    If optShowLine.Value = True Then
       prmShowLine = 1
       ExpData_vs_Simulated = 0
    Else
        prmShowLine = 0
    End If
   
    If OptSmall.Value = True Then
        Marker_Small = 1
    Else
        Marker_Small = Val(Me.txtMarkerSize.Text)
    End If
    
    If Me.OptGridShow.Value = True Then
        Show_Grid = 1
    Else
        Show_Grid = 0
    End If
    
    If ShowX_Axis = 0 Then
        prmShowLine = 0
    End If
   
    New_DataBasePath = Trim(txtFolder.Text)
    New_FilePath = Trim(txtFilePath.Text)

    
    If ChkEditChart.Value = 1 Then
        file_prmShowLine = prmShowLine
        file_Marker_Small = Marker_Small
        file_Show_Grid = Show_Grid
    End If
    If ChkExpShow.Value = 1 Then
        file_ShowExperimentalData = ShowExperimentalData
        file_ShowStatistic = 1
    End If
    If ChkFile.Value = 1 Then
        'file_ApplicationDrive = ApplicationDrive
        file_DataBasePath = Trim(txtFolder.Text)
        file_FilePath = Trim(txtFilePath.Text)
       ' file_PartOfFullPath = PartOfFullPath
    End If
    If ChkXAxis.Value = 1 Then
        file_ShowX_Axis = ShowX_Axis
        file_ExpData_vs_Simulated = ExpData_vs_Simulated
        file_ShowRealDates = ShowRealDates
    End If
        
    
    Open ApplicationPathOption & "\Option.txt" For Output As #11
        Print #11, file_ShowExperimentalData, file_ShowStatistic, file_ShowRealDates, _
            file_ShowX_Axis, file_prmShowLine, file_ExpData_vs_Simulated, file_Show_Grid, file_Marker_Small
        If default_ApplicationDrive = 1 Or file_DataBasePath = "" Then
            Print #11, "@1"
        ElseIf file_FilePath <> "" Then
            Print #11, "@" & file_DataBasePath & "," & file_FilePath
        Else
            Print #11, "@" & file_DataBasePath
        End If
        Close 11
        If PaulsVersion Then frmOpenFileShown = True
        
        Unload frmGraph
        
        Unload frmSelection
        If SetDefaults = False And frmOpenFileShown = True Then
        'If ExpData_vs_Simulated = 0 And SetDefaults = False Then
            On Error Resume Next
            If CloseISselected = False Then
                frmSelection.Show
            End If
        End If

    Unload Me
End Sub

Private Sub Command1_Click()
    Screen.MousePointer = vbHourglass
    With CommonDialog1
        .CancelError = True
        On Error GoTo ProcExit
        .Filter = "All Files (*.*)|*.*"
        '.InitDir = ApplicationDrive & ":" & "\" & DataBasePath
        .FileName = ""
        .Flags = cdlOFNHideReadOnly
        .ShowOpen
        Screen.MousePointer = vbDefault
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        If InStrRev(.FileName, "\") > 1 Then
            txtFilePath.Text = Mid(.FileName, 1, InStrRev(.FileName, "\") - 1)
        End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ProcExit:
Screen.MousePointer = vbDefault

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
    Case 14
        cmdCancel_Click
    Case 15
        cmdOK_Click
    Case 4
        cmdDefaults_Click
    End Select

End Sub

Private Sub Form_Load()
Dim i
    LoadResStrings Me
    Me.Top = 0
    Me.Left = 0
    Me.Width = 9450
    Me.Height = 6400
   ' If gsglXFactor <> 1 Or gsglYFactor <> 1 Then
        Call SetDeviceIndependentWindow(Me, 1)
   ' End If
    If Marker_Size <> 0 Then txtMarkerSize.Text = Marker_Size
    
    If NumberOfOUTfiles > 1 Then
        Me.chkSimExp.Enabled = False
    End If
    
    If Marker_Size = 0 Then
        Marker_Small = 1
    Else
        Marker_Small = Val(txtMarkerSize.Text)
    End If

    default_ApplicationDrive = 0
    SetDefaults = False
    If ShowExperimentalData = 1 Then
        OptPlot.Value = True
        OptDontPlot.Value = False
    Else
        OptPlot.Value = False
        OptDontPlot.Value = True
        chkSimExp.Value = 0
    End If
    If ShowRealDates = 1 Then
        optDate.Value = True
        OptDaysAfterPlant.Value = False
        OptDaysOfYear.Value = False
    ElseIf ShowRealDates = 0 Then
        optDate.Value = False
        OptDaysAfterPlant.Value = True
        OptDaysOfYear.Value = False
    ElseIf ShowRealDates = 2 Then
        optDate.Value = False
        OptDaysAfterPlant.Value = False
        OptDaysOfYear.Value = True
    End If
    If ShowX_Axis = 1 Then
        OptTime.Value = True
        OptGenetic.Value = False
        chkSimExp.Value = 0
    Else
        OptTime.Value = False
        OptGenetic.Value = True
    End If
    
    If ExpData_vs_Simulated = 1 Then
        chkSimExp.Value = 1
        OptGenetic.Value = True
        OptPlot.Value = True
    Else
        chkSimExp.Value = 0
    End If
    
    If ExpExists = False Then
        chkSimExp.Enabled = False
    End If

    
    If FileType = "SUM" Then
        prmShowLine = 0
        OptTime.Value = False
        OptGenetic.Value = True
        OptTime.Enabled = False
    End If
    If prmShowLine = 1 Then
        optShowLine.Value = True
        chkSimExp.Value = 0
        optDontShowLine.Value = False
    Else
        optShowLine.Value = False
        optDontShowLine.Value = True
    End If
    
    If Marker_Small = 1 Then
        OptSmall.Value = True
        OptLarge.Value = False
    Else
        OptSmall.Value = False
        OptLarge.Value = True
    End If
    
    If Show_Grid = 1 Then
        OptGridShow.Value = True
        OptGridDontShow = False
    Else
        OptGridShow.Value = False
        OptGridDontShow = True
    End If
    
    txtFolder.Text = New_DataBasePath
    txtFilePath.Text = New_FilePath
    
    
    If FileType = "T-file" Then
        optShowLine.Value = False
        optDontShowLine.Value = True
        OptPlot.Value = True
        OptDontPlot.Value = False
        optDate.Value = True
        OptDaysAfterPlant.Value = False
        OptGenetic.Enabled = False
        chkSimExp.Enabled = False
        OptDaysAfterPlant.Enabled = False
    End If
   
    ChkEditChart.Value = 0
    ChkExpShow.Value = 0
    ChkFile.Value = 0
    ChkXAxis.Value = 0
   
    default_prmShowLine = 1
    default_ShowExperimentalData = 1
    default_ShowStatistic = 1
    
    If Dir(Mid(App.path, 1, InStr(App.path, ":") - 1) & ":" & "\" & "DSSAT46" & "\DATA.CDE") <> "" Then
         default_DataBasePath = "DSSAT46"
    'ElseIf Dir(Mid(App.path, 1, InStr(App.path, ":") - 1) & ":" & "\" & "DSSAT35" & "\DATA.CDE") <> "" Then
       '   default_DataBasePath = "DSSAT35"
    End If
    default_ShowX_Axis = 1
    default_ExpData_vs_Simulated = 0
    default_ShowRealDates = 1
    default_Marker_Small = 1
    default_Show_Grid = 0
    For i = 0 To fraTab.Count - 1
    With fraTab(i)
      .Move tabRTF.ClientLeft, _
      tabRTF.ClientTop, _
      tabRTF.ClientWidth, _
      tabRTF.ClientHeight
    End With
    Next i
    If PaulsVersion = True Then
        txtFilePath.Enabled = False
        txtFolder.Enabled = False
        ChkFile.Enabled = False
        Label2.Enabled = False
        Label1.Enabled = False
        cmbBrowse.Enabled = False
        Command1.Enabled = False
    Else
        txtFilePath.Enabled = True
        txtFolder.Enabled = True
        ChkFile.Enabled = True
        Label2.Enabled = True
        Label1.Enabled = True
        cmbBrowse.Enabled = True
        Command1.Enabled = True
    End If
   Shape2.Height = Me.Height - 400
   ' Bring the first fraTab control to the front.
   fraTab(0).ZOrder 0
    With cmdCancel
        .Top = Me.Height - (3 + 1 / 4) * .Height
        .Left = Me.Width - (1 + 3 / 4) * .Width - cmdOK.Width - cmdDefaults.Width
    End With
    With cmdDefaults
        .Top = Me.Height - (3 + 1 / 4) * .Height
        .Left = Me.Width - (1 + 2 / 4) * .Width - cmdOK.Width
    End With
    With cmdOK
        .Top = Me.Height - (3 + 1 / 4) * .Height
        .Left = Me.Width - (1 + 1 / 4) * .Width
    End With
    
    If CloseISselected = True Then
        Me.OptTime.Enabled = False
        Me.OptGenetic.Enabled = False
        Me.chkSimExp.Enabled = False
        Me.optDate.Enabled = False
        Me.OptDaysAfterPlant.Enabled = False
        Me.OptDaysOfYear.Enabled = False
        Me.OptPlot.Enabled = False
        Me.OptDontPlot.Enabled = False
        Me.OptStatShow.Enabled = False
        Me.optShowLine.Enabled = False
        Me.optDontShowLine.Enabled = False
        Me.OptSmall.Enabled = False
        Me.OptSmall.Enabled = False
        Me.OptLarge.Enabled = False
        Me.OptGridShow.Enabled = False
        Me.OptGridDontShow.Enabled = False
        cmdDefaults.Enabled = False
        Me.cmdCancel.Enabled = False
        Me.ChkXAxis.Enabled = False
        Me.ChkEditChart.Enabled = False
        Me.ChkExpShow.Enabled = False
        txtMarkerSize.Enabled = False
    End If
    Unload frmDocument
End Sub



Private Sub OptDataPath_Click()

End Sub

Private Sub OptDontPlot_Click()
    ChkExpShow.Value = 0
End Sub



Private Sub OptGenetic_Click()
    optDate.Value = True
    optDate.Enabled = False
    Me.OptDaysAfterPlant.Enabled = False
    Me.OptDaysOfYear.Enabled = False
End Sub

Private Sub optShowLine_Click()
    If optShowLine.Value = True Then
        chkSimExp.Value = 0
    End If
End Sub


Private Sub OptTime_Click()
    If OptTime.Value = True Then
        Me.optShowLine.Value = True
        chkSimExp.Value = 0
    
      '  OptDaysAfterPlant.Value = True
        optDate.Enabled = True
        Me.OptDaysAfterPlant.Enabled = True
      '  Me.OptDaysOfYear.Enabled = True
    
    End If
End Sub

Private Sub tabRTF_Click()
   fraTab(tabRTF.SelectedItem.Index - 1).ZOrder 0
End Sub



Private Sub tabRTF_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    If Dir(App.path & "\Manual_GBuild.chm") <> "" Then
        App.HelpFile = App.path & "\Manual_GBuild.chm"
        Call HtmlHelp(Me.hWnd, App.HelpFile, HH_DISPLAY_TOC, 0)
    Else
        MsgBox "Cannot find" & " " & App.path & "\Manual_GBuild.chm"
    End If
 End If

End Sub

Private Sub txtFolder_Click()
txtFolder.ForeColor = &H80000012
End Sub

Private Sub txtMarkerSize_Change()
    If IsNumeric(txtMarkerSize.Text) = True Then
        Me.OptSmall.Value = False
        Me.OptLarge.Value = True
    End If
End Sub
