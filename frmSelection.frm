VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelection 
   Caption         =   "Selection"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12405
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   12405
   Tag             =   "1034"
   Begin VB.CheckBox chkSelectAllRuns 
      Caption         =   "Select All Runs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   31
      Top             =   5400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox chkSelectAllTr 
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Cancel          =   -1  'True
      Caption         =   "Clear All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      TabIndex        =   29
      Tag             =   "1135"
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdReloadData 
      Caption         =   "Reload Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4560
      TabIndex        =   22
      Tag             =   "1146"
      Top             =   5280
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3735
      Left            =   240
      TabIndex        =   17
      Top             =   1440
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   6588
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5160
      TabIndex        =   16
      Tag             =   "1003"
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox cmbOutFiles 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   2175
   End
   Begin VB.Data DatTreatment 
      Caption         =   "Data3"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7680
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6000
      TabIndex        =   1
      Tag             =   "1134"
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmdOK1 
      Caption         =   "Next"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6960
      TabIndex        =   0
      Tag             =   "1069"
      Top             =   5280
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   3735
      Left            =   5760
      TabIndex        =   18
      Top             =   1440
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   6588
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   7056
      EndProperty
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   3735
      Left            =   3000
      TabIndex        =   19
      Top             =   1440
      Width           =   30000
      _ExtentX        =   52917
      _ExtentY        =   6588
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   17639
      EndProperty
   End
   Begin VB.Label lblToolOutFiles 
      Caption         =   "Label6"
      Height          =   255
      Left            =   8160
      TabIndex        =   28
      Tag             =   "1021"
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblToolPreview 
      Caption         =   "Label6"
      Height          =   255
      Left            =   8040
      TabIndex        =   27
      Tag             =   "1020"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblToolNext 
      Caption         =   "Label6"
      Height          =   255
      Left            =   7920
      TabIndex        =   26
      Tag             =   "1013"
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblToolClose 
      Caption         =   "Label6"
      Height          =   255
      Left            =   7680
      TabIndex        =   25
      Tag             =   "1008"
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblToolReload 
      Caption         =   "Label6"
      Height          =   255
      Left            =   7560
      TabIndex        =   24
      Tag             =   "1007"
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblToolClear 
      Caption         =   "Label6"
      Height          =   255
      Left            =   7320
      TabIndex        =   23
      Tag             =   "1006"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H8000000A&
      Height          =   5775
      Left            =   0
      Top             =   0
      Width           =   9375
   End
   Begin VB.Label lblYAxis 
      Caption         =   "Y - Axis"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   135
      Left            =   3360
      TabIndex        =   21
      Tag             =   "1115"
      Top             =   1245
      Width           =   1455
   End
   Begin VB.Label lblXAxis 
      Caption         =   "X - Axis"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   135
      Left            =   720
      TabIndex        =   20
      Tag             =   "1114"
      Top             =   1245
      Width           =   1455
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   240
      Left            =   2280
      TabIndex        =   15
      Tag             =   "1096"
      Top             =   480
      Width           =   360
   End
   Begin VB.Label lblTimeSeries 
      Caption         =   "Time Series Plot"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Tag             =   "1095"
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblXYPlotting 
      Caption         =   "Scatter Plot"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Tag             =   "1145"
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblRunDescription 
      Caption         =   "Run Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Tag             =   "1094"
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblVariableName 
      Caption         =   "Variable Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Tag             =   "1093"
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblTreatments 
      Caption         =   "Treatments"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5280
      TabIndex        =   10
      Tag             =   "1090"
      Top             =   960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblRuns 
      Caption         =   "Runs"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5280
      TabIndex        =   9
      Tag             =   "1083"
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblVariables 
      Caption         =   "Variables"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Tag             =   "1068"
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Crop"
      Height          =   495
      Left            =   -480
      TabIndex        =   7
      Tag             =   "1073"
      Top             =   7320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Experiment"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Tag             =   "1072"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Run"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Tag             =   "1062"
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "TRT"
      Height          =   255
      Left            =   10440
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim prmTableOutName As String
Dim NumberArrayItems As Integer
Dim NumberArrayItems2 As Integer
Dim ErrorFound As Boolean
Dim ErrorFile1 As Boolean


Private Sub chkSelectAllRuns_Click()
' VSH
    Dim i As Integer
    Dim item As ListItem
    
    If ((chkSelectAllRuns.Value = 1) And (gAllSelectedRuns = False)) Then
        For Each item In ListView2.ListItems
            item.Selected = True
            item.Checked = True
            PutVariablesInArray2
            item.Selected = False
        Next
        chkSelectAllRuns.Enabled = False
        gAllSelectedRuns = True
    End If
End Sub

Private Sub chkSelectAllTr_Click()
' VSH
    Dim i As Integer
    Dim item As ListItem
    
    If ((chkSelectAllTr.Value = 1) And (gAllSelectedTr = False)) Then
        For Each item In ListView2.ListItems
            item.Selected = True
            item.Checked = True
            PutVariablesInArray2
            item.Selected = False
        Next
        chkSelectAllTr.Enabled = False
        gAllSelectedTr = True
    End If
    
End Sub

Private Sub cmbOutFiles_Click()
    On Error GoTo Error1
    ListView1.ListItems.Clear
    ListView3.ListItems.Clear
        PreviewFile = DirectoryToPreview & cmbOutFiles.Text
        
    
    If FileType = "SUM" Then
        NameSUM_File = cmbOutFiles.Text
        
        prmTableOutName = "Evaluate_List"
        Data1.RecordSource = "Select [Selected], [VariableDescription] From [Evaluate_List]"
        Data1.Refresh
        With Data1.Recordset
            If Not (.EOF And .BOF) Then
                .MoveLast
                NumberArrayItems = .RecordCount
            End If
        End With
        ListView1.Visible = False
        ListView2.Visible = False
        ListView3.Visible = True
    End If
    Shape2.Height = Me.Height - 400
    Shape2.Width = Me.Width - 100
    If FileType = "OUT" Then
        If ExpData_vs_Simulated = 1 Then
            prmTableOutName = Mid(cmbOutFiles.Text, 1, Len(cmbOutFiles.Text) - 4) & "_List_EXPOUT"
        Else
            prmTableOutName = Mid(cmbOutFiles.Text, 1, Len(cmbOutFiles.Text) - 4) & "_List"
        End If
        Data1.RecordSource = "Select [Selected], [VariableDescription] From [" & prmTableOutName & "]"
        Data1.Refresh
        With Data1.Recordset
            If Not (.EOF And .BOF) Then
                .MoveLast
                NumberArrayItems = .RecordCount
            End If
        End With
        Data2.RecordSource = "Select [Selected], [RunNumber], [RunDescription], " _
            & "[ExperimentID], [CropID], [TRNO] " _
            & "From [" & Mid(cmbOutFiles.Text, 1, Len(cmbOutFiles.Text) - 4) _
            & "_File_Info" & "]"
        Data2.Refresh
        With Data2.Recordset
            If Not (.EOF And .BOF) Then
                .MoveLast
                NumberArrayItems2 = .RecordCount
            End If
        End With
    ElseIf FileType = "T-file" Then  'T-File
        prmTableOutName = cmbOutFiles.Text & "_List"
        Data1.RecordSource = "Select [VariableDescription] From [" & _
            prmTableOutName & "]"
        Data1.Refresh
        With Data1.Recordset
            If Not (.EOF And .BOF) Then
                .MoveLast
                NumberArrayItems = .RecordCount
            End If
        End With
        Data2.RecordSource = "Select [Selected], [TRNO] " _
            & "From [" & cmbOutFiles.Text & "_File_Info]"
        Data2.Refresh
        With Data2.Recordset
            If Not (.EOF And .BOF) Then
                .MoveLast
                NumberArrayItems2 = .RecordCount
            End If
        End With
    End If
    
    
    PreviewFile = DirectoryToPreview & Replace(cmbOutFiles.Text, " ", ".")
    Err.Clear
    On Error Resume Next
    ErrorFound = False
    If FileType <> "SUM" Then
        AddList
    End If
    If FileType <> "SUM" Then
        If ErrorFound = False Then
            AddList2
        Else
            Unload Me
        End If
    End If
    If ErrorFound = False Then
        AddList3
    Else
        Unload Me
    End If
    
    
    
    
    cmdPreview.ToolTipText = lblToolPreview.Caption & " " & cmbOutFiles.Text
    If ErrorFound = True Then Unload Me
    Exit Sub
Error1: MsgBox Err.Description & " in cmbOutFiles/frmSelection."
    
End Sub


Private Sub cmdCancel_Click()
        FileIsClosed = True
        CloseISselected = True
       ' ShowRealDates = 0
        ShowX_Axis = 1
        ExpData_vs_Simulated = 0
    Unload Me
    frmDocument.Show
End Sub


Private Sub cmdClear_Click()
Dim i As Integer
    On Error Resume Next
    For i = 1 To NumberOutTables
        dbXbuild.Execute ("Update [" & _
            Mid(OUTTableNames(i), 1, Len(OUTTableNames(i)) - 4) & _
            "_File_Info] Set [Selected] = Null")
        dbXbuild.Execute ("Update [" & _
            Mid(OUTTableNames(i), 1, Len(OUTTableNames(i)) - 4) & _
            "_List] Set [Selected] = Null")
    Next i
   
    For i = 1 To NumberOutTables
        dbXbuild.Execute ("Update [" & _
            OUTTableNames(i) & _
            "_File_Info] Set [Selected] = Null")
        dbXbuild.Execute ("Update [" & _
            OUTTableNames(i) & _
            "_List] Set [Selected] = Null")
    Next i

    If ExpData_vs_Simulated = 1 Then
        For i = 1 To NumberOutTables
            dbXbuild.Execute ("Update [" & _
            Mid(OUTTableNames(i), 1, Len(OUTTableNames(i)) - 4) & _
            "_List_EXPOUT] Set [Selected] = Null")

        Next i
    
    End If

   AddList
   If FileType <> "SUM" Then
    AddList2
   End If
   
   AddList3
   
   ' VSH
   chkSelectAllTr.Value = 0
   chkSelectAllTr.Enabled = True
   gAllSelectedTr = False
   
   chkSelectAllRuns.Value = 0
   chkSelectAllRuns.Enabled = True
   gAllSelectedRuns = False
   
    Exit Sub
Error1: MsgBox Err.Description & " in cmdClear_Click."
End Sub

Private Sub cmdOK1_Click()
Dim i As Integer
Dim n As Integer
Dim nn As Integer
Dim myrsEXPlist As Recordset
Dim myrsEXPinfo As Recordset
Dim myrsGraphData As Recordset
Dim MyrsMyTemp As Recordset
Dim bb As Integer
     Show_Sim = 1
    CurrentFileSelected = cmbOutFiles.Text
    ShowStatistic = 1
    ErrorFound = False
    If FileType <> "SUM" Then
        PutVariablesInArray
    End If
    PutVariablesInArray3
    On Error Resume Next
    dbXbuild.Execute "Drop Table [Graph_Data]"
    dbXbuild.Execute "Drop Table [Exp_Out]"
    Err.Clear
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    ReDim XYVariables(2)
    ReDim XYRuns(2)
    XYTitle = ""
    Screen.MousePointer = vbHourglass
    If XYAxisApproved = False Then
        MsgBox "You must select a variable."
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Y_Variable_Selected = False Then
        MsgBox "You must select a variable."
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    dbXbuild.Execute "Create Table [Graph_Data] ([TheOrder] Text, [Date] single, [RealDate] Date)"
       ' frmWait.Show
       ' DoEvents
       ' frmWait.prgLoad.Visible = True

1:
    If FileType = "OUT" Then
        CreateExpOUTTable
        If ErrorFound = False Then
            DoEvents
        Else
            Exit Sub
        End If
        If ShowX_Axis = 1 Or ExpData_vs_Simulated = 1 Then
            DoEvents
        Else
            AddXvariables
            AddXColumnsToGraph_Data
        End If
2:      AddYColumnsToGraph_Data
        AddExpVariables
        If ErrorFound = False Then
            DoEvents
        Else
            Exit Sub
        End If
        

5:      If ShowRealDates = 1 Then
            frmWait.Show
            DoEvents
            frmWait.prgLoad.Visible = True

            FillOUTGraphData_RealDates1
            If ErrorFound = True Then
                'MsgBox "Selection is not complete."
                MsgBox "Please select Run(s)."
                Exit Sub
            End If
            
            
            
            If ShowX_Axis = 1 Or ExpData_vs_Simulated = 1 Then
                DoEvents
            Else
                FillOUT_X_GraphData_RealDates1
            End If
            ChangeDateFormat
            If FileType <> "SUM" Then
              MakePlotData1
            End If
            
            If prmShowLine = 1 Then
                If ExpData_vs_Simulated = 0 Then
                    FillNullDataInOut1
                End If
            End If

            AddExpDataToGraphData_YearDate1
            If ErrorFound = False Then
                DoEvents
            Else
                Exit Sub
            End If
            ChangeDateFormat
        ElseIf ShowRealDates = 0 Or ShowRealDates = 2 Then 'Show Days after planting
            frmWait.Show
            DoEvents
            frmWait.prgLoad.Visible = True

            FillOUTGraphData_AfterPlanting
            If ErrorFound = True Then
                'MsgBox "Selection is not complete."
                MsgBox "Please select Run(s)."
                Exit Sub
            End If
            frmWait.Show
            DoEvents
            frmWait.prgLoad.Visible = True
            
            If ShowX_Axis = 1 Or ExpData_vs_Simulated = 1 Then
                DoEvents
            Else
                FillOUT_X_GraphData_AfterPlanting
                If ErrorFound = False Then
                    DoEvents
                Else
                    Exit Sub
                End If
            End If
            If FileType <> "SUM" Then
              MakePlotData1
            End If
            If prmShowLine = 1 Then
                If ExpData_vs_Simulated = 0 Then
                    FillNullDataInOut1
                End If
            End If

            AddExpDataToGraphData_AfterPlanting
                If ErrorFound = False Then
                    DoEvents
                Else
                    Exit Sub
                End If
        End If 'Show Days after planting
        
        Set myrsGraphData = dbXbuild.OpenRecordset("Graph_Data")
        If myrsGraphData.EOF And myrsGraphData.BOF Then
            myrsGraphData.Close
            Unload frmWait
            MsgBox "Please select a variable/Run for plotting."
            dbXbuild.Execute "Drop Table [Graph_Data]"
            dbXbuild.Execute "Drop Table [Exp_Out]"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    If FileType = "T-file" Then
        FindSelectedOutTables
        NumberExpVariables_X = 0
        If ShowX_Axis = 1 Then
            DoEvents
        Else
            AddXColumns_GraphData_T_File
            ChangeDateFormat
        End If
        AddYColumns_GraphData_T_File
        If ErrorFound = True Then
            'MsgBox "Selection is not complete."
            MsgBox "Please select Treatment(s)."
            Exit Sub
        End If
        ChangeDateFormat
    End If
    
    If FileType = "SUM" Then
        NumberExpVariables_X = 0
        AddYColumns_GraphData_SUM_File
        If ErrorFound = True Then
            MsgBox "Selection is not complete."
            Exit Sub
        End If
    End If
    
    
    If ShowX_Axis <> 1 Then
        ShowStatistic = 0
    End If
    If ExpData_vs_Simulated = 1 Then
        ShowStatistic = 1
    End If
    If FileType = "T-file" Then
        ShowStatistic = 0
    End If
    If FileType = "SUM" Then
       ShowStatistic = 1
    End If
    Close
    '***
    If AnyDataExists = False Then
        Screen.MousePointer = vbDefault
       frmWait.Visible = False
         Call MsgBox("The file does not have any data for plotting.", vbOKOnly + vbCritical)
        
        Exit Sub
    End If
   '
    MakeDataOrder
    If ShowX_Axis <> 1 Then
        If FileType <> "SUM" Then
            FillNullDataInOut0
        End If
    End If
    FindExcelData
    
    If FileType = "OUT" And ShowX_Axis = 1 Then
        FindExcelDataOUT_Time

    End If
    If FileType = "T-file" Then
       ' FindExcelDataOUT_Time

    End If
    
    Screen.MousePointer = vbDefault
    Err.Clear
    On Error Resume Next
    myrsGraphData.Close
    MyrsMyTemp.Close
    myrsEXPlist.Close
    myrsEXPinfo.Close
    Err.Clear
    On Error GoTo Error1
   
  
    Dim OUT_Field_Name()
    Dim NumberOUTLabal_to_delete As Integer
    Dim j As Integer
    If Show_Sim = 0 Then
        Set MyrsMyTemp = dbXbuild.OpenRecordset("Exp_OUT")
        With MyrsMyTemp
            If Not (.EOF And .BOF) Then
                .MoveLast
                NumberOUTLabal_to_delete = .RecordCount
            End If
            .Close
        End With
        ReDim OUT_Field_Name(NumberOUTLabal_to_delete + 1)
        Set myrsGraphData = dbXbuild.OpenRecordset("Graph_Data")
        With myrsGraphData
            For j = 1 To NumberOUTLabal_to_delete
                OUT_Field_Name(j) = .Fields(j + 2).Name
            Next j
            .Close
        End With
        For j = 1 To NumberOUTLabal_to_delete
            dbXbuild.Execute "Alter Table [Graph_Data] Drop Column [" & _
                    OUT_Field_Name(j) & "]"
        Next j
    End If
    frmGraph.Show
    
    
    Unload Me
    Exit Sub
Error1:     MsgBox Err.Description & " in cmdOK/frmSelection(1)."
            Screen.MousePointer = vbDefault
            Exit Sub
Error2:     MsgBox Err.Description & " in cmdOK/frmSelection(2)."
            Screen.MousePointer = vbDefault
            Exit Sub
Error3:     MsgBox Err.Description & " in cmdOK/frmSelection(3)."
            Screen.MousePointer = vbDefault
            Exit Sub
Error4:     MsgBox Err.Description & " in cmdOK/frmSelection(4)."
            Screen.MousePointer = vbDefault
            Exit Sub
Error5:     MsgBox Err.Description & " in cmdOK/frmSelection(5)."
            Screen.MousePointer = vbDefault
            Exit Sub
Error6:     MsgBox Err.Description & " in cmdOK/frmSelection(6)."
            Screen.MousePointer = vbDefault
            Exit Sub
Error7:     MsgBox Err.Description & " in cmdOK/frmSelection(7)."
            Screen.MousePointer = vbDefault
            Exit Sub
Error8:     MsgBox Err.Description & " in cmdOK/frmSelection(8)."
            Exit Sub
Error10:     MsgBox Err.Description & " in cmdOK/frmSelection(10)."
            Exit Sub
Error7a:   MsgBox Err.Description & " in cmdOK/frmSelection(71)."
            Exit Sub

Error9:     MsgBox Err.Description & " in cmdOK/frmSelection(9)."
            Screen.MousePointer = vbDefault

End Sub


Private Sub cmdPreview_Click()
    frmFilePreview.Show
End Sub


Private Sub cmdReloadData_Click()
    Unload Me
    ReloadData
    Me.Show
End Sub

Private Sub Form_Activate()
' VSH
If chkSelectAllTr.Visible And gAllSelectedTr Then
    chkSelectAllTr.Enabled = False
    chkSelectAllTr.Value = 1
End If

If chkSelectAllRuns.Visible And gAllSelectedRuns Then
    chkSelectAllRuns.Enabled = False
    chkSelectAllRuns.Value = 1
End If

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
    Case 12
        cmdCancel_Click
    Case 14
        cmdOK1_Click
    Case 1
        cmdClear_Click
    Case 18
        cmdReloadData_Click
    Case 22
        cmdPreview_Click
    End Select
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim prmTheOutFileName As String
Dim prmMyExperiment As String
Dim col1 As Column
    Me.Top = 0
    Me.Left = 0
    FileIsClosed = False
    Me.Width = 9450
    Me.Height = 6400
    If FileType = "SUM" Then
        ListView3.Width = 2775 * 2
    Else
        ListView3.Width = 2775
        
    End If
   
    If LongFile = True And ShowRealDates = 1 Then
        LongFile = False
        'If ShowRealDates = 1 Then
            Call MsgBox("  The time period for all Treatments is more then 3 years." & Chr(10) & Chr(13) & " " & _
            "Creating the plot may take a few minutes." & Chr(10) & Chr(13) & " " & _
            "It is recomended to change the option to 'Days after Planting'.", vbCritical, "Warning")

        'End If
    End If
    
    Me.Caption = FrmSelectionName
    If FileType = "SUM" Then
        lblRuns.Visible = False
    
    End If
    ExpExists = False
    Screen.MousePointer = vbHourglass
    If FileType = "T-file" Then
        ShowX_Axis = 1
    End If
    NumberExpVariables_Y = 0
    
    If ShowX_Axis = 1 Or ExpData_vs_Simulated = 1 Or FileType = "SUM" Then
        ListView1.Visible = False
        ListView3.Left = 1958
        lblVariables.Top = 1080
        lblRuns.Top = 1080
        lblRuns.Left = 5280
    Else
        ListView1.Visible = True
        'ListView1.Left = 938
        ListView1.Left = 538
        lblVariables.Top = 840
        lblRuns.Top = 840
        lblRuns.Left = 6200
        
        ' VSH Scatter plot
        chkSelectAllRuns.Top = lblRuns.Top
        chkSelectAllRuns.Left = lblRuns.Left + lblRuns.Width + 400
        chkSelectAllRuns.Visible = True
        
    End If
    
    LoadResStrings Me
    
   ' If gsglXFactor <> 1 Or gsglYFactor <> 1 Then
        Call SetDeviceIndependentWindow(Me, 1)
   ' End If
    Dim mm, gg
    
    Me.lblRuns.FontSize = 12
    Me.lblTreatments.FontSize = 12
    Me.lblVariables.FontSize = 12
    Me.lblXYPlotting.FontSize = 14
     Me.lblTimeSeries.FontSize = 14
     
     
    If ExpData_vs_Simulated = 1 And FileType <> "SUM" Then
        ErrorFile1 = False
        If KeepSim_vs_Obs = False Then
            Create_New_List_Table
        End If
        If ErrorFile1 = True Then
            ErrorFile1 = False
            Unload Me
            Screen.MousePointer = vbDefault
            Unload frmWait2
            Exit Sub
        End If
    End If
    If OneTimeDAP_Exists = 1 And FileType = "OUT" And NumberOutTables > 1 Then
        Add_DAP_Collumn_To_OUT
        
    End If
    If ShowX_Axis = 1 Or ExpData_vs_Simulated = 1 Or FileType = "SUM" Then
        lblTimeSeries.Visible = True
        ListView2.Left = ListView3.Left + _
        ListView3.Width
        lblXAxis.Visible = False
        lblYAxis.Visible = False
        
        ' VSH Scatter plot
        chkSelectAllRuns.Top = lblRuns.Top
        chkSelectAllRuns.Left = lblRuns.Left + lblRuns.Width + 100
        chkSelectAllRuns.Visible = True
    ElseIf FileType <> "SUM" Then
        lblXYPlotting.Visible = True
        ListView3.Left = ListView1.Left + _
        ListView1.Width
        ListView2.Left = ListView3.Left + _
        ListView3.Width
        lblXAxis.Visible = True
        lblYAxis.Visible = True
    End If
    If ExpData_vs_Simulated = 1 Or FileType = "SUM" Then
        lblTimeSeries.Visible = False
        lblXYPlotting.Visible = True
    End If
    frmSelectionIsUnloaded = False
    Data1.DatabaseName = ApplicationPathOption + "\GraphCreateTEMP.mdb"
    Data2.DatabaseName = ApplicationPathOption + "\GraphCreateTEMP.mdb"
    Me.Left = 0
    Me.Top = 0
    cmbOutFiles.Clear
    For i = 1 To NumberOutTables
        cmbOutFiles.AddItem OUTTableNames(i)
    Next i
    
    If NumberOutTables > 0 Then
        cmbOutFiles.ListIndex = 0
    End If
    
    If ErrorFound = True Then
        Unload Me
        Exit Sub
    End If
    ''''!!!Selected Output File
    If FileType = "OUT" And FileType <> "SUM" Then
        lblTreatments.Visible = False
    ElseIf FileType = "T-file" Then
        Me.lblTreatments.Visible = True
        
        ' VSH
        chkSelectAllRuns.Visible = False
        lblTreatments.Top = lblVariables.Top
        lblTreatments.Left = ListView2.Left + 300
        chkSelectAllTr.Top = lblTreatments.Top
        chkSelectAllTr.Left = ListView2.Left + ListView2.Width - chkSelectAllTr.Width / 2
        chkSelectAllTr.Enabled = True
        chkSelectAllTr.Visible = True
        
        Me.lblRuns.Visible = False
    End If
    
    
    With cmdCancel
        '.Top = Me.Height - (2 + 1 / 4) * .Height
        .Top = Me.ListView1.Top + Me.ListView1.Height + (1 / 8) * .Height
    End With
    With cmdOK1
        '.Top = Me.Height - (2 + 1 / 4) * .Height
        .Top = Me.ListView1.Top + Me.ListView1.Height + (1 / 8) * .Height
    End With
    With cmdClear
        '.Top = Me.Height - (2 + 1 / 4) * .Height
        .Top = Me.ListView1.Top + Me.ListView1.Height + (1 / 8) * .Height
    End With
    Me.cmdPreview.Height = Me.cmbOutFiles.Height
    cmdReloadData.Height = Me.cmdClear.Height
    cmdReloadData.Top = cmdClear.Top
    cmbOutFiles.Top = cmdPreview.Top
    cmdPreview.Left = cmbOutFiles.Left + (1 + 1 / 8) * cmbOutFiles.Width
    
    'Label7.Top = cmbOutFiles.Top
    ' VSH
    Label7.Top = cmbOutFiles.Top + (cmbOutFiles.Height - Label7.Height) / 2
    Label7.Left = cmbOutFiles.Left - Label7.Width - 100
    
    If CurrentFileSelected <> "" Then
        cmbOutFiles.Text = CurrentFileSelected
    End If
    cmdClear.ToolTipText = lblToolClear.Caption
    cmdReloadData.ToolTipText = lblToolReload.Caption
    cmdCancel.ToolTipText = lblToolClose.Caption
    cmdOK1.ToolTipText = lblToolNext.Caption
    cmbOutFiles.ToolTipText = lblToolOutFiles.Caption
    
    cmbOutFiles_Click
    
    Unload frmWait2
    Unload frmWait3
    Unload frmWait
    Screen.MousePointer = vbDefault
    Unload frmDocument
    Exit Sub
Error1:
    FileIsClosed = True
    Unload frmWait2
    Screen.MousePointer = vbDefault
    MsgBox Err.Description & " in frmSelection_Load."
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmSelectionIsUnloaded = True
End Sub



Public Function FindExpVariable(prmMyCropID As String, prmMyExpermentID As String, prmMyVariableID As String) As String
Dim i As Integer
Dim myrs As Recordset
    On Error GoTo Error1
    Set myrs = dbXbuild.OpenRecordset("Select * From [" & prmMyExpermentID & " " & prmMyCropID & "T]")
    With myrs
        For i = 1 To myrs.Fields.Count - 1
            If prmMyVariableID = .Fields(i).Name Then
                FindExpVariable = "YES"
                Exit Function
            End If
        Next i
        .Close
    End With
    Exit Function
Error1:
    FindExpVariable = "NO"
End Function

Public Sub MakeDataOrder()
Dim myrs As Recordset
Dim i As Integer
Dim prmNumberOfFields As Integer
Dim EmptyField() As String
Dim prmFieldsName() As String
Dim NumberOfRecordsDataGraph As Integer
Dim NumberNullRecords As Integer
Dim MyVariable As String
Dim myFile As String
Dim myFileExtention As String
Dim myTRT As Integer
Dim myFileLong As String
Dim k As Integer
On Error GoTo Error1
    Set myrs = dbXbuild.OpenRecordset("Graph_Data")
    If ShowRealDates = 0 Then
        With myrs
            If Not (.EOF And .BOF) Then
                .MoveLast
                NumberOfRecordsDataGraph = .RecordCount
            End If
        End With
        With myrs
            prmNumberOfFields = .Fields.Count - 1
            ReDim prmFieldsName(1 To prmNumberOfFields)
            For i = 3 To prmNumberOfFields
                prmFieldsName(i) = .Fields(i).Name
            Next i
            .Close
        End With
        For i = 3 To prmNumberOfFields
            Set myrs = dbXbuild.OpenRecordset("Select * from Graph_Data Where [" & _
                prmFieldsName(i) & "] Is Null")
            With myrs
                If Not (.EOF And .BOF) Then
                    .MoveLast
                    NumberNullRecords = .RecordCount
                End If
                .Close
            End With
            ReDim Preserve EmptyField(i + 1)
            If NumberOfRecordsDataGraph = NumberNullRecords Then
                EmptyField(i) = prmFieldsName(i)
            Else
                EmptyField(i) = ""
            End If
        Next i
        Close
        Err.Clear
        On Error Resume Next
        For i = 3 To prmNumberOfFields
            If EmptyField(i) <> "" Then
                dbXbuild.Execute "Alter Table [Graph_Data] Drop Column [" & _
                    EmptyField(i) & "]"
                If InStr(EmptyField(i), "TRT") <> 0 Then
                    MyVariable = Trim(Mid(EmptyField(i), 1, InStr(EmptyField(i), "(") - 1))
                    myFileLong = Trim(Mid(EmptyField(i), Len(MyVariable) + 2, (InStr(EmptyField(i), ")") - _
                        Len(MyVariable) - 3)))
                    myFile = Trim(Mid(myFileLong, 1, InStr(myFileLong, " ")))
                    myFileExtention = Trim(Mid(myFileLong, InStr(myFileLong, " ") + 1))
                    myTRT = Val(Mid(EmptyField(i), InStrRev(EmptyField(i), " ")))
                    Set myrs = dbXbuild.OpenRecordset("Select * From [Exp_Out] Where " & _
                        "[Variable] = '" & MyVariable & "' And [ExperimentID] = '" & _
                        myFile & "' And [CropID] = '" & myFileExtention & "' And [TRNO] = " & _
                        myTRT)
                    With myrs
                        If Not (.EOF And .BOF) Then
                            .MoveFirst
                            .Edit
                            !ExpCheck = "NO"
                            .Update
                        End If
                        .Close
                    End With
                End If
            End If
        Next i
    End If
    Err.Clear
    On Error Resume Next
    myrs.Close
    Exit Sub
Error1:
    MsgBox Err.Description & " in MakeDataOrder"
    Unload frmWait2
    Unload frmWait3
    Unload frmWait
    Screen.MousePointer = vbDefault
End Sub


Public Function XYAxisApproved() As Boolean
Dim myrsOUT As Recordset
Dim NotEmpty As Integer
Dim i As Integer
Dim TheOutFile As String
    On Error GoTo Error1
    NotEmpty = 0
    If FileType = "SUM" Then
        XYAxisApproved = True
        Exit Function
    End If
    
    If ShowX_Axis = 1 Or ExpData_vs_Simulated = 1 Then
        XYAxisApproved = True
    Else
        NotEmpty = 0
        For i = 1 To NumberOutTables
            If FileType = "OUT" Then
                TheOutFile = Mid(OUTTableNames(i), 1, Len(OUTTableNames(i)) - 4)
            Else 'T-File
                TheOutFile = OUTTableNames(i)
            End If
            Set myrsOUT = dbXbuild.OpenRecordset("Select * From [" & TheOutFile _
                & "_List] Where Selected = 'X' Or Selected = 'XY'")
            With myrsOUT
                If Not (.EOF And .BOF) Then
                   NotEmpty = 1 + NotEmpty
                End If
            End With
        Next i
        If NotEmpty <> 1 Then
            XYAxisApproved = False
            myrsOUT.Close
            Exit Function
        Else
            XYAxisApproved = True
        End If
        Err.Clear
        On Error Resume Next
        myrsOUT.Close
    End If
    Err.Clear
    On Error Resume Next
    myrsOUT.Close
    Exit Function
Error1:
    MsgBox Err.Description & " in XYAxisApproved/frmSelection."
End Function



Public Sub AddList()
Dim ArraySelectedVariables() As String
Dim Selection() As String
Dim myrs As Recordset
Dim itmX As ListItem
Dim i As Integer
Dim mm1 As ListItem
   'MsgBox "Adding variables to the list."
   On Error GoTo Error1
    'Fill info for OUT file Variables
        ListView1.ListItems.Clear
        ListView1.View = lvwReport
    If FileType = "OUT" Then
        'Set myrs = dbXbuild.OpenRecordset("Select [Selected], [VariableDescription] From [" & prmTableOutName & "]")
        Set myrs = dbXbuild.OpenRecordset("Select * From [" & prmTableOutName & "]")
    
    ElseIf FileType = "T-file" Then
       Set myrs = dbXbuild.OpenRecordset("Select [Selected], [VariableDescription] From [" & _
            cmbOutFiles.Text & "_List]")
        myrs.MoveFirst
        While Not myrs.EOF
            Set itmX = ListView3.ListItems. _
                Add(, , CStr(myrs!VariableDescription))
            
            myrs.MoveNext   ' Move to next record.
        Wend
        With myrs
            ReDim ArraySelectedVariables(1)
            ReDim Selection(1)
            .MoveFirst
            i = 1
            Do While Not (.EOF)
                ReDim Preserve ArraySelectedVariables(i + 1)
                ReDim Preserve Selection(i + 1)
                ArraySelectedVariables(i) = !VariableDescription
                Selection(i) = IIf(IsNull(!Selected) = True, "", !Selected)
                i = i + 1
                .MoveNext
            Loop
        End With
        NumberArrayItems = i - 1
        For i = 1 To NumberArrayItems
            Set mm1 = ListView3.FindItem(ArraySelectedVariables(i))
            If Selection(i) = "X" Then
                mm1.Checked = True
            ElseIf Selection(i) = "XY" Then
                mm1.Checked = True
            Else
                mm1.Checked = False
            End If
        Next i
    ElseIf FileType = "SUM" Then
       Set myrs = dbXbuild.OpenRecordset("Select [Selected], [VariableDescription] From [Evaluate_List]")
    End If
        myrs.MoveFirst
        
        
        
        While Not myrs.EOF
            Set itmX = ListView1.ListItems. _
                Add(, , CStr(myrs!VariableDescription))
           ' MsgBox CStr(myrs!VariableDescription)
            If FileType = "OUT" And ExpData_vs_Simulated = 0 Then
                itmX.Bold = True_ExpExist(prmTableOutName, myrs!VariableID)
            End If
            If ExpData_vs_Simulated = 1 Then
                ExpExists = True
                itmX.Bold = True
            End If
            myrs.MoveNext   ' Move to next record.
        Wend
        With myrs
            ReDim ArraySelectedVariables(1)
            ReDim Selection(1)
            .MoveFirst
            i = 1
            Do While Not (.EOF)
                ReDim Preserve ArraySelectedVariables(i + 1)
                ReDim Preserve Selection(i + 1)
                ArraySelectedVariables(i) = !VariableDescription
                Selection(i) = IIf(IsNull(!Selected) = True, "", !Selected)
                i = i + 1
                .MoveNext
            Loop
        End With
        NumberArrayItems = i - 1
        For i = 1 To NumberArrayItems
            Set mm1 = ListView1.FindItem(ArraySelectedVariables(i))
            If Selection(i) = "X" Then
                mm1.Checked = True
            ElseIf Selection(i) = "XY" Then
                mm1.Checked = True
            Else
                mm1.Checked = False
            End If
        Next i
    Err.Clear
    On Error Resume Next
    myrs.Close
Exit Sub
Error1:
Screen.MousePointer = vbDefault
MsgBox Err.Description & " in AddList/frmSimulation."
'MsgBox "Error with file."
ErrorFound = True
    Unload frmWait2
    Unload frmWait3
    Unload frmWait
    Screen.MousePointer = vbDefault
End Sub





Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
   ' If KeyCode = 32 And Shift = 0 Then
        ListView1_Click
   ' End If
End Sub

Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 And Shift = 0 Then
        If ListView2.SelectedItem.Checked = False Then
            ListView2.SelectedItem.Checked = True
        Else
            ListView2.SelectedItem.Checked = False
        End If
        PutVariablesInArray2
    End If
End Sub


Private Sub ListView2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        If ListView2.SelectedItem.Checked = False Then
            ListView2.SelectedItem.Checked = True
        Else
            ListView2.SelectedItem.Checked = False
        End If
        PutVariablesInArray2
    End If
End Sub

Private Sub ListView3_KeyDown(KeyCode As Integer, Shift As Integer)
  
 ' If KeyCode = 32 And Shift = 0 Then
        ListView3_Click
 '   End If
End Sub

Private Sub ListView3_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub ListView1_Click()
Dim i As Integer
Dim myrs As Recordset
Dim mm1 As ListItem
    On Error GoTo Error1
    ListView1.FullRowSelect = False
    For i = 1 To NumberArrayItems
        If FileType = "OUT" Then
            Set myrs = dbXbuild.OpenRecordset("Select * From " & "[" & _
                prmTableOutName & "] Where [Selected] = 'X' Or [Selected] = 'XY'")
        ElseIf FileType = "T-file" Then
            Set myrs = dbXbuild.OpenRecordset("Select * From " & "[" & _
                cmbOutFiles.Text & "_List] Where [Selected] = 'X' Or [Selected] = 'XY'")
        End If
        With myrs
            If Not (.EOF And .BOF) Then
                .MoveFirst
                .Edit
                Set mm1 = ListView1.FindItem(!VariableDescription)
                !Selected = Null
                mm1.Checked = False
                .Update
            End If
       End With
    Next i
    myrs.Close
    If FileType = "OUT" Then
        For i = 1 To NumberOutTables
            dbXbuild.Execute "Update [" & Mid(OUTTableNames(i), 1, Len(OUTTableNames(i)) - 4) & _
                "_List" & "] Set [Selected] = Null Where [Selected] = 'X'"
            dbXbuild.Execute "Update [" & Mid(OUTTableNames(i), 1, Len(OUTTableNames(i)) - 4) & _
                "_List" & "] Set [Selected] = null Where [Selected] = 'XY'"
        Next i
    ElseIf FileType = "T-file" Then
        For i = 1 To NumberOutTables
            dbXbuild.Execute "Update [" & OUTTableNames(i) & "_List] Set [Selected] = Null Where [Selected] = 'X'"
            dbXbuild.Execute "Update [" & OUTTableNames(i) & "_List] Set [Selected] = null Where [Selected] = 'XY'"
        Next i
    End If
    Err.Clear
    On Error Resume Next
    myrs.Close
    PutVariablesInArray
    Exit Sub
Error1:
    Unload frmWait2
    Unload frmWait3
    Unload frmWait
    Screen.MousePointer = vbDefault
MsgBox Err.Description & " in ListView1."

End Sub


Private Sub ListView3_Click()
Dim i As Integer
Dim myrs As Recordset
Dim mm1 As ListItem
    On Error GoTo Error1
    ListView3.FullRowSelect = True
        If ExpData_vs_Simulated = 1 Then
            ListView3.MultiSelect = False
            For i = 1 To NumberOutTables
                Set myrs = dbXbuild.OpenRecordset("Select * From " & "[" & _
                prmTableOutName & "] Where [Selected] = 'Y' Or [Selected] = 'XY'")
                With myrs
                    If Not (.EOF And .BOF) Then
                        .MoveFirst
                        .Edit
                        Set mm1 = ListView3.FindItem(!VariableDescription)
                        !Selected = Null
                        mm1.Checked = False
                        mm1.Selected = False
                        ListView3.Refresh
                        .Update
                    End If
                    .Close
                End With
            
            Next i
        Err.Clear
        On Error Resume Next
            For i = 1 To NumberOutTables
                dbXbuild.Execute "Update [" & Mid(OUTTableNames(i), 1, Len(OUTTableNames(i)) - 4) & _
                "_List_EXPOUT" & "] Set [Selected] = Null Where [Selected] = 'Y'"
                dbXbuild.Execute "Update [" & Mid(OUTTableNames(i), 1, Len(OUTTableNames(i)) - 4) & _
                "_List_EXPOUT" & "] Set [Selected] = Null Where [Selected] = 'XY'"
            Next i
        End If
        Err.Clear
        On Error Resume Next
        myrs.Close
    PutVariablesInArray3
    Exit Sub
Error1: MsgBox Err.Description & " in ListView3/frmSelection."
End Sub

Public Sub PutVariablesInArray()
Dim ArraySelectedVariables() As String
Dim i As Integer
Dim mm
Dim myrs As Recordset
Dim mm1 As ListItem
    On Error GoTo Error1
        With Data1.Recordset
            ReDim ArraySelectedVariables(1 To NumberArrayItems + 1)
            .MoveFirst
            i = 1
            Do While Not (.EOF)
                ArraySelectedVariables(i) = !VariableDescription
                i = i + 1
                .MoveNext
            Loop
        End With
        
        For i = 1 To NumberArrayItems
            Set mm1 = ListView1.FindItem(ArraySelectedVariables(i))
            If FileType = "OUT" Then
                Set myrs = dbXbuild.OpenRecordset("Select [Selected], [VariableDescription] " & _
                    "From [" & prmTableOutName & "] Where [VariableDescription] = " & _
                    "'" & ArraySelectedVariables(i) & "'")
            ElseIf FileType = "T-file" Then
                Set myrs = dbXbuild.OpenRecordset("Select [Selected], [VariableDescription] " & _
                    "From [" & cmbOutFiles.Text & "_List] Where [VariableDescription] = " & _
                    "'" & ArraySelectedVariables(i) & "'")
            End If
            With myrs
                .Edit
                If mm1.Checked = True Then
                    If !Selected = "Y" Then
                        !Selected = "XY"
                    ElseIf !Selected = "XY" Then
                        !Selected = "XY"
                    Else
                        !Selected = "X"
                    End If
                Else
                    If !Selected = "XY" Then
                        !Selected = "Y"
                    ElseIf !Selected = "X" Then
                        !Selected = Null
                    End If
                End If
                .Update
                .Close
            End With
        Next i
    Err.Clear
    On Error Resume Next
    myrs.Close
    Exit Sub
Error1: MsgBox Err.Description & " in PutVariablesInArray/frmSelection."
    Unload frmWait2
    Unload frmWait3
    Unload frmWait
    Screen.MousePointer = vbDefault
End Sub

Public Sub AddList2()
Dim ArraySelectedVariables() As Integer
Dim Selection() As String
Dim myrs As Recordset
Dim itmX As ListItem
Dim i As Integer
Dim mm1 As ListItem
    ListView2.View = lvwReport
    ListView2.ListItems.Clear
    On Error GoTo Error1
    If FileType = "OUT" Then
        Set myrs = dbXbuild.OpenRecordset("Select [Selected], [RunNumber], [RunDescription], " _
            & "[ExperimentID], [CropID], [TRNO] " _
            & "From [" & Mid(cmbOutFiles.Text, 1, Len(cmbOutFiles.Text) - 4) _
            & "_File_Info" & "]")
        myrs.MoveFirst
        While Not myrs.EOF
            Set itmX = ListView2.ListItems. _
                Add(, , CStr(myrs!RunNumber))
            itmX.SubItems(1) = CStr(myrs!RunDescription)
            myrs.MoveNext   ' Move to next record.
        Wend
    
        With myrs
            ReDim ArraySelectedVariables(1)
            ReDim Selection(1)
            .MoveFirst
            i = 1
            Do While Not (.EOF)
                ReDim Preserve ArraySelectedVariables(i + 1)
                ReDim Preserve Selection(i + 1)
                ArraySelectedVariables(i) = !RunNumber
                Selection(i) = IIf(IsNull(!Selected) = True, "", !Selected)
                i = i + 1
                .MoveNext
            Loop
        End With
        NumberArrayItems2 = i - 1
        For i = 1 To NumberArrayItems2
            Set mm1 = ListView2.FindItem(ArraySelectedVariables(i))
            If Selection(i) = "Y" Then
                mm1.Checked = True
            ElseIf Selection(i) = "XY" Then
                mm1.Checked = True
            Else
                mm1.Checked = False
            End If
        Next i
    Else
        Set myrs = dbXbuild.OpenRecordset("Select [Selected], [TRNO], [Description] " _
            & "From [" & cmbOutFiles.Text & "_File_Info]")
        myrs.MoveFirst
        While Not myrs.EOF
            Set itmX = ListView2.ListItems. _
                Add(, , CStr(myrs!TRNO))
                itmX.SubItems(1) = CStr(myrs!Description)
            myrs.MoveNext   ' Move to next record.
        Wend
    
        With myrs
            ReDim ArraySelectedVariables(1)
            ReDim Selection(1)
            .MoveFirst
            i = 1
            Do While Not (.EOF)
                ReDim Preserve ArraySelectedVariables(i + 1)
                ReDim Preserve Selection(i + 1)
                ArraySelectedVariables(i) = !TRNO
                Selection(i) = IIf(IsNull(!Selected) = True, "", !Selected)
                i = i + 1
                .MoveNext
            Loop
        End With
         
         NumberArrayItems2 = i - 1
        For i = 1 To NumberArrayItems2
            Set mm1 = ListView2.FindItem(ArraySelectedVariables(i))
            If Selection(i) = "Y" Then
                mm1.Checked = True
            ElseIf Selection(i) = "XY" Then
                mm1.Checked = True
            Else
                mm1.Checked = False
            End If
        Next i
    
    End If
    Err.Clear
    On Error Resume Next
    myrs.Close
Exit Sub
Error1:
    Unload frmWait2
    Unload frmWait3
    Unload frmWait
    Screen.MousePointer = vbDefault
MsgBox Err.Description & " in AddList2/frmSelection."
'MsgBox "Error with file."
ErrorFound = True

End Sub



Private Sub ListView3_ItemClick(ByVal item As MSComctlLib.ListItem)
    item.Selected = False
End Sub

Private Sub ListView2_ItemClick(ByVal item As MSComctlLib.ListItem)
    item.Selected = False
End Sub


Private Sub ListView1_ItemClick(ByVal item As MSComctlLib.ListItem)
    item.Selected = False
End Sub

Private Sub ListView2_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub ListView2_Click()
   ' ListView2.FullRowSelect = True
   ' ListView2.SelectedItem.Checked = True
    PutVariablesInArray2
    
    ' VSH
    ' VSH
    Dim i As Integer
    Dim item As ListItem
        
    i = 0
    For Each item In ListView2.ListItems
        If item.Checked Then
          i = i + 1
        End If
    Next
    
    
    Select Case i
    Case ListView2.ListItems.Count
        chkSelectAllTr.Value = 1
        chkSelectAllTr.Enabled = False
        gAllSelectedTr = True
        
        chkSelectAllRuns.Value = 1
        chkSelectAllRuns.Enabled = False
        gAllSelectedRuns = True
    Case Else
        chkSelectAllTr.Value = 0
        chkSelectAllTr.Enabled = True
        gAllSelectedTr = False
        
        chkSelectAllRuns.Value = 0
        chkSelectAllRuns.Enabled = True
        gAllSelectedRuns = False
    End Select
    
        
End Sub



Public Sub PutVariablesInArray2()
Dim ArraySelectedVariables() As Integer
Dim i As Integer
Dim mm
Dim myrs As Recordset
Dim mm1 As ListItem
Dim SelectedRuns() As Integer
Dim NumberSelectedRuns As Integer
Dim j As Integer
Dim TheOutFile As String
        On Error GoTo Error1
        With Data2.Recordset
            ReDim ArraySelectedVariables(1 To NumberArrayItems2 + 1)
            .MoveFirst
            i = 1
            Do While Not (.EOF)
                If FileType = "OUT" Then
                    ArraySelectedVariables(i) = !RunNumber
                Else
                    ArraySelectedVariables(i) = !TRNO
                End If
                i = i + 1
                .MoveNext
            Loop
        End With
        For i = 1 To NumberArrayItems2
            Set mm1 = ListView2.FindItem(ArraySelectedVariables(i))
            If FileType = "OUT" Then
                Set myrs = dbXbuild.OpenRecordset("Select [Selected], [RunNumber], [RunDescription], " _
                & "[ExperimentID], [CropID], [TRNO] " _
                & "From [" & Mid(cmbOutFiles.Text, 1, Len(cmbOutFiles.Text) - 4) _
                & "_File_Info" & "] Where [RunNumber] = " & ArraySelectedVariables(i))
            Else
                Set myrs = dbXbuild.OpenRecordset("Select [Selected], " _
                & "[TRNO] " _
                & "From [" & cmbOutFiles.Text _
                & "_File_Info" & "] Where [TRNO] = " & ArraySelectedVariables(i))
            End If
            
            With myrs
                .Edit
                If mm1.Checked = True Then
                    If !Selected = "X" Then
                        !Selected = "XY"
                    ElseIf !Selected = "XY" Then
                        !Selected = "XY"
                    Else
                        !Selected = "Y"
                    End If
                Else
                    If !Selected = "X" Then
                        !Selected = "X"
                    ElseIf !Selected = "Y" Then
                        !Selected = Null
                    End If
                End If
                .Update
                .Close
            End With
        Next i
    j = 1
    ReDim SelectedRuns(1)
    If FileType = "OUT" Then
        If NumberOutTables > 0 Then
       'Check the change!!!!!!
       ' If NumberOutTables > 1 Then
            Set myrs = dbXbuild.OpenRecordset("Select * " _
                & "From [" & Mid(cmbOutFiles.Text, 1, Len(cmbOutFiles.Text) - 4) _
                & "_File_Info" & "] Where [Selected] = 'Y'")
            With myrs
                If Not (.EOF And .BOF) Then
                    .MoveFirst
                    Do While Not (.EOF)
                        ReDim Preserve SelectedRuns(j + 1)
                        SelectedRuns(j) = !RunNumber
                        j = j + 1
                    .MoveNext
                    Loop
                End If
            End With
            NumberSelectedRuns = j - 1
            myrs.Close
            For i = 1 To NumberOutTables
                TheOutFile = Mid(OUTTableNames(i), 1, Len(OUTTableNames(i)) - 4)
                dbXbuild.Execute "Update [" & TheOutFile & _
                    "_File_Info] Set " & "[Selected] = Null"
                If NumberSelectedRuns > 0 Then
                    For j = 1 To NumberSelectedRuns
                        dbXbuild.Execute "Update [" & TheOutFile & _
                        "_File_Info] Set " & "[Selected] = 'Y' Where [RunNumber] = " & SelectedRuns(j)
                    Next j
                End If
            Next i
        End If
    Else
        'For t-file
        'Codes go here
    End If
    Err.Clear
    On Error Resume Next
    myrs.Close
Exit Sub
Error1:
    Unload frmWait2
    Unload frmWait3
    Unload frmWait
    Screen.MousePointer = vbDefault
MsgBox Err.Description & " in PutVariablesInArray2/frmSelection."
End Sub

Public Sub AddList3()
Dim ArraySelectedVariables() As String
Dim Selection() As String
Dim myrs As Recordset
Dim itmX As ListItem
Dim i As Integer
Dim mm1 As ListItem
   On Error GoTo Error1
    'Fill info for OUT file Variables
        ListView3.ListItems.Clear
        ListView3.View = lvwReport
    If FileType = "OUT" Then
        Set myrs = dbXbuild.OpenRecordset("Select * From [" & prmTableOutName & "]")
    ElseIf FileType = "T-file" Then
       prmTableOutName = cmbOutFiles.Text & "_List"
       Set myrs = dbXbuild.OpenRecordset("Select [Selected], [VariableDescription] From [" & _
            prmTableOutName & "]")
    ElseIf FileType = "SUM" Then
        
        Set myrs = dbXbuild.OpenRecordset("Select [Selected], [VariableDescription] From [" & prmTableOutName & "]")
        If myrs.EOF And myrs.BOF Then
             Call MsgBox("The file does not have any data for plotting.", vbOKOnly + vbCritical)
            GoTo Error1
        End If
    
    End If
        myrs.MoveFirst
        While Not myrs.EOF
            Set itmX = ListView3.ListItems. _
                Add(, , CStr(myrs!VariableDescription))
            'Debug.Print CStr(myrs!VariableDescription)
            If FileType = "OUT" And ExpData_vs_Simulated = 0 Then
                itmX.Bold = True_ExpExist(prmTableOutName, myrs!VariableID)
            End If
            If ExpData_vs_Simulated = 1 Then
                itmX.Bold = True
            End If
            myrs.MoveNext   ' Move to next record.
        Wend
    
        With myrs
            ReDim ArraySelectedVariables(1)
            ReDim Selection(1)
            .MoveFirst
            i = 1
            Do While Not (.EOF)
                ReDim Preserve ArraySelectedVariables(i + 1)
                ReDim Preserve Selection(i + 1)
                ArraySelectedVariables(i) = !VariableDescription
                Selection(i) = IIf(IsNull(!Selected) = True, "", !Selected)
                i = i + 1
                .MoveNext
            Loop
        End With
        NumberArrayItems = i - 1
        For i = 1 To NumberArrayItems
            Set mm1 = ListView3.FindItem(ArraySelectedVariables(i))
            If Selection(i) = "Y" Then
                mm1.Checked = True
            ElseIf Selection(i) = "XY" Then
                mm1.Checked = True
            Else
                mm1.Checked = False
            End If
        Next i
    Exit Sub
    Err.Clear
    On Error Resume Next
    myrs.Close
Error1:
    Unload frmWait2
    Unload frmWait3
    Unload frmWait
    Screen.MousePointer = vbDefault
'MsgBox Err.Description & " in AddList3."
'MsgBox "Error with file."
ErrorFound = True

End Sub



Public Sub PutVariablesInArray3()
Dim ArraySelectedVariables() As String
Dim i As Integer
Dim mm
Dim myrs As Recordset
Dim mm1 As ListItem
    On Error GoTo Error1
        With Data1.Recordset
            ReDim ArraySelectedVariables(1 To NumberArrayItems + 1)
            .MoveFirst
            i = 1
            Do While Not (.EOF)
                ArraySelectedVariables(i) = !VariableDescription
                i = i + 1
                .MoveNext
            Loop
        End With
        For i = 1 To NumberArrayItems
            Set mm1 = ListView3.FindItem(ArraySelectedVariables(i))
            If FileType = "OUT" Then
                Set myrs = dbXbuild.OpenRecordset("Select [Selected], [VariableDescription] " & _
                    "From [" & prmTableOutName & "] Where [VariableDescription] = " & _
                    "'" & ArraySelectedVariables(i) & "'")
            ElseIf FileType = "T-file" Then
                Set myrs = dbXbuild.OpenRecordset("Select [Selected], [VariableDescription] " & _
                    "From [" & Me.cmbOutFiles.Text & "_List" & "] Where [VariableDescription] = " & _
                    "'" & ArraySelectedVariables(i) & "'")
            ElseIf FileType = "SUM" Then
                Set myrs = dbXbuild.OpenRecordset("Select [Selected], [VariableDescription] " & _
                    "From [Evaluate_List] Where [VariableDescription] = " & _
                    "'" & ArraySelectedVariables(i) & "'")
            End If
            
            With myrs
                .Edit
                If mm1.Checked = True Then
                    If !Selected = "X" Then
                        !Selected = "XY"
                    ElseIf !Selected = "XY" Then
                        !Selected = "XY"
                    Else
                        !Selected = "Y"
                    End If
                Else
                    If !Selected = "XY" Then
                        !Selected = "X"
                    ElseIf !Selected = "X" Then
                        !Selected = "X"
                    Else
                        !Selected = Null
                    End If
                End If
                .Update
            End With
        Next i
    Err.Clear
    On Error Resume Next
    myrs.Close
    Exit Sub
    Unload frmWait2
    Unload frmWait3
    Unload frmWait
    Screen.MousePointer = vbDefault
Error1: MsgBox Err.Description & " in PutVariablesInArray3."
End Sub


Public Sub CreateExpOUTTable()
Dim n As Integer
Dim myrsOUT As Recordset
Dim myrsOutInfo As Recordset
Dim myrsOUTvsEXP As Recordset
Dim i As Integer
Dim TheOutFile As String

    On Error GoTo Error1
    Err.Clear
    On Error Resume Next

        n = 0
        dbXbuild.Execute "Create Table [Exp_OUT] ([OUT_File] Text, [Variable] Text, " _
            & "[Run] Integer, [TRNO] Integer, [ExperimentID] Text, [CropID] Text, [ExpCheck] Text, [Axis] Text)"
        Set myrsOUTvsEXP = dbXbuild.OpenRecordset("Exp_OUT")
        
        For i = 1 To NumberOutTables
            TheOutFile = Mid(OUTTableNames(i), 1, Len(OUTTableNames(i)) - 4)
             If ExpData_vs_Simulated = 1 Then
                Set myrsOUT = dbXbuild.OpenRecordset("Select * From [" & TheOutFile _
                    & "_List_EXPOUT] Where Selected = 'Y' Or Selected = 'XY'")
             Else
                Set myrsOUT = dbXbuild.OpenRecordset("Select * From [" & TheOutFile _
                    & "_List] Where Selected = 'Y' Or Selected = 'XY'")
             End If
                Set myrsOutInfo = dbXbuild.OpenRecordset("Select * From [" & TheOutFile _
                    & "_File_Info] Where Selected = 'Y' Or Selected = 'XY' Order By [RunNumber]")
            
            If Not (myrsOutInfo.EOF And myrsOutInfo.BOF) Then
                myrsOutInfo.MoveFirst
                Do While Not myrsOutInfo.EOF
                    With myrsOUT
                        If Not (.EOF And .BOF) Then
                            .MoveFirst
                            Do While Not .EOF
                                n = n + 1
                                myrsOUTvsEXP.AddNew
                                myrsOUTvsEXP.Fields("OUT_File") = TheOutFile
                                myrsOUTvsEXP.Fields("Variable") = !VariableID
                                ReDim Preserve XYVariables(n + 1)
                                ReDim Preserve XYRuns(n + 1)
                                XYVariables(n) = Trim(!VariableDescription)
                                XYTitle = TheOutFile
                                myrsOUTvsEXP.Fields("Run") = myrsOutInfo.Fields("RunNumber").Value
                                XYRuns(n) = myrsOutInfo.Fields("RunNumber").Value & " " & "Run" & _
                                    " " & myrsOutInfo.Fields("RunDescription").Value
                                myrsOUTvsEXP.Fields("TRNO") = myrsOutInfo.Fields("TRNO").Value
                                myrsOUTvsEXP.Fields("ExperimentID") = myrsOutInfo.Fields("ExperimentID").Value
                                myrsOUTvsEXP.Fields("CropID") = myrsOutInfo.Fields("CropID").Value
                                If ShowExperimentalData = 1 Then
                                    myrsOUTvsEXP.Fields("ExpCheck") = _
                                        FindExpVariable(myrsOutInfo.Fields("CropID").Value, _
                                        myrsOutInfo.Fields("ExperimentID").Value, !VariableID)
                                Else
                                    myrsOUTvsEXP.Fields("ExpCheck") = "NO"
                                End If
                                myrsOUTvsEXP.Update
                                .MoveNext
                            Loop
                        End If
                    End With
                    myrsOutInfo.MoveNext
                Loop
            End If
        Next i
        myrsOutInfo.Close
        myrsOUT.Close
        
        Err.Clear
        
         myrsOUTvsEXP.Close
            dbXbuild.Execute "UPDATE [Exp_OUT]" & _
                " SET [Axis] = 'Y'"
       
    Err.Clear
    On Error Resume Next
    myrsOUT.Close
    myrsOutInfo.Close
    myrsOUTvsEXP.Close
    Exit Sub
Error1:
    Unload frmWait2
    Unload frmWait3
    Unload frmWait
    Screen.MousePointer = vbDefault
'MsgBox Err.Description & " CreateExpOUTTable"
'MsgBox "Error in the experimental data."
ErrorFound = True
End Sub

Public Sub AddXvariables()
Dim i As Integer
Dim myrsOUT As Recordset
Dim myrsOutInfo As Recordset
Dim myrsOUTvsEXP As Recordset
Dim TheOutFile As String
Dim n As Integer
    n = 0
    On Error GoTo Error1
        Set myrsOUTvsEXP = dbXbuild.OpenRecordset("Select * From [Exp_OUT]")
        For i = 1 To NumberOutTables
            TheOutFile = Mid(OUTTableNames(i), 1, Len(OUTTableNames(i)) - 4)
                Set myrsOUT = dbXbuild.OpenRecordset("Select * From [" & TheOutFile _
                    & "_List] Where Selected = 'X' Or Selected = 'XY'")
                Set myrsOutInfo = dbXbuild.OpenRecordset("Select * From [" & TheOutFile _
                    & "_File_Info] Where Selected = 'Y' Order By [RunNumber]")
            If Not (myrsOutInfo.EOF And myrsOutInfo.BOF) Then
                myrsOutInfo.MoveFirst
                Do While Not myrsOutInfo.EOF
                    With myrsOUT
                        If Not (.EOF And .BOF) Then
                            .MoveFirst
                            Do While Not .EOF
                                n = n + 1
                                myrsOUTvsEXP.AddNew
                                myrsOUTvsEXP.Fields("OUT_File") = TheOutFile
                                myrsOUTvsEXP.Fields("Variable") = !VariableID
                                ReDim Preserve XYVariables(n + 1)
                                ReDim Preserve XYRuns(n + 1)
                                XYVariables(n) = Trim(!VariableDescription)
                                XYTitle = TheOutFile
                                myrsOUTvsEXP.Fields("Run") = myrsOutInfo.Fields("RunNumber").Value
                                XYRuns(n) = myrsOutInfo.Fields("RunNumber").Value & " " & "Run" & _
                                    " " & myrsOutInfo.Fields("RunDescription").Value
                                myrsOUTvsEXP.Fields("TRNO") = myrsOutInfo.Fields("TRNO").Value
                                myrsOUTvsEXP.Fields("ExperimentID") = myrsOutInfo.Fields("ExperimentID").Value
                                myrsOUTvsEXP.Fields("CropID") = myrsOutInfo.Fields("CropID").Value
                                '
                                If ShowExperimentalData = 1 Then
                                    myrsOUTvsEXP.Fields("ExpCheck") = _
                                        FindExpVariable(myrsOutInfo.Fields("CropID").Value, _
                                        myrsOutInfo.Fields("ExperimentID").Value, !VariableID)
                                Else
                                    myrsOUTvsEXP.Fields("ExpCheck") = "NO"
                                End If
                                
                                'myrsOUTvsEXP.Fields("ExpCheck") = "NO"
                                myrsOUTvsEXP.Fields("Axis") = "X"
                                myrsOUTvsEXP.Update
                                .MoveNext
                            Loop
                        End If
                    End With
                    myrsOutInfo.MoveNext
                Loop
            End If
        Next i
        Err.Clear
        On Error Resume Next
        myrsOutInfo.Close
        myrsOUT.Close
        myrsOUTvsEXP.Close
    Exit Sub
Error1:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description & " in AddXvariables/frmSelection."
End Sub

Public Sub AddXColumnsToGraph_Data()
Dim myrsOUTvsEXP As Recordset
Dim i As Integer
    On Error GoTo Error1
    Set myrsOUTvsEXP = dbXbuild.OpenRecordset("Select * From [Exp_OUT] Where [Axis] = 'X'")
    i = 1
    NumberOUTVariables = 0
    With myrsOUTvsEXP
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do While Not .EOF
                dbXbuild.Execute "Alter Table [Graph_Data] Add Column [" & _
                    "X_Axis_" & .Fields("Variable").Value & _
                    "(" & .Fields("OUT_File") & ") Run " & .Fields("RUN") & "] Single"
                .MoveNext
                i = i + 1
            Loop
        End If
        NumberOUTVariables = i - 1
    End With
    Err.Clear
    On Error Resume Next
    myrsOUTvsEXP.Close
    Exit Sub
Error1:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description & " in AddXcolumnsToGraph_Data/frmSelection."
End Sub

Public Sub AddYColumnsToGraph_Data()
Dim myrsOUTvsEXP As Recordset
Dim i As Integer
    On Error GoTo Error1
    Set myrsOUTvsEXP = dbXbuild.OpenRecordset("Select * From [Exp_OUT] Where [Axis] = 'Y'")
    i = 1
    NumberOUTVariables = 0
    With myrsOUTvsEXP
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do While Not .EOF
                dbXbuild.Execute "Alter Table [Graph_Data] Add Column [" & _
                      .Fields("Variable").Value & _
                    "(" & .Fields("OUT_File") & ") Run " & .Fields("RUN") & "] Single"
                .MoveNext
                i = i + 1
            Loop
        End If
        NumberOUTVariables = i - 1
    End With
    Err.Clear
    On Error Resume Next
    myrsOUTvsEXP.Close
    Exit Sub
Error1:
    Screen.MousePointer = vbDefault
   ' MsgBox Err.Description & " in AddYcolumnsToGraph_Data/frmSelection."
    ErrorFound = True
    Unload frmWait2
    Unload frmWait3
Unload frmWait
End Sub

Public Sub AddExpVariables()
Dim myrsOUTvsEXP_Y As Recordset
Dim myrsOUTvsEXP_X As Recordset
Dim i As Integer
Dim Experiments As Boolean
    
    On Error GoTo Error1
    If ExpData_vs_Simulated = 1 Then
        Set myrsOUTvsEXP_Y = dbXbuild.OpenRecordset("Exp_OUT")
    Else
        Set myrsOUTvsEXP_Y = dbXbuild.OpenRecordset("Select * From [Exp_OUT] Where [ExpCheck] = 'YES' And [Axis] = 'Y'")
    End If
    With myrsOUTvsEXP_Y
        If .EOF And .BOF Then
            Experiments = False
        Else
            .MoveLast
            Date_NumberExpVariables = .RecordCount
            Set myrsOUTvsEXP_X = dbXbuild.OpenRecordset("Select * From [Exp_OUT] Where [ExpCheck] = 'YES' " & _
            "And [Axis] = 'X'")
                If myrsOUTvsEXP_X.EOF And myrsOUTvsEXP_X.BOF Then
                    Experiments = False
                Else
                    Experiments = True
                End If
        End If
    End With
    If Experiments = True Then
        i = 1
        With myrsOUTvsEXP_X
            .MoveFirst
            Do While Not .EOF
                dbXbuild.Execute "Alter Table [Graph_Data] Add Column [" & _
                    "X_Exp_Axis_" & .Fields("Variable").Value & _
                    "(" & .Fields("ExperimentID") & " " & .Fields("CropID") _
                    & "T) TRT " & .Fields("TRNO") & "/" & .Fields("Run") & "] Single"
                    .MoveNext
                i = i + 1
            Loop
            NumberExpVariables_X = .RecordCount
            '.Close
        End With
        'Paul
        If ExpData_vs_Simulated <> 1 Then
            With myrsOUTvsEXP_Y
                .MoveFirst
                Do While Not .EOF
                    dbXbuild.Execute "Alter Table [Graph_Data] Add Column [" & _
                        .Fields("Variable").Value & _
                        "(" & .Fields("ExperimentID") & " " & .Fields("CropID") _
                        & "T) TRT " & .Fields("TRNO") & "/" & .Fields("Run") & "] Single"
                        .MoveNext
                        i = i + 1
                Loop
                NumberExpVariables_Y = .RecordCount
                '.Close
            End With
            NumberExpVariables = NumberExpVariables_X + NumberExpVariables_Y
        End If
    Else
        If ShowX_Axis = 0 And ExpData_vs_Simulated = 0 Then
            NumberExpVariables_X = 0
            NumberExpVariables_Y = 0
            NumberExpVariables = 0
            dbXbuild.Execute "Update [Exp_OUT] Set [ExpCheck] = 'NO'"
        Else
            With myrsOUTvsEXP_Y
                If Not (.EOF And .BOF) Then
                    .MoveFirst
                    Do While Not .EOF
                        dbXbuild.Execute "Alter Table [Graph_Data] Add Column [" & _
                            .Fields("Variable").Value & _
                            "(" & .Fields("ExperimentID") & " " & .Fields("CropID") _
                            & "T) TRT " & .Fields("TRNO") & "/" & .Fields("Run") & "] Single"
                            .MoveNext
                        'i = i + 1
                    Loop
                    NumberExpVariables_Y = .RecordCount
                End If
            End With
        End If
    End If
    Err.Clear
    On Error Resume Next
    myrsOUTvsEXP_X.Close
    myrsOUTvsEXP_Y.Close
    Exit Sub
Error1:
Screen.MousePointer = vbDefault
    'MsgBox Err.Description & " in AddExpVariables/frmSelection."
    'MsgBox "Error with experimental data."
    ErrorFound = True
    Unload frmWait2
    Unload frmWait3
    Unload frmWait
End Sub



Public Sub FillOUTGraphData_AfterPlanting()
Dim VariableCount As Integer
Dim myrsTest As Recordset
Dim myrsGraphData  As Recordset
Dim myrsOUTvsEXP  As Recordset
Dim myrsInfo  As Recordset
Dim myrsDateTemp  As Recordset
Dim myrsGraphSelect  As Recordset
    On Error GoTo Error1
    Set myrsGraphData = dbXbuild.OpenRecordset("Graph_Data")
    Set myrsOUTvsEXP = dbXbuild.OpenRecordset("Select * From Exp_OUT Where Axis = 'Y'")
    VariableCount = 0
    
    ReDim FirstPlantDay(2)
    ReDim LastPlantDay(2)

    FirstPlantDay(1) = 9999
    
    With myrsOUTvsEXP
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Set myrsTest = dbXbuild.OpenRecordset("Select * From [" & _
                .Fields("OUT_File").Value & "_File_Info] Where [RunNumber] = " _
                & .Fields("Run").Value & " And [TRNO] = " _
                & .Fields("TRNO").Value)

            If Not (myrsTest.EOF And myrsTest.BOF) Then
              '  Set myrsInfo = dbXbuild.OpenRecordset("Select * From [" & _
                    .Fields("OUT_File").Value & "_OUT] Where [RunNumber] = " _
                    & .Fields("Run").Value)
                If CDAYexists = True Then
                    Set myrsInfo = dbXbuild.OpenRecordset("Select * From [" & _
                        .Fields("OUT_File").Value & "_OUT] Where [RunNumber] = " _
                        & .Fields("Run").Value & " Order By CDAY")
                ElseIf DAPexists = True Then
                    
                    Set myrsInfo = dbXbuild.OpenRecordset("Select * From [" & _
                        .Fields("OUT_File").Value & "_OUT] Where [RunNumber] = " _
                        & .Fields("Run").Value & " Order By DAP")
                Else
                    Set myrsInfo = dbXbuild.OpenRecordset("Select * From [" & _
                        .Fields("OUT_File").Value & "_OUT] Where [RunNumber] = " _
                        & .Fields("Run").Value & " Order By DAS")
                End If
                    
                    
                    
                    
                    If Not (myrsInfo.EOF And myrsInfo.BOF) Then
                        myrsInfo.MoveFirst
                        Do While Not (myrsInfo.EOF)
                            myrsGraphData.AddNew
                            
                            myrsGraphData.Fields(.Fields("Variable").Value & _
                                "(" & .Fields("OUT_File") & ") Run " & .Fields("RUN")).Value = _
                                myrsInfo.Fields(.Fields("Variable")).Value
                            If ShowRealDates = 0 Then
                                If CDAYexists = True Then
                                    myrsGraphData.Fields("Date").Value = myrsInfo.Fields("CDAY").Value
                                    If FirstPlantDay(1) = 9999 Then
                                        FirstPlantDay(1) = myrsInfo.Fields("CDAY").Value
                                    End If
                                ElseIf DAPexists = True Then
                                    myrsGraphData.Fields("Date").Value = myrsInfo.Fields("DAP").Value
                                    Dim mn
                                    If myrsInfo.Fields("DAP").Value = 94 Then
                                        mn = 0
                                    End If
                                    If FirstPlantDay(1) = 9999 Then FirstPlantDay(1) = myrsInfo.Fields("DAP").Value
                                Else
                                    myrsGraphData.Fields("Date").Value = myrsInfo.Fields("DAS").Value
                                    If FirstPlantDay(1) = 9999 Then FirstPlantDay(1) = myrsInfo.Fields("DAS").Value
                                End If
                            ElseIf ShowRealDates = 2 Then
                                myrsGraphData.Fields("Date").Value = Val(Mid(myrsInfo.Fields("Date").Value, 6))
                            End If
                               ' myrsGraphData.Fields("TheOrder").Value = Mid(myrsInfo.Fields("Date").Value, 6)
                                myrsGraphData.Fields("TheOrder").Value = myrsInfo.Fields("Date").Value
                            myrsGraphData.Update
                            myrsInfo.MoveNext
                        Loop
                    End If
                    .MoveLast
                End If
                VariableCount = .RecordCount
            End If
        End With
        myrsGraphData.MoveLast
         LastPlantDay(1) = myrsGraphData.Fields("Date").Value
        'Add the rest of out variables/data adjusting the dates
        Dim bb As Integer
        bb = 1
        If VariableCount > 1 Then
            ReDim Preserve FirstPlantDay(VariableCount + 1)
            ReDim Preserve LastPlantDay(VariableCount + 1)

            With myrsOUTvsEXP
                .MoveFirst
                .MoveNext
                Set myrsTest = dbXbuild.OpenRecordset("Select * From [" & _
                    .Fields("OUT_File").Value & "_File_Info] Where [RunNumber] = " _
                    & .Fields("Run").Value & " And [TRNO] = " _
                    & .Fields("TRNO").Value)
                    If Not (myrsTest.EOF And myrsTest.BOF) Then
                        Do While Not .EOF
                            frmWait.prgLoad.Value = .PercentPosition
                            
                            'Set myrsInfo = dbXbuild.OpenRecordset("Select * From [" & _
                                .Fields("OUT_File").Value & "_OUT] Where [RunNumber] = " _
                                & .Fields("Run").Value)
                If CDAYexists = True Then
                    Set myrsInfo = dbXbuild.OpenRecordset("Select * From [" & _
                        .Fields("OUT_File").Value & "_OUT] Where [RunNumber] = " _
                        & .Fields("Run").Value & " Order By CDAY")
                ElseIf DAPexists = True Then
                    Set myrsInfo = dbXbuild.OpenRecordset("Select * From [" & _
                        .Fields("OUT_File").Value & "_OUT] Where [RunNumber] = " _
                        & .Fields("Run").Value & " Order By DAP")
                Else
                    Set myrsInfo = dbXbuild.OpenRecordset("Select * From [" & _
                        .Fields("OUT_File").Value & "_OUT] Where [RunNumber] = " _
                        & .Fields("Run").Value & " Order By DAS")
                End If
                            If Not (myrsInfo.EOF And myrsInfo.BOF) Then
                                myrsGraphData.MoveFirst
                                
                                myrsInfo.MoveFirst
                                FirstPlantDay(bb + 1) = 9999
                                Do While Not (myrsInfo.EOF)
                                    If ShowRealDates = 0 Then
                                       ' Set myrsDateTemp = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [TheOrder] = '" _
                                       '     & Mid(myrsInfo.Fields("Date"), 6) & "'")
                                            If CDAYexists = True Then
                                                Set myrsDateTemp = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [DATE] = " _
                                                    & myrsInfo.Fields("CDAY"))
                                                If FirstPlantDay(bb + 1) = 9999 Then FirstPlantDay(bb + 1) = myrsInfo.Fields("CDAY")
                                                LastPlantDay(bb + 1) = myrsInfo.Fields("CDAY").Value
                                            ElseIf DAPexists = True Then
                                                Set myrsDateTemp = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [DATE] = " _
                                                    & myrsInfo.Fields("DAP"))
                                                If FirstPlantDay(bb + 1) = 9999 Then FirstPlantDay(bb + 1) = myrsInfo.Fields("DAP")
                                                LastPlantDay(bb + 1) = myrsInfo.Fields("DAP").Value
                                            Else
                                                Set myrsDateTemp = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [DATE] = " _
                                                    & myrsInfo.Fields("DAS"))
                                                If FirstPlantDay(bb + 1) = 9999 Then FirstPlantDay(bb + 1) = myrsInfo.Fields("DAS")
                                                LastPlantDay(bb + 1) = myrsInfo.Fields("DAS").Value
                                            End If
                                    ElseIf ShowRealDates = 2 Then
                                        Set myrsDateTemp = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [Date] = " _
                                            & Val(Mid(myrsInfo.Fields("Date"), 6)))
                                    End If
                                    If myrsDateTemp.EOF And myrsDateTemp.BOF Then
                                        myrsGraphData.AddNew
                                        myrsGraphData.Fields(.Fields("Variable").Value & _
                                            "(" & .Fields("OUT_File") & ") Run " & .Fields("RUN")).Value = _
                                            myrsInfo.Fields(.Fields("Variable")).Value
                                        If ShowRealDates = 0 Then
                                            If CDAYexists = True Then
                                                myrsGraphData.Fields("Date").Value = myrsInfo.Fields("CDAY").Value
                                                
                                            ElseIf DAPexists = True Then
                                                myrsGraphData.Fields("Date").Value = myrsInfo.Fields("DAP").Value
                                            Else
                                                myrsGraphData.Fields("Date").Value = myrsInfo.Fields("DAS").Value
                                            End If
                                        ElseIf ShowRealDates = 2 Then
                                            myrsGraphData.Fields("Date").Value = Val(Mid(myrsInfo.Fields("Date").Value, 6))
                                        End If
                                       ' myrsGraphData.Fields("TheOrder").Value = Mid(myrsInfo.Fields("Date"), 6)
                                        myrsGraphData.Fields("TheOrder").Value = myrsInfo.Fields("Date")
                                        myrsGraphData.Update
                                    Else
                                        If ShowRealDates = 0 Then
                                            If CDAYexists = True Then
                                                Set myrsGraphSelect = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [Date] = " _
                                                & myrsInfo.Fields("CDAY").Value)
                                            ElseIf DAPexists = True Then
                                                Set myrsGraphSelect = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [Date] = " _
                                                & myrsInfo.Fields("DAP").Value)
                                            Else
                                                Set myrsGraphSelect = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [Date] = " _
                                                & myrsInfo.Fields("DAS").Value)
                                                
                                                
                                            End If
                                        ElseIf ShowRealDates = 2 Then
                                            Set myrsGraphSelect = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [Date] = " _
                                                & Val(Mid(myrsInfo.Fields("Date"), 6)))
                                        End If
                                        myrsGraphSelect.Edit
                                        myrsGraphSelect.Fields(.Fields("Variable").Value & _
                                            "(" & .Fields("OUT_File") & ") Run " & .Fields("RUN")).Value = _
                                            myrsInfo.Fields(.Fields("Variable")).Value
                                        myrsGraphSelect.Update
                                    End If
                                    myrsInfo.MoveNext
                                Loop
                            End If
                                'myrsInfo.MoveLast
                                'myrsInfo.MovePrevious
                                'If CDAYexists = True Then
                                '        LastPlantDay(bb + 1) = myrsInfo.Fields("CDAY").Value
                                '    ElseIf DAPexists = True Then
                                '        LastPlantDay(bb + 1) = myrsInfo.Fields("DAP").Value
                                '    Else
                                '        LastPlantDay(bb + 1) = myrsInfo.Fields("DAS").Value
                                'End If
                           ' LastPlantDay(bb + 1) = myrsGraphSelect.Fields("Date").Value
                            
                            
                            bb = bb + 1
                            .MoveNext
                        Loop
                        frmWait.prgLoad.Value = 100
                    End If
                    myrsTest.Close
                End With
            End If
    
    Err.Clear
    
    On Error Resume Next
    myrsTest.Close
    myrsGraphData.Close
    myrsOUTvsEXP.Close
    myrsInfo.Close
    myrsDateTemp.Close
    myrsGraphSelect.Close
    Exit Sub
Error1:
    Screen.MousePointer = vbDefault
    'MsgBox Err.Description & " in FillOUTGraphData_AfterPlanting/frmSelection."
    ErrorFound = True
    Unload frmWait2
    Unload frmWait3
Unload frmWait
End Sub


Public Sub FillOUT_X_GraphData_AfterPlanting()
Dim myrsTest As Recordset
Dim myrsGraphData As Recordset
Dim myrsOUTvsEXP As Recordset
Dim myrsInfo As Recordset
Dim myrsDateTemp As Recordset
Dim myrsGraphSelect As Recordset

    On Error GoTo Error1
    Set myrsGraphData = dbXbuild.OpenRecordset("Graph_Data")
    Set myrsOUTvsEXP = dbXbuild.OpenRecordset("Select * From Exp_OUT Where Axis = 'X'")
    With myrsOUTvsEXP
        If Not (.EOF And .BOF) Then
           DoEvents
         Else
            myrsOUTvsEXP.Close
            myrsGraphData.Close
            Exit Sub
         End If
        .MoveFirst
        Set myrsTest = dbXbuild.OpenRecordset("Select * From [" & _
            .Fields("OUT_File").Value & "_File_Info] Where [RunNumber] = " _
            & .Fields("Run").Value & " And [TRNO] = " _
            & .Fields("TRNO").Value)
        If Not (myrsTest.EOF And myrsTest.BOF) Then
            Do While Not .EOF
                Set myrsInfo = dbXbuild.OpenRecordset("Select * From [" & _
                    .Fields("OUT_File").Value & "_OUT] Where [RunNumber] = " _
                    & .Fields("Run").Value)
                If Not (myrsInfo.EOF And myrsInfo.BOF) Then
                    myrsGraphData.MoveFirst
                    myrsInfo.MoveFirst
                    Do While Not (myrsInfo.EOF)
                        If ShowRealDates = 0 Then
                           ' Set myrsDateTemp = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [TheOrder] = '" _
                                & Mid(myrsInfo.Fields("Date"), 6) & "'")
                            Set myrsDateTemp = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [TheOrder] = '" _
                                & myrsInfo.Fields("Date") & "'")
                        ElseIf ShowRealDates = 2 Then
                            Set myrsDateTemp = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [Date] = " _
                                & Val(Mid(myrsInfo.Fields("Date"), 6)))
                        End If
                        If myrsDateTemp.EOF And myrsDateTemp.BOF Then
                            myrsGraphData.AddNew
                            myrsGraphData.Fields("X_Axis_" & .Fields("Variable").Value & _
                                "(" & .Fields("OUT_File") & ") Run " & .Fields("RUN")).Value = _
                                myrsInfo.Fields(.Fields("Variable")).Value
                            If ShowRealDates = 0 Then
                                If CDAYexists = True Then
                                    myrsGraphData.Fields("Date").Value = myrsInfo.Fields("CDAY").Value
                                ElseIf DAPexists = True Then
                                    myrsGraphData.Fields("Date").Value = myrsInfo.Fields("DAP").Value
                                Else
                                    myrsGraphData.Fields("Date").Value = myrsInfo.Fields("DAS").Value
                                End If
                            ElseIf ShowRealDates = 2 Then
                                myrsGraphData.Fields("Date").Value = Val(Mid(myrsInfo.Fields("Date"), 6))
                            End If
                            myrsGraphData.Update
                        Else
                            If ShowRealDates = 0 Then
                                Set myrsGraphSelect = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [TheOrder] = '" _
                                    & myrsInfo.Fields("Date") & "'")
                            ElseIf ShowRealDates = 2 Then
                                Set myrsGraphSelect = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [Date] = " _
                                    & Val(Mid(myrsInfo.Fields("Date"), 6)))
                            End If
                            myrsGraphSelect.Edit
                            myrsGraphSelect.Fields("X_Axis_" & .Fields("Variable").Value & _
                                "(" & .Fields("OUT_File") & ") Run " & .Fields("RUN")).Value = _
                                myrsInfo.Fields(.Fields("Variable")).Value
                            myrsGraphSelect.Update
                        End If
                            myrsInfo.MoveNext
                    Loop
                End If
                .MoveNext
            Loop
        End If
    End With
    
    Err.Clear
    On Error Resume Next
    myrsTest.Close
    myrsGraphData.Close
    myrsOUTvsEXP.Close
    myrsInfo.Close
    myrsDateTemp.Close
    myrsGraphSelect.Close
    Exit Sub
Error1:
Screen.MousePointer = vbDefault
MsgBox Err.Description & " in FillOUT_X_GraphData_AfterPlanting/frmSelection."
   ' MsgBox "Error with experimental data."
    ErrorFound = True
    Unload frmWait2
    Unload frmWait3
Unload frmWait
End Sub

Public Sub AddExpDataToGraphData_AfterPlanting()
Dim myrsOUTvsEXP As Recordset
Dim myrsDateTemp As Recordset
Dim myrsGraphSelect As Recordset
Dim myrsGraphData As Recordset
Dim myrsOUTvsEXP_Sel As Recordset
Dim ExpTableName As String
Dim myrsExp As Recordset
Dim myrsPlantingDay As Recordset
Dim myDaysAfterPlanting As String
Dim myrstempData As Recordset
Dim bbb

    On Error GoTo Error1
    Set myrsGraphData = dbXbuild.OpenRecordset("Graph_Data")
    myrsGraphData.MoveFirst
    Set myrsOUTvsEXP_Sel = dbXbuild.OpenRecordset("Select * From [Exp_OUT] Where " & _
    "[ExpCheck] = 'YES' And [Axis] = 'X'")
    With myrsOUTvsEXP_Sel
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do While Not .EOF
                ExpTableName = .Fields("ExperimentID").Value & _
                    " " & .Fields("CropID").Value & "T"
                Set myrsExp = dbXbuild.OpenRecordset("Select * From [" & _
                    ExpTableName & "] Where [TRNO] = " & .Fields("TRNO") & " And " & _
                    .Fields("Variable") & " Is Not Null")
                If Not (myrsExp.EOF And myrsExp.BOF) Then
                    myrsGraphData.MoveFirst
                    myrsExp.MoveFirst
                   
               '    bbb = myrsExp.Fields("TheOrder")
  
               '  Dim myrsbbb As Recordset
               '  myrsbbb = dbXbuild.OpenRecordset("Select * From")
                 
                 '   myDaysAfterPlanting = CalculateDays(myrsGraphData.Fields("Date").Value, _
                         myrsGraphData.Fields("TheOrder").Value)
                    Do While Not (myrsExp.EOF)
                            Set myrsDateTemp = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [TheOrder] = '" _
                                & myrsExp.Fields("TheOrder") & "' and " & _
                                "[X_Exp_Axis_" & .Fields("Variable").Value & _
                                "(" & ExpTableName & ") TRT " & .Fields("TRNO").Value & _
                                "/" & .Fields("Run").Value & "] Is Null")
                        If myrsDateTemp.EOF And myrsDateTemp.BOF Then
                            myrsGraphData.AddNew
                            myrsGraphData.Fields("X_Exp_Axis_" & .Fields("Variable").Value & _
                                "(" & ExpTableName & ") TRT " & .Fields("TRNO").Value & "/" & .Fields("Run").Value) = _
                                myrsExp.Fields(.Fields("Variable")).Value
                                '
                                myrsGraphData.Fields("TheOrder") = myrsExp.Fields("TheOrder")
                              '  If ShowRealDates = 0 Then
                                '    myrsGraphData.Fields("Date").Value = RemoveDaysAfterPlantingToTheOrder(myDaysAfterPlanting, myrsExp.Fields("TheOrder"))
                            '    ElseIf ShowRealDates = 2 Then
                                '    myrsGraphData.Fields("Date") = Mid(myrsExp.Fields("TheOrder"), 6)
                               ' End If
                            
                            myrsGraphData.Update
                        Else
                            myrsDateTemp.Edit
                            myrsDateTemp.Fields("X_Exp_Axis_" & .Fields("Variable").Value & _
                                "(" & ExpTableName & ") TRT " & .Fields("TRNO").Value & "/" & .Fields("Run").Value) = _
                                myrsExp.Fields(.Fields("Variable")).Value
                            myrsDateTemp.Update
                        End If
                        myrsExp.MoveNext
                    Loop
                End If
                .MoveNext
            Loop
        End If
    End With
    ''
    Dim myMYRS As Recordset
    
    Set myrsOUTvsEXP_Sel = dbXbuild.OpenRecordset("Select * From [Exp_OUT] Where " & _
    "[ExpCheck] = 'YES' And [Axis] = 'Y'")
    With myrsOUTvsEXP_Sel
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do While Not .EOF
                ExpTableName = .Fields("ExperimentID").Value & _
                    " " & .Fields("CropID").Value & "T"
                Set myrsExp = dbXbuild.OpenRecordset("Select * From [" & _
                    ExpTableName & "] Where [TRNO] = " & .Fields("TRNO") & " And [" & _
                    .Fields("Variable").Value & "] Is Not Null")
               ' Debug.Print "Select * From [" & _
                    ExpTableName & "] Where [TRNO] = " & .Fields("TRNO") & " And [" & _
                    .Fields("Variable").Value & "] Is Not Null"
                
                If Not (myrsExp.EOF And myrsExp.BOF) Then
                    myrsGraphData.MoveFirst
                    myrsExp.MoveFirst
                  '   bbb = myrsExp.Fields("TheOrder")
                    ''''''''''''
                    myrsGraphData.MoveFirst
                   ' Dim vv
                      ' vv = "Select * From [" & _
                        .Fields("OUT_File") & "_OUT] Where RunNumber = " & .Fields("RUN").Value
                       
                       Set myrstempData = dbXbuild.OpenRecordset("Select * From [" & _
                        .Fields("OUT_File") & "_OUT] Where RunNumber = " & .Fields("RUN").Value)
                       myrstempData.MoveFirst
                     '  vv = myrstempData.Fields("Date").Value
                     If CDAYexists = True Then
                         
                      '   bbb = myrsExp.Fields("TheOrder")
                         
                            myDaysAfterPlanting = CalculateDays(myrstempData.Fields("CDAY").Value, _
                            myrstempData.Fields("Date").Value)
                    ElseIf DAPexists = True Then
                        myDaysAfterPlanting = CalculateDays(myrstempData.Fields("DAP").Value, _
                        myrstempData.Fields("Date").Value)
                    Else
                        myDaysAfterPlanting = CalculateDays(myrstempData.Fields("DAS").Value, _
                        myrstempData.Fields("Date").Value)
                    End If
                    myrsExp.MoveFirst
                    Do While Not (myrsExp.EOF)
                           
                           Dim FindDate As Single
                           'If ShowRealDates = 0 Then
                                FindDate = RemoveDaysAfterPlantingToTheOrder(myDaysAfterPlanting, myrsExp.Fields("TheOrder"))
                           ' ElseIf ShowRealDates = 2 Then
                            '    myrsGraphData.Fields("Date") = Mid(myrsExp.Fields("TheOrder"), 6)
                           ' End If
                            
                            
                        '    Set myrsDateTemp = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [TheOrder] = '" _
                        '        & myrsExp.Fields("TheOrder").Value & "' And [" & .Fields("Variable").Value & _
                         '       "(" & ExpTableName & ") TRT " & .Fields("TRNO").Value & "/" & _
                         '       .Fields("Run").Value & "] Is Null")
                        
                            Set myrsDateTemp = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where Date = " & _
                            FindDate)
                        
                        
                        
                        'If myrsDateTemp.EOF And myrsDateTemp.BOF Then
                                myrsGraphData.AddNew
                                myrsGraphData.Fields(.Fields("Variable").Value & _
                                    "(" & ExpTableName & ") TRT " & .Fields("TRNO").Value & "/" & _
                                    .Fields("Run").Value).Value = _
                                    myrsExp.Fields(.Fields("Variable")).Value
                                    myrsGraphData.Fields("TheOrder").Value = myrsExp.Fields("TheOrder")
                                    myrsGraphData.Fields("Date").Value = FindDate
                            myrsGraphData.Update

                       ' ElseIf IsNull(myrsDateTemp.Fields(.Fields("Variable").Value & _
                      '          "(" & ExpTableName & ") TRT " & .Fields("TRNO").Value & "/" & _
                      '          .Fields("Run").Value)) = True Then
                      '      myrsGraphData.AddNew
                      '      myrsGraphData.Fields(.Fields("Variable").Value & _
                       '         "(" & ExpTableName & ") TRT " & .Fields("TRNO").Value & "/" & _
                       '         .Fields("Run").Value) = myrsExp.Fields(.Fields("Variable")).Value
                       '     myrsGraphData.Fields("Date").Value = FindDate
                        '    myrsGraphData.Fields("TheOrder").Value = "11111"
                        '    myrsGraphData.Update
                        'ElseIf IsNull(myrsDateTemp.Fields(.Fields("Variable").Value & _
                         '       "(" & ExpTableName & ") TRT " & .Fields("TRNO").Value & "/" & _
                         '       .Fields("Run").Value)) = True Then
                          '  myrsDateTemp.Edit
                          '  myrsDateTemp.Fields(.Fields("Variable").Value & _
                          '      "(" & ExpTableName & ") TRT " & .Fields("TRNO").Value & "/" & _
                           '     .Fields("Run").Value) = myrsExp.Fields(.Fields("Variable")).Value
                           ' myrsDateTemp.Fields("TheOrder").Value = "11111"
                           ' myrsDateTemp.Update
                        
                        'End If
                        myrsExp.MoveNext
                    Loop
                End If
                .MoveNext
            Loop
        End If
    End With
     ' dbXbuild.Execute "Delete * From [Graph_Data] Where [TheOrder] Is Null"
    

    
    
    
    
    
    On Error Resume Next
    Err.Clear
    myrsOUTvsEXP.Close
    myrsDateTemp.Close
    myrsGraphSelect.Close
    myrsGraphData.Close
    myrsExp.Close
    myrsOUTvsEXP_Sel.Close
    
    
    'ChangeDateFormat
    
    Exit Sub
Error1:
Screen.MousePointer = vbDefault
MsgBox Err.Description & " in AddExpDataToGraphData_AfterPlanting/frmSelection."
    'MsgBox "Error with experimental data."
    ErrorFound = True
    Unload frmWait2
    Unload frmWait3
Unload frmWait
End Sub

Public Sub ChangeDateFormat()
Dim myrsGraphData As Recordset
    On Error GoTo Error1
    If ShowRealDates = 1 Then
        Set myrsGraphData = dbXbuild.OpenRecordset("Graph_Data")
        With myrsGraphData
            If Not (.EOF And .BOF) Then
                .MoveFirst
                Do While Not .EOF
                    .Edit
                        If IsNull(!TheOrder) = False Then
                            !realdate = Format(TheDate(!TheOrder), "mm/dd/yyyy")
                        End If
                    .Update
                    .MoveNext
                Loop
            End If
            .Close
        End With
    End If
    Exit Sub
    Err.Clear
    On Error Resume Next
    myrsGraphData.Close
Error1:
Screen.MousePointer = vbDefault
MsgBox Err.Description & " in ChangeDateFormat/frmSelection."
End Sub




Public Sub AddYColumns_GraphData_T_File()
Dim myrsEXPlist As Recordset
Dim myrsDataTemp As Recordset
Dim myrsDATA As Recordset
Dim myrsEXPinfo As Recordset
Dim myrsInfo As Recordset
Dim i As Integer
Dim bb As Integer
Dim nn As Integer
Dim kk As Integer
Dim hh As Integer
Dim gg As Integer
Dim ss As Integer
Dim NumberYSelected As Integer
Dim prmVariables() As Integer
Dim prmNumberTRNO() As Integer
Dim ArrVarID() As String
Dim ArrTRNO() As Integer
    ErrorFound = False
    On Error GoTo Error1
    i = 1
    bb = 1
    kk = 1
    ss = 1
    hh = 1
    gg = 1
    NumberYSelected = 1
    ReDim ArrVarID(10, 1000)
    ReDim prmVariables(10)
    ReDim ArrTRNO(10, 1000)
    ReDim prmNumberTRNO(10)
    For i = 1 To T_File_Y_NumberOutTables
            Set myrsEXPlist = dbXbuild.OpenRecordset("Select * From [" & _
                Y_ArrTableNames(i) & "_List] Where [Selected] = 'Y'")
            If Not (myrsEXPlist.EOF And myrsEXPlist.BOF) Then
                With myrsEXPlist
                    .MoveFirst
                    Do While Not .EOF
                        ArrVarID(i, hh) = !VariableID
                        NumberYSelected = NumberYSelected + 1
                        hh = hh + 1
                        .MoveNext
                    Loop
                    prmVariables(i) = hh - 1
                End With
            End If
    Next i
    NumberYSelected = NumberYSelected - 1
    If NumberYSelected > 0 Then
        DoEvents
    Else
        MsgBox "No variables have been selected."
        Exit Sub
    End If
    For i = 1 To T_File_Y_NumberOutTables
                Set myrsEXPinfo = dbXbuild.OpenRecordset("Select * From [" & _
                    Y_ArrTableNames(i) & "_File_Info] Where  [Selected] = 'Y'")
                With myrsEXPinfo
                    If Not (.EOF And .BOF) Then
                        bb = 1
                        .MoveFirst
                        Do While Not .EOF
                            ArrTRNO(i, bb) = !TRNO
                            bb = bb + 1
                            .MoveNext
                        Loop
                    End If
                End With
            prmNumberTRNO(i) = bb - 1
    Next i
    Err.Clear
    On Error Resume Next
    myrsEXPlist.Close
    myrsEXPinfo.Close
    Err.Clear
    On Error GoTo Error1
    For i = 1 To T_File_Y_NumberOutTables
            For bb = 1 To prmVariables(i)
                If ArrVarID(i, bb) <> "" Then
                    For ss = 1 To prmNumberTRNO(i)
                        If ArrTRNO(i, ss) <> 0 Then
                            dbXbuild.Execute "Alter Table [Graph_Data] Add Column [" & _
                                ArrVarID(i, bb) & _
                                "(" & Y_ArrTableNames(i) & ") TRT " & ArrTRNO(i, ss) & "] Single"
                        End If
                    Next ss
                End If
            Next bb
    Next i
    Set myrsDATA = dbXbuild.OpenRecordset("Graph_Data")
    NumberExpVariables_Y = myrsDATA.Fields.Count - 3 - NumberExpVariables_X
    If T_File_Y_NumberOutTables > 0 Then
                
                    '??????????????????
                        Set myrsDataTemp = dbXbuild.OpenRecordset("Graph_Data")
                        Set myrsInfo = dbXbuild.OpenRecordset("Select * From [" & Y_ArrTableNames(1) & _
                        "] Where [TRNO] = " & ArrTRNO(1, 1))
                    If NumberExpVariables_X = 0 Then

                        If ArrVarID(1, 1) <> "" Then
                            If Not (myrsInfo.BOF And myrsInfo.EOF) Then
                                myrsInfo.MoveFirst
                                Do While Not myrsInfo.EOF
                                        myrsDATA.AddNew
                                        If ShowRealDates = 1 Then
                                            myrsDATA.Fields("TheOrder") = myrsInfo.Fields("TheOrder").Value
                                            myrsDATA.Fields("RealDate") = TheDate(myrsInfo.Fields("TheOrder").Value)
                                        ElseIf ShowRealDates = 2 Then
                                            myrsDATA.Fields("Date") = myrsInfo.Fields("Date").Value
                                        End If
                                        myrsDATA.Fields(ArrVarID(1, 1) & "(" & Y_ArrTableNames(1) & ") TRT " & _
                                            ArrTRNO(1, 1)) = myrsInfo.Fields(ArrVarID(1, 1)).Value
                                        myrsDATA.Update
                                    myrsInfo.MoveNext
                                Loop
                            End If
                        End If
                End If
        
        
        
        ''''''''''''''''
        myrsDATA.Close
        
        dbXbuild.Execute "Delete * From [Graph_Data] Where [" & _
            ArrVarID(1, 1) & _
            "(" & Y_ArrTableNames(1) & ") TRT " & ArrTRNO(1, 1) & "] is Null"
        
        
        ''''''''''''''''
        
        Dim myrsTest As Recordset             '?????????????????
        Dim ContinueFillingTable As Boolean
        
        For i = 1 To T_File_Y_NumberOutTables
            'For ss = 2 To prmNumberTRNO(i)
            For ss = 1 To prmNumberTRNO(i)
                Set myrsInfo = dbXbuild.OpenRecordset("Select * From [" & Y_ArrTableNames(i) & _
                        "] Where [TRNO] = " & ArrTRNO(i, ss))
                For bb = 1 To prmVariables(i)
                
                    Set myrsTest = dbXbuild.OpenRecordset("Select * From " & _
                        "[Graph_Data] Where [" & ArrVarID(i, bb) & "(" & Y_ArrTableNames(i) & ") TRT " & _
                                                ArrTRNO(i, ss) & "] is Null ")
                   
                    With myrsTest
                        If Not (.EOF And .BOF) Then
                            ContinueFillingTable = True
                        Else
                            ContinueFillingTable = False
                        End If
                    End With
                If ContinueFillingTable = True Then

                        Dim myrsDATa_Graph_select As Recordset
                        If ArrVarID(i, bb) <> "" Then
                            If Not (myrsInfo.BOF And myrsInfo.EOF) Then
                                myrsInfo.MoveFirst
                                Do While Not myrsInfo.EOF
                                    frmWait.prgLoad.Value = myrsInfo.PercentPosition
                                    If ShowRealDates = 1 Then
                                        Set myrsDataTemp = dbXbuild.OpenRecordset("Select * From " & _
                                            "[Graph_Data] Order By [RealDate]")
                                    ElseIf ShowRealDates = 2 Then
                                        Set myrsDataTemp = dbXbuild.OpenRecordset("Select * From " & _
                                            "[Graph_Data] Order By Date")
                                    End If
                                    If Not (myrsDataTemp.EOF And myrsDataTemp.BOF) Then
                                        If ShowRealDates = 1 Then
                                            Set myrsDATa_Graph_select = dbXbuild.OpenRecordset("Select * From " & _
                                                "[Graph_Data] Where TheOrder = '" & myrsInfo.Fields("TheOrder").Value & "' " & _
                                                "And [" & ArrVarID(i, bb) & "(" & Y_ArrTableNames(i) & ") TRT " & _
                                                ArrTRNO(i, ss) & "] is Null ")
                                            
                                            If Not (myrsDATa_Graph_select.EOF And myrsDATa_Graph_select.BOF) Then
                                                myrsDATa_Graph_select.Edit
                                                myrsDATa_Graph_select.Fields(ArrVarID(i, bb) & "(" & Y_ArrTableNames(i) & ") TRT " & _
                                                    ArrTRNO(i, ss)) = myrsInfo.Fields(ArrVarID(i, bb)).Value
                                                myrsDATa_Graph_select.Update
                                            Else
                                                myrsDataTemp.AddNew
                                                myrsDataTemp.Fields("TheOrder") = myrsInfo.Fields("TheOrder").Value
                                                myrsDataTemp.Fields("RealDate") = TheDate(myrsInfo.Fields("TheOrder").Value)
                                                myrsDataTemp.Fields(ArrVarID(i, bb) & "(" & Y_ArrTableNames(i) & ") TRT " & _
                                                    ArrTRNO(i, ss)) = myrsInfo.Fields(ArrVarID(i, bb)).Value
                                                myrsDataTemp.Update
                                            End If
                                       Else
                                            Set myrsDATa_Graph_select = dbXbuild.OpenRecordset("Select * From " & _
                                                "[Graph_Data] Where Date ='" & myrsInfo.Fields("Date").Value & "' " & _
                                                "And [" & ArrVarID(i, bb) & "(" & Y_ArrTableNames(i) & ") TRT " & _
                                                ArrTRNO(i, ss) & "] is Null ")
                                            If Not (myrsDATa_Graph_select.EOF And myrsDATa_Graph_select.BOF) Then
                                                myrsDATa_Graph_select.Edit
                                                myrsDATa_Graph_select.Fields(ArrVarID(i, bb) & "(" & Y_ArrTableNames(i) & ") TRT " & _
                                                    ArrTRNO(i, ss)) = myrsInfo.Fields(ArrVarID(i, bb)).Value
                                                myrsDATa_Graph_select.Update
                                            Else
                                                myrsDataTemp.AddNew
                                                myrsDataTemp.Fields("Date") = myrsInfo.Fields("Date").Value
                                                myrsDataTemp.Fields(ArrVarID(i, bb) & "(" & Y_ArrTableNames(i) & ") TRT " & _
                                                    ArrTRNO(i, ss)) = myrsInfo.Fields(ArrVarID(i, bb)).Value
                                                myrsDataTemp.Update
                                            End If
                                       End If
                                   End If
                                    myrsDataTemp.MoveNext
                                    myrsInfo.MoveNext
                                Loop
                                frmWait.prgLoad.Value = 100
                            End If
                        End If
                    End If
                    Next bb
                
            Next ss
        Next i
    End If
    Exit Sub
    Err.Clear
    On Error Resume Next
    myrsEXPlist.Close
    myrsDataTemp.Close
    myrsDATA.Close
    myrsEXPinfo.Close
    myrsInfo.Close
Error1:
    Screen.MousePointer = vbDefault
    'MsgBox Err.Description & " in AddYColumns_GraphData_T_File/frmSelection."
    ErrorFound = True
    Unload frmWait2
    Unload frmWait3
Unload frmWait

    
End Sub

Public Sub AddXColumns_GraphData_T_File()
Dim myrsEXPlist As Recordset
Dim myrsDataTemp As Recordset
Dim myrsDATA As Recordset
Dim myrsEXPinfo As Recordset
Dim myrsInfo As Recordset
Dim i As Integer
Dim bb As Integer
Dim nn As Integer
Dim kk As Integer
Dim hh As Integer
Dim gg As Integer
Dim ss As Integer
Dim NumberYSelected As Integer
Dim prmSections() As Integer
Dim prmVariables() As Integer
Dim prmNumberTRNO() As Integer
Dim X_ArrVarID() As String
Dim X_ArrTRNO() As Integer
    On Error GoTo Error1
    i = 1
    bb = 1
    kk = 1
    ss = 1
    hh = 1
    gg = 1
    NumberYSelected = 1
    ReDim X_ArrVarID(10, 1000)
    ReDim prmVariables(10)
    'ReDim X_ArrSection(10)
    ReDim X_ArrTRNO(10, 1000)
    ReDim prmNumberTRNO(10)
    For i = 1 To T_File_X_NumberOutTables
            Set myrsEXPlist = dbXbuild.OpenRecordset("Select * From [" & _
                X_ArrTableNames(i) & "_List] Where  [Selected] = 'X' Or [Selected] = 'XY'")
            If Not (myrsEXPlist.EOF And myrsEXPlist.BOF) Then
                With myrsEXPlist
                    .MoveFirst
                    Do While Not .EOF
                        X_ArrVarID(i, hh) = !VariableID
                        NumberYSelected = NumberYSelected + 1
                        hh = hh + 1
                        .MoveNext
                    Loop
                    prmVariables(i) = hh - 1
                End With
            End If
    Next i
    NumberYSelected = NumberYSelected - 1
    If NumberYSelected > 0 Then
        DoEvents
    Else
        MsgBox "No variables have been selected."
        Exit Sub
    End If
    For i = 1 To T_File_X_NumberOutTables
                Set myrsEXPinfo = dbXbuild.OpenRecordset("Select * From [" & _
                    X_ArrTableNames(i) & "_File_Info] Where [Selected] = 'Y'")
                With myrsEXPinfo
                    If Not (.EOF And .BOF) Then
                        bb = 1
                        .MoveFirst
                        Do While Not .EOF
                            X_ArrTRNO(i, bb) = !TRNO
                            bb = bb + 1
                            .MoveNext
                        Loop
                    End If
                End With
            prmNumberTRNO(i) = bb - 1
    Next i
    Err.Clear
    On Error Resume Next
    myrsEXPlist.Close
    myrsEXPinfo.Close
    Err.Clear
    On Error GoTo Error1
    For i = 1 To T_File_X_NumberOutTables
            For bb = 1 To prmVariables(i)
                If X_ArrVarID(i, bb) <> "" Then
                    For ss = 1 To prmNumberTRNO(i)
                        If X_ArrTRNO(i, ss) <> 0 Then
                            dbXbuild.Execute "Alter Table [Graph_Data] Add Column [X_Axis_" & _
                                X_ArrVarID(i, bb) & _
                                "(" & X_ArrTableNames(i) & ") TRT " & X_ArrTRNO(i, ss) & "] Single"
                        End If
                    Next ss
                End If
            Next bb
    Next i
    
    Set myrsDATA = dbXbuild.OpenRecordset("Graph_Data")
     Set myrsDataTemp = dbXbuild.OpenRecordset("Graph_Data")
    NumberExpVariables_X = myrsDATA.Fields.Count - 3
    If T_File_Y_NumberOutTables > 0 Then
        Set myrsInfo = dbXbuild.OpenRecordset("Select * From [" & X_ArrTableNames(1) & _
            "] Where [TRNO] = " & X_ArrTRNO(1, 1))
        If Not (myrsInfo.BOF And myrsInfo.EOF) Then
            myrsInfo.MoveFirst
            Do While Not myrsInfo.EOF
                With myrsDATA
                    .AddNew
                    If ShowRealDates = 1 Then
                        .Fields("TheOrder") = myrsInfo.Fields("TheOrder").Value
                        .Fields("RealDate") = TheDate(myrsInfo.Fields("TheOrder").Value)
                    ElseIf ShowRealDates = 2 Then
                        .Fields("Date") = myrsInfo.Fields("Date").Value
                    End If
                    .Fields("X_Axis_" & X_ArrVarID(1, 1) & "(" & X_ArrTableNames(1) & ") TRT " & _
                        X_ArrTRNO(1, 1)) = myrsInfo.Fields(X_ArrVarID(1, 1)).Value
                    .Update
                End With
                myrsInfo.MoveNext
            Loop
        End If
    End If
    
   Dim myrsTest As Recordset             '?????????????????
        Dim ContinueFillingTable As Boolean
        For i = 1 To T_File_Y_NumberOutTables
            For ss = 1 To prmNumberTRNO(i)
                Set myrsInfo = dbXbuild.OpenRecordset("Select * From [" & X_ArrTableNames(i) & _
                        "] Where [TRNO] = " & X_ArrTRNO(i, ss))
                For bb = 1 To prmVariables(i)
                    Set myrsTest = dbXbuild.OpenRecordset("Select * From " & _
                        "[Graph_Data] Where [X_Axis_" & X_ArrVarID(i, bb) & "(" & X_ArrTableNames(i) & ") TRT " & _
                                                X_ArrTRNO(i, ss) & "] is Null ")
                   
                    With myrsTest
                        If Not (.EOF And .BOF) Then
                            ContinueFillingTable = True
                        Else
                            ContinueFillingTable = False
                        End If
                    End With
                If ContinueFillingTable = True Then

                        Dim myrsDATa_Graph_select As Recordset
                        If X_ArrVarID(i, bb) <> "" Then
                            If Not (myrsInfo.BOF And myrsInfo.EOF) Then
                                myrsInfo.MoveFirst
                                Do While Not myrsInfo.EOF
                                    frmWait.prgLoad.Value = myrsInfo.PercentPosition
                                    If ShowRealDates = 1 Then
                                        Set myrsDataTemp = dbXbuild.OpenRecordset("Select * From " & _
                                            "[Graph_Data] Order By [RealDate]")
                                    ElseIf ShowRealDates = 2 Then
                                        Set myrsDataTemp = dbXbuild.OpenRecordset("Select * From " & _
                                            "[Graph_Data] Order By Date")
                                    End If
                                    If Not (myrsDataTemp.EOF And myrsDataTemp.BOF) Then
                                        If ShowRealDates = 1 Then
                                            Set myrsDATa_Graph_select = dbXbuild.OpenRecordset("Select * From " & _
                                                "[Graph_Data] Where TheOrder = '" & myrsInfo.Fields("TheOrder").Value & "' " & _
                                                "And [X_Axis_" & X_ArrVarID(i, bb) & "(" & X_ArrTableNames(i) & ") TRT " & _
                                                X_ArrTRNO(i, ss) & "] is Null ")
                                            
                                            If Not (myrsDATa_Graph_select.EOF And myrsDATa_Graph_select.BOF) Then
                                                myrsDATa_Graph_select.Edit
                                                myrsDATa_Graph_select.Fields("X_Axis_" & X_ArrVarID(i, bb) & "(" & X_ArrTableNames(i) & ") TRT " & _
                                                    X_ArrTRNO(i, ss)) = myrsInfo.Fields(X_ArrVarID(i, bb)).Value
                                                myrsDATa_Graph_select.Update
                                            Else
                                                myrsDataTemp.AddNew
                                                myrsDataTemp.Fields("TheOrder") = myrsInfo.Fields("TheOrder").Value
                                                myrsDataTemp.Fields("RealDate") = TheDate(myrsInfo.Fields("TheOrder").Value)
                                                myrsDataTemp.Fields("X_Axis_" & X_ArrVarID(i, bb) & "(" & X_ArrTableNames(i) & ") TRT " & _
                                                    X_ArrTRNO(i, ss)) = myrsInfo.Fields(X_ArrVarID(i, bb)).Value
                                                myrsDataTemp.Update
                                            End If
                                       Else
                                            Set myrsDATa_Graph_select = dbXbuild.OpenRecordset("Select * From " & _
                                                "[Graph_Data] Where Date ='" & myrsInfo.Fields("Date").Value & "' " & _
                                                "And [X_Axis_" & X_ArrVarID(i, bb) & "(" & X_ArrTableNames(i) & ") TRT " & _
                                                X_ArrTRNO(i, ss) & "] is Null ")
                                            If Not (myrsDATa_Graph_select.EOF And myrsDATa_Graph_select.BOF) Then
                                                myrsDATa_Graph_select.Edit
                                                myrsDATa_Graph_select.Fields("X_Axis_" & X_ArrVarID(i, bb) & "(" & X_ArrTableNames(i) & ") TRT " & _
                                                    X_ArrTRNO(i, ss)) = myrsInfo.Fields(X_ArrVarID(i, bb)).Value
                                                myrsDATa_Graph_select.Update
                                            Else
                                                myrsDataTemp.AddNew
                                                myrsDataTemp.Fields("Date") = myrsInfo.Fields("Date").Value
                                                myrsDataTemp.Fields("X_Axis_" & X_ArrVarID(i, bb) & "(" & X_ArrTableNames(i) & ") TRT " & _
                                                    X_ArrTRNO(i, ss)) = myrsInfo.Fields(X_ArrVarID(i, bb)).Value
                                                myrsDataTemp.Update
                                            End If
                                       End If
                                   End If
                                    myrsDataTemp.MoveNext
                                    myrsInfo.MoveNext
                                Loop
                                frmWait.prgLoad.Value = 100
                            End If
                        End If
                    End If
                    Next bb
                
            Next ss
        Next i
    'End If
    Exit Sub
Error1:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description & " in AddXColunms_GaphData_T_File/frmSelection."
    
End Sub

Public Sub FindSelectedOutTables()
Dim bn As Integer
Dim i As Integer
Dim myrsEXPinfo As Recordset

ReDim Y_ArrTableNames(NumberOutTables + 1)
ReDim X_ArrTableNames(NumberOutTables + 1)

    On Error GoTo Error1
    bn = 1
    For i = 1 To NumberOutTables
        Set myrsEXPinfo = dbXbuild.OpenRecordset("Select * From [" & _
                    OUTTableNames(i) & "_List] Where [Selected] = 'X' Or Selected = 'XY'")
        With myrsEXPinfo
            If Not (.EOF And .BOF) Then
                X_ArrTableNames(bn) = OUTTableNames(i)
                bn = bn + 1
            End If
            .Close
        End With
    Next i
    T_File_X_NumberOutTables = bn - 1
    
    bn = 1
    For i = 1 To NumberOutTables
        Set myrsEXPinfo = dbXbuild.OpenRecordset("Select * From [" & _
                    OUTTableNames(i) & "_List] Where [Selected] = 'Y' Or Selected = 'XY'")
        With myrsEXPinfo
            If Not (.EOF And .BOF) Then
                Y_ArrTableNames(bn) = OUTTableNames(i)
                bn = bn + 1
            End If
            .Close
        End With
    Next i
    T_File_Y_NumberOutTables = bn - 1
    Err.Clear
    On Error Resume Next
    myrsEXPinfo.Close
    Exit Sub
Error1:
    MsgBox Err.Description & " in FindSelectedOUTTables."
    
End Sub

Public Function Y_Variable_Selected() As Boolean
Dim myrsOUT As Recordset
Dim NotEmpty As Integer
Dim i As Integer
Dim TheOutFile As String
    
    On Error GoTo Error1
    NotEmpty = 0
    If FileType = "SUM" Then
        Set myrsOUT = dbXbuild.OpenRecordset("Select * From [Evaluate_List] Where Selected = 'Y'")
        With myrsOUT
            If Not (.EOF And .BOF) Then
                NotEmpty = 1 + NotEmpty
            End If
        End With
        If NotEmpty <> 1 Then
            Y_Variable_Selected = False
            myrsOUT.Close
        Else
            Y_Variable_Selected = True
            myrsOUT.Close
        End If
        Exit Function
    End If
    
    
    For i = 1 To NumberOutTables
        If FileType = "OUT" Then
            If ExpData_vs_Simulated = 1 Then
                TheOutFile = Mid(OUTTableNames(i), 1, Len(OUTTableNames(i)) - 4) & "_List_EXPOUT"
            Else
                TheOutFile = Mid(OUTTableNames(i), 1, Len(OUTTableNames(i)) - 4) & "_List"
            End If
        Else 'T-File
            TheOutFile = OUTTableNames(i) & "_List"
        End If
       Err.Clear
        On Error Resume Next

        Set myrsOUT = dbXbuild.OpenRecordset("Select * From [" & TheOutFile _
            & "] Where Selected = 'Y' Or Selected = 'XY'")
        With myrsOUT
        If Not (.EOF And .BOF) Then
            NotEmpty = 1 + NotEmpty
            End If
        End With
            
    Next i
    If NotEmpty = 0 Then
        Y_Variable_Selected = False
        myrsOUT.Close
        Exit Function
    Else
        Y_Variable_Selected = True
    End If
    myrsOUT.Close
    
    Exit Function
Error1:
    MsgBox Err.Description & " in Y_Variable_Selected/frmSelection."
End Function





Public Sub Create_New_List_Table()
Dim i As Integer
Dim k As Integer
Dim h As Integer
Dim ListVariablesOUT() As String
Dim myrsInfo As Recordset
Dim TheOutFileName As String
Dim NumberOfListVariables As Integer
Dim NumberOf_EXP_ListVariables As Integer
Dim LisOUT_ExpVariables() As String
Dim ExperimentalFiles() As String
Dim myNumberExpFiles As Integer
Dim Number_OUT_ExpVariables As Integer
Dim nn As Integer
Dim dd As Integer
Dim bb As Integer
Dim myrsNew As Recordset
Dim myrsDetail As Recordset
Dim myrs As Recordset
    On Error GoTo Error1
    'ErrorFile = False
    For i = 1 To NumberOutTables
        k = 1
        
        TheOutFileName = Mid(OUTTableNames(i), 1, Len(OUTTableNames(i)) - 4)
        Err.Clear
        On Error Resume Next
        dbXbuild.Execute ("Drop Table [" & TheOutFileName & "_List_EXPOUT]")
        Err.Clear
        On Error GoTo Error1
        dbXbuild.Execute ("Create Table [" & TheOutFileName & "_List_EXPOUT]( " _
            & "[VariableID] Text, [VariableDescription] Text, [Selected] Text)")
         'Debug.Print TheOutFileName & "_List_EXPOUT]( "
        Set myrs = dbXbuild.OpenRecordset(TheOutFileName & "_List")
        Set myrsInfo = dbXbuild.OpenRecordset("Select Distinct [ExperimentID],[CropID] From " _
            & " [" & TheOutFileName & "_File_Info]")
        With myrs
            If Not (.EOF And .BOF) Then
                .MoveLast
                NumberOfListVariables = .RecordCount + 1
                .MoveFirst
                If !VariableID <> "DAP" Then
                    ReDim ListVariablesOUT(NumberOfListVariables)
                    Do While Not .EOF
                        ListVariablesOUT(k) = !VariableID
                        k = k + 1
                        .MoveNext
                    Loop
                Else
                    
                End If
            End If
            .Close
        End With
        nn = 1
        With myrsInfo
            If Not (.EOF And .BOF) Then
                .MoveLast
                ReDim ExperimentalFiles(.RecordCount + 1)
                .MoveFirst
                Do While Not (.EOF)
                    ExperimentalFiles(nn) = !ExperimentID & " " & !CropID & "T"
                    nn = nn + 1
                    .MoveNext
                Loop
            End If
            .Close
        End With
        myNumberExpFiles = nn - 1
        For h = 1 To myNumberExpFiles
            bb = 1
            Err.Clear
            On Error Resume Next
            dbXbuild.Execute "Drop Table DistinctNames"
            Err.Clear
            On Error GoTo Error1
            dbXbuild.Execute "Create Table DistinctNames (VariableID Text, VariableDescription Text)"
            Err.Clear
            On Error GoTo 22
            Set myrsInfo = dbXbuild.OpenRecordset(ExperimentalFiles(h))
            
            With myrsInfo
                If Not (.EOF And .BOF) Then
                    NumberOf_EXP_ListVariables = .Fields.Count - 1
                    ReDim LisOUT_ExpVariables(NumberOf_EXP_ListVariables)
                    For k = 3 To NumberOf_EXP_ListVariables
                        For dd = 1 To NumberOfListVariables
                            
                            If UCase(ListVariablesOUT(dd)) = UCase(.Fields(k).Name) Then
                                LisOUT_ExpVariables(bb) = ListVariablesOUT(dd)
                                bb = bb + 1
                            End If
                        Next dd
                    Next k
                    Number_OUT_ExpVariables = bb - 1
                    bb = 1
                    
                    If Number_OUT_ExpVariables > 0 Then
                        For bb = 1 To Number_OUT_ExpVariables
                            Set myrsNew = dbXbuild.OpenRecordset(TheOutFileName & "_List_EXPOUT")
                                myrsNew.AddNew
                                myrsNew.Fields("VariableID").Value = LisOUT_ExpVariables(bb)
                                Set myrsDetail = dbXbuild.OpenRecordset("Select * From [DATA] " _
                                    & "Where [Code] = '" & LisOUT_ExpVariables(bb) & "'")
                                If Not (myrsDetail.EOF And myrsDetail.BOF) Then
                                    myrsDetail.MoveFirst
                                    myrsNew.Fields("VariableDescription").Value = _
                                        myrsDetail.Fields("Description").Value
                                Else
                                    myrsNew.Fields("VariableDescription").Value = LisOUT_ExpVariables(bb)
                                End If
                                myrsNew.Update
                                myrsNew.Close
                                
                        Next bb
                        
                    Else
                        MsgBox "The file" & " " & TheOutFileName & " " & "does not have data for Exp/Sim plotting."
                        ExpData_vs_Simulated = 0
                        ShowX_Axis = 1
                        prmShowLine = 1
                        'frmSelection.Show
                        'ErrorFile1 = True
                        Exit Sub
                    End If
                End If
            End With
22:        Next h
    
    Dim myrsDist As Recordset
    Err.Clear
    On Error GoTo Error1
    Set myrsNew = dbXbuild.OpenRecordset("Select Distinct VariableID, " & _
        " VariableDescription From " & TheOutFileName & "_List_EXPOUT")
    dbXbuild.Execute "Delete * From DistinctNames"
    Set myrsDist = dbXbuild.OpenRecordset("DistinctNames")
    With myrsNew
       If Not (.EOF And .BOF) Then
          .MoveFirst
           Do While Not .EOF
              myrsDist.AddNew
              myrsDist!VariableID = !VariableID
              myrsDist!VariableDescription = !VariableDescription
              myrsDist.Update
              .MoveNext
           Loop
       End If
    End With
                                                            
    myrsNew.Close
    dbXbuild.Execute "Delete * From " & TheOutFileName & "_List_EXPOUT"
    Set myrsNew = dbXbuild.OpenRecordset(TheOutFileName & "_List_EXPOUT")

    With myrsDist
        If Not (.EOF And .BOF) Then
           .MoveFirst
           Do While Not .EOF
              myrsNew.AddNew
              myrsNew!VariableID = !VariableID
              myrsNew!VariableDescription = !VariableDescription
              myrsNew.Update
              .MoveNext
            Loop
        End If
    End With
    myrsDist.Close
    myrsNew.Close
    
    Next i
    
    Exit Sub
    Err.Clear
    On Error Resume Next
    myrsNew.Close
    myrsDetail.Close
    myrs.Close
Error1:
    Unload frmWait2
Unload frmWait3
MsgBox Err.Description & " in Create_New_List_Table/frmMain."
Error2:
End Sub





Public Sub ReloadData()
Dim X As Integer
Dim myrsOUTDATA As Recordset
Dim myrsOUT As Recordset
Dim myrsTempNames As Recordset
Dim myrsTemExp As Recordset
Dim myrsFile As Recordset
Dim myrsSection As Recordset
Dim j As Integer
Dim SpacePlace() As Integer
Dim m As Integer
Dim FileToOpen As String
Dim TableName1 As String
Dim TableName2 As String
Dim h As Integer
Dim BackSlashPlace As Integer
Dim OpenFileDir As String
Dim CannotOpenFileName As String
Dim TheOutFileName As String
Dim prmMyExperiment As String
Dim ii As Integer
Dim prmTheOutFileName As String
Dim prmMyExperiment1 As String
Dim s As Integer
Dim FilesString As String
Dim BackShlashPlace1 As Integer
Dim BackShlashPlace2 As Integer
Dim BackShlashPlace3 As Integer
Dim ArrOutFileTemp() As String
Dim bb As Integer
Dim nn As Integer
Dim mm As Integer
Dim ArrayTRNO() As Integer
Dim i As Integer
Dim NumberNotOpenedFiles As Integer
Dim myrsTreat As Recordset
Dim myExt As String

    Screen.MousePointer = vbHourglass
    ReDim Preserve ArrayT_Tables(1)
    CurrentFileSelected = ""
    On Error Resume Next
    NumberNotOpenedFiles = 0
    For i = 1 To NumberOfOUTfiles
        prmTheOutFileName = Mid(ArrOutFile(i), 1, Len(ArrOutFile(i)) - 4)
        dbXbuild.Execute "Drop Table [" & prmTheOutFileName & "_List]"
        dbXbuild.Execute "Drop Table [" & prmTheOutFileName & "_File_Info]"
        dbXbuild.Execute "Drop Table [" & prmTheOutFileName & "_OUT]"
    Next i
    For i = 1 To NumberOfExpFiles
        'Err.Clear
        prmMyExperiment = Replace(Mid(ArrayExpFiles(i), InStr(1, _
            ArrayExpFiles(i), "\") + 1), ".", " ")
            dbXbuild.Execute "Drop Table [" & prmMyExperiment & "]"
            dbXbuild.Execute "Drop Table [" & prmMyExperiment & "_File_Info]"
            dbXbuild.Execute "Drop Table [" & prmMyExperiment & "_List]"
    Next i
    If FileType = "SUM" Then
        dbXbuild.Execute "Drop Table [Evaluate_List]"
        dbXbuild.Execute "Drop Table [Evaluate_SUM]"
    End If
    
    NumberOfOUTfiles = 0
    NumberOfExpFiles = 0
    Err.Clear
    CropName = FrmSelectionName
    
    myExt = UCase(Mid(CropName, InStrRev(CropName, ".") + 1))
    frmWait3.Show
    DoEvents
    If UCase(Mid(CropName, InStrRev(CropName, "\") + 1)) = "EVALUATE.OUT" Or _
    Right(UCase(Mid(CropName, InStrRev(CropName, "\") + 1)), 3) = "OEV" Then
        FileType = "SUM"
    ElseIf UCase(Mid(CropName, InStrRev(CropName, ".") + 1)) = "OUT" Then
        FileType = "OUT"
    ElseIf UCase(Mid(CropName, InStrRev(CropName, ".") + 3)) = "T" Then
        FileType = "T-file"
    Else
        FileType = "OUT"
    End If



    Err.Clear
    On Error Resume Next
    For ii = 1 To NumberOfExpFiles
        prmMyExperiment1 = Replace(Mid(ArrayExpFiles(ii), InStr(1, _
            ArrayExpFiles(ii), "\") + 1), ".", " ")
        dbXbuild.Execute "Drop Table [" & prmMyExperiment1 & "]"
    Next ii
    Err.Clear
    On Error GoTo Error2
    m = 0
    '''''''''''''
    If InStrRev(CropName, "\") <> 0 Then
        BackShlashPlace1 = InStr(CropName, "\")
        BackShlashPlace2 = InStr(BackShlashPlace1 + 1, CropName, "\")
        If BackShlashPlace2 <> 0 Then
            Do Until BackShlashPlace2 = 0
                BackShlashPlace3 = BackShlashPlace2
                BackShlashPlace2 = InStr(BackShlashPlace2 + 1, CropName, "\")
            Loop
        End If
            BackSlashPlace = InStr(BackShlashPlace3 + 1, CropName, " ")
    End If
    ''''''''''''''''
    Dim ss As Integer
    ss = InStr(CropName, ".")
    If InStr(ss + 1, CropName, ".") = 0 Then
        m = 1
        ReDim Preserve SpacePlace(m)
        SpacePlace(m) = InStrRev(CropName, "\") + 1
        FilesString = Trim(Mid(CropName, SpacePlace(m)))
    Else
        FilesString = Trim(Mid(CropName, BackSlashPlace + 1))
        m = 0
        ReDim SpacePlace(2)
        SpacePlace(0) = 1
        For X = 0 To Len(FilesString)
            ReDim Preserve SpacePlace(X + 2)
            SpacePlace(X + 1) = InStr(SpacePlace(X) + 1, FilesString, " ")
            m = m + 1
            If SpacePlace(X + 1) = 0 Then Exit For
        Next X
    End If
    FilesName_Open = FilesString
    If FileType = "SUM" Then
    'If UCase(FilesString) = "EVALUATE.SUM" Then
        myReadFile CropName
        CreateEvaluateTable
        CreateEvaluateListTable
        If CreateEvaluateListTable_DontDoThis = True Then
            Screen.MousePointer = vbDefault
            Unload frmWait2
            Unload frmWait3
            Exit Sub
        End If
        NumberOutTables = 1
        ReDim OUTTableNames(2)
         
        OUTTableNames(1) = UCase(Mid(CropName, InStrRev(CropName, "\") + 1))
       ' OUTTableNames(1) = "EVALUATE.OUT"
        'OpenFileDir = Replace(Trim(UCase(CropName)), "\EVALUATE.OUT", "")
        OpenFileDir = Replace(Trim(UCase(CropName)), "\" & OUTTableNames(1), "")
 
         
        
        'OUTTableNames(1) = "EVALUATE.OUT"
       ' OpenFileDir = Replace(Trim(Mid(CropName, 1, BackSlashPlace)), " ", "")
        If Right$(OpenFileDir, 1) = "\" Then
            OpenFileDir = OpenFileDir
        Else
            OpenFileDir = Trim(OpenFileDir) & "\"
        End If
        DirectoryToPreview = OpenFileDir
        prmShowLine = 0
         frmOpenFileShown = True
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    NumberOfOUTfiles = m
    SpacePlace(NumberOfOUTfiles) = Len(FilesString) + 1
    ReDim ArrOutFileTemp(1 To NumberOfOUTfiles + 1)
    If NumberOfOUTfiles = 1 Then
        ArrOutFileTemp(1) = FilesString
    Else
        For j = 0 To NumberOfOUTfiles - 1
            ArrOutFileTemp(j + 1) = Trim(Mid(FilesString, SpacePlace(j), SpacePlace(j + 1) - SpacePlace(j)))
        Next j
    End If
    bb = 0
    For j = 1 To NumberOfOUTfiles
        If FileType = "OUT" Then
            ReDim Preserve ArrOutFile(j + 1)
            ArrOutFile(j) = ArrOutFileTemp(j)
            bb = bb + 1
        If ObservedWasOpened = True Then
            prmShowLine = 1
            ObservedWasOpened = False
        End If
    End If
        If FileType = "T-file" Then
            ReDim Preserve ArrOutFile(j + 1)
            ArrOutFile(j) = ArrOutFileTemp(j)
            bb = bb + 1
            prmShowLine = 0
            If ShowRealDates = 0 Then
                ShowRealDates = 1
            End If
        End If
        ObservedWasOpened = True
    Next j
    NumberOfOUTfiles = bb
    Err.Clear
    CannotOpenFileName = ""
    OpenFileDir = Replace(Trim(Mid(CropName, 1, BackSlashPlace)), " ", "")
    If Right$(OpenFileDir, 1) = "\" Then
        OpenFileDir = OpenFileDir
    Else
        OpenFileDir = Trim(OpenFileDir) & "\"
    End If
    DirectoryToPreview = OpenFileDir
    If FileType = "OUT" Then
        Err.Clear
        On Error Resume Next
        dbXbuild.Execute "Drop Table [NamesTemp]"
        Err.Clear
        On Error GoTo Error2
        dbXbuild.Execute "Create Table [NamesTemp] ([Experiment] Text (25))"
        Set myrsTempNames = dbXbuild.OpenRecordset("NamesTemp")
        h = 1
        For j = 1 To NumberOfOUTfiles
            frmWait3.prgLoad.Visible = True
            frmWait3.prgLoad.Value = 100 / NumberOfOUTfiles * (j - 1)
            FileToOpen = OpenFileDir & ArrOutFile(j)
            TheOutFileName = Mid(ArrOutFile(j), 1, Len(ArrOutFile(j)) - 4)
            TableName1 = TheOutFileName & "_File_Info"
            TableName2 = TheOutFileName & "_OUT"
            If NumberOfOUTfiles = 1 Then
                FileToOpen = CropName
                 DirectoryToPreview = Mid(CropName, 1, Len(CropName) - Len(ArrOutFile(1)))
            End If
          '  Debug.Print Trim(DirectoryToPreview) & "ll"
            'myReadFile FileToOpen
          ' Dim mmm As String
          ' mmm = Replace(FileToOpen, "\ ", "\")
            'myReadFile mmm
          ' Debug.Print mmm
           ' myReadFile FileToOpen
            myReadFile Trim(DirectoryToPreview) & TheOutFileName & ".out"
            
            
            frmWait2.prgLoad.Value = (100 / NumberOfOUTfiles) * (j - 1)
            '*****
            frmWait2.prgLoad.Visible = True
            '****
            dbXbuild.Execute "Create Table [" & TheOutFileName & "_File_Info] ([CropID] Text (2), [Crop] Text (10), [CultivarID] Text (8)," _
            & " [Cultivar] Text (25), [ExperimentID] Text (8), [ExpDescription] Text (225), [TRNO] Integer, " & _
            "[TreatmentDescription] Text (225), [RunNumber] Integer, [RunDescription] Text (225))"

            CreateOutDataTable (TheOutFileName)
            Call FillCreateOUTTable(TableName1, TheOutFileName, (100 / NumberOfOUTfiles) * (j - 1) + 1)
            
            Create_OneOutTable (TheOutFileName)
'Call FillCreateOUTTable(TableName1, TheOutFileName, 100 / NumberOfOUTfiles * (j - 1) + 1)
            dbXbuild.Execute "Alter Table [" & TheOutFileName & "_File_Info] Add Column [Selected] Text"
            Set myrsOUT = dbXbuild.OpenRecordset("Select Distinct [ExperimentID], [CropID],[Crop] From [" & _
                TableName1 & "]")
            Set myrsOUTDATA = dbXbuild.OpenRecordset(TableName2)
            If myrsOUTDATA.EOF And myrsOUTDATA.BOF Then
                myrsOUTDATA.Close
                myrsOUT.Close
                dbXbuild.Execute "Drop Table " & TableName1
                dbXbuild.Execute "Drop Table " & TableName2
                NumberNotOpenedFiles = NumberNotOpenedFiles + 1
                CannotOpenFileName = CannotOpenFileName & "  " & TheOutFileName
                If InStr(FrmSelectionName, TheOutFileName & ".out") <> 0 Then
                    FrmSelectionName = Replace(FrmSelectionName, TheOutFileName & ".out", "")
                End If
                If InStr(FrmSelectionName, TheOutFileName & ".OUT") <> 0 Then
                    FrmSelectionName = Replace(FrmSelectionName, TheOutFileName & ".OUT", "")
                End If
            Else
                With myrsOUT
                    If Not (.EOF And .BOF) Then
                        .MoveFirst
                        Do While Not .EOF
                            myrsTempNames.AddNew
                            myrsTempNames.Fields("Experiment") = !ExperimentID & "." & _
                                !CropID & "T"
                            myrsTempNames.Update
                            .MoveNext
                        Loop
                    Else
                        If InStr(FrmSelectionName, FileToOpen & ".out") <> 0 Then
                            FrmSelectionName = Replace(FrmSelectionName, FileToOpen & ".out", "")
                        End If
                        If InStr(FrmSelectionName, FileToOpen & ".OUT") <> 0 Then
                            FrmSelectionName = Replace(FrmSelectionName, FileToOpen & ".OUT", "")
                        End If
                        MsgBox "Cannot locate necessary information. Error in file" & " " & _
                            FileToOpen
                    End If
                End With
                myrsOUT.Close
                CreateOUTListTable (ArrOutFile(j))
                h = h + 1
                ReDim Preserve OUTTableNames(h)
                OUTTableNames(h - 1) = ArrOutFile(j)
            End If
        Next j
        frmWait3.prgLoad.Value = 100
        NumberOutTables = h - 1
        myrsTempNames.Close
        Set myrsTempNames = dbXbuild.OpenRecordset("Select Distinct [Experiment] From [NamesTemp]")
        With myrsTempNames
            If Not (.EOF And .BOF) Then
                .MoveLast
                NumberOfExpFiles = .RecordCount
                .MoveFirst
                ReDim ArrayExpFiles(NumberOfExpFiles + 1)
                For j = 1 To NumberOfExpFiles
                    ArrayExpFiles(j) = !Experiment
                    If Not .EOF Then .MoveNext
                Next j
            End If
        End With
        myrsTempNames.Close
        If CannotOpenFileName <> "" Then
            If NumberOutTables = 0 And NumberOfExpFiles = 0 Then
                Screen.MousePointer = vbDefault
                Unload frmWait3
            End If
        End If
        s = 0
        For j = 1 To NumberOfExpFiles
        If Dir(Trim(DirectoryToPreview) & ArrayExpFiles(j)) <> "" Then
                myReadFile Trim(DirectoryToPreview) & ArrayExpFiles(j)
                prmMyExperiment = ArrayExpFiles(j)
                ErrorWithExp = False
                CreateTables prmMyExperiment
                CreateExpTable prmMyExperiment
                If ErrorWithExp = True Then NumberOfExpFiles = 0
                s = s + 1
            End If
        Next j
    End If
    Err.Clear
    On Error Resume Next
    dbXbuild.Execute "Drop Table [NamesTemp] "
    NumberOfExpFiles = s
    Err.Clear
    
    If FileType = "T-file" Then
        'Selected are T-files. We are using arrays to store t-files info:
         'NumberOfOUTFiles=Number of t-files opened
         'ArrOutFile(j)- t-files
        h = 1
        Unload frmWait3
        NumberOfExpFiles = NumberOfOUTfiles
        ReDim ArrayExpFiles(NumberOfExpFiles + 1)
        For j = 1 To NumberOfOUTfiles
            
            ArrayExpFiles(j) = ArrOutFile(j)
           ' myReadFile DirectoryToPreview & ArrOutFile(j)
            myReadFile CropName
            CreateTables ArrOutFile(j)
            CreateExpTable ArrOutFile(j)
            Set myrsTemExp = dbXbuild.OpenRecordset(Replace(ArrOutFile(j), ".", " "))
            If myrsTemExp.EOF And myrsTemExp.BOF Then
                myrsTemExp.Close
                dbXbuild.Execute "Drop Table [" & Replace(ArrOutFile(j), ".", " ") & "]"
            Else
                CreateExpList ArrOutFile(j)
                h = h + 1
                ReDim Preserve OUTTableNames(h)
                OUTTableNames(h - 1) = ArrOutFile(j)
            End If
        Next j
        NumberOutTables = h - 1
        NumberOfExpFiles = NumberOfOUTfiles
        Err.Clear
        Dim TreatmentDescription As String
        On Error GoTo Error1
        For j = 1 To NumberOfOUTfiles
            dbXbuild.Execute "Create Table [" & _
                Replace(ArrOutFile(j), ".", " ") & "_File_Info] ([TRNO] Integer, [Selected] Text, [Description] Text)"
            nn = 1
            dbXbuild.Execute "Delete * From [" & _
                Replace(ArrOutFile(j), ".", " ") & "_File_Info] "
                Set myrsSection = dbXbuild.OpenRecordset("Select Distinct TRNO From [" & _
                    Replace(ArrOutFile(j), ".", " ") & "]")
                FileType = "X"
                If Dir(Mid(Replace(ArrOutFile(j), " ", "."), 1, Len(ArrOutFile(j)) - 1) & "x") <> "" Then
                    'myReadFile DirectoryToPreview & (Mid(Replace(ArrOutFile(j), " ", "."), 1, Len(ArrOutFile(j)) - 1) & "x")
                    myReadFile Mid(CropName, 1, Len(CropName) - 1) & "X"
                End If
                    CreateTempTreatmentTable
                FileType = "T-file"
                Set myrsFile = dbXbuild.OpenRecordset(Replace(ArrOutFile(j), ".", " ") & "_File_Info")
                If Not (myrsSection.BOF And myrsSection.EOF) Then
                    myrsSection.MoveFirst
                    Do While Not myrsSection.EOF
                    myrsFile.AddNew
                    Set myrsTreat = dbXbuild.OpenRecordset("Select * From [Treatment_Names] " & _
                        "Where TRNO = " & myrsSection.Fields("TRNO").Value)
                    If Not (myrsTreat.EOF And myrsTreat.BOF) Then
                        myrsTreat.MoveFirst
                        TreatmentDescription = myrsTreat!Description
                    Else
                        TreatmentDescription = ""
                    End If
                    myrsFile.Fields("TRNO").Value = myrsSection.Fields("TRNO").Value
                    myrsFile.Fields("Description").Value = TreatmentDescription
                    myrsFile.Update
                    myrsSection.MoveNext
                    Loop
                    
                End If
            myrsFile.Close
            myrsSection.Close
            myrsTreat.Close
            dbXbuild.Execute "Drop Table [Treatment_Names]"
3:        Next j
    End If
    frmOpenFileShown = True
    FileIsClosed = False
    Screen.MousePointer = vbDefault
Exit Sub
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number = 3078 Then GoTo 3
    Unload frmWait3
Error2:
    MsgBox Err.Description & " in ReloadData/frmSelection."
    Screen.MousePointer = vbDefault
End Sub


Public Sub AddYColumns_GraphData_SUM_File()
Dim myrsEXPlist As Recordset
Dim myrsDATA As Recordset
Dim myrsInfo As Recordset
Dim bb As Integer
Dim hh As Integer
Dim NumberYSelected As Integer
Dim ArrVarID() As String
Dim ArrFieldName() As Integer
    On Error GoTo Error1
    bb = 1
    hh = 1
    ErrorFound = False
    NumberYSelected = 1
    Set myrsEXPlist = dbXbuild.OpenRecordset("Select * From [Evaluate_List] Where [Selected] = 'Y'")
    If Not (myrsEXPlist.EOF And myrsEXPlist.BOF) Then
        With myrsEXPlist
            .MoveFirst
            Do While Not .EOF
                ReDim Preserve ArrVarID(hh + 1)
                ArrVarID(hh) = !VariableID
                NumberYSelected = NumberYSelected + 1
                hh = hh + 1
                .MoveNext
            Loop
        End With
    End If
    NumberYSelected = NumberYSelected - 1
    If NumberYSelected > 0 Then
        DoEvents
    Else
        MsgBox "No variables have been selected."
        Exit Sub
    End If
    For bb = 1 To NumberYSelected
        If ArrVarID(bb) <> "" Then
            dbXbuild.Execute "Alter Table [Graph_Data] Add Column [X_Axis_" & _
                ArrVarID(bb) & "] Single"
        End If
    Next bb
    For bb = 1 To NumberYSelected
        If ArrVarID(bb) <> "" Then
            dbXbuild.Execute "Alter Table [Graph_Data] Add Column [" & _
                ArrVarID(bb) & "] Single"
        End If
    Next bb
    
    Set myrsDATA = dbXbuild.OpenRecordset("Graph_Data")
    Set myrsInfo = dbXbuild.OpenRecordset("Evaluate_SUM")
            If Not (myrsInfo.BOF And myrsInfo.EOF) Then
                myrsInfo.MoveFirst
                Do While Not myrsInfo.EOF
                    myrsDATA.AddNew
                    For bb = 1 To NumberYSelected
                        If ArrVarID(bb) <> "" Then
                            'myrsDATA.Fields("X_Axis_" & ArrVarID(bb)).Value = myrsInfo.Fields(ArrVarID(bb) & "P").Value
                            'myrsDATA.Fields(ArrVarID(bb)).Value = myrsInfo.Fields(ArrVarID(bb) & "O").Value
                            myrsDATA.Fields("X_Axis_" & ArrVarID(bb)).Value = myrsInfo.Fields(ArrVarID(bb) & "S").Value
                            myrsDATA.Fields(ArrVarID(bb)).Value = myrsInfo.Fields(ArrVarID(bb) & "M").Value
                        End If
                    Next bb
                    myrsDATA.Update
                    myrsInfo.MoveNext
                Loop
            End If
    Exit Sub
    Err.Clear
    On Error Resume Next
    myrsEXPlist.Close
    myrsDATA.Close
    myrsInfo.Close
Error1:
    Screen.MousePointer = vbDefault
   ' MsgBox Err.Description & " in AddYColumns_GraphData_SUM_File/frmSelection."
    ErrorFound = True
    Unload frmWait2
    Unload frmWait3
Unload frmWait

End Sub

Public Function CalculateDays(prmDaysNumberData As Integer, prmOrderDATA As String)
    On Error GoTo Error1
    Dim myRealDate_prmOrderExp
    Dim myRealDate_prmOrderDATA
    Dim prmPlantingDate_forEXP
    myRealDate_prmOrderDATA = TheDate(prmOrderDATA)
    prmPlantingDate_forEXP = DateAdd("d", prmDaysNumberData * (-1), Format(myRealDate_prmOrderDATA, "mm/dd/yyyy"))
   ' CalculateDays = DateDiff("d", myRealDate_prmOrderExp, prmPlantingDate_forEXP)
    CalculateDays = prmPlantingDate_forEXP
    Exit Function
Error1:     MsgBox Err.Description & " in CalculateDays."
End Function

Public Function RemoveDaysAfterPlantingToTheOrder(prmCalculateDays, prmOrderDATA As String)
    Dim myRealDate_prmOrderDATA As Date
    Dim valStringDays As Single
    Dim StringDays As String
    Dim prmPlantingDate_forEXP As Single
    On Error GoTo Error1
    myRealDate_prmOrderDATA = TheDate(prmOrderDATA)
    prmPlantingDate_forEXP = DateDiff("d", Format(prmCalculateDays, "mm/dd/yyyy"), Format(myRealDate_prmOrderDATA, "mm/dd/yyyy"))
    RemoveDaysAfterPlantingToTheOrder = prmPlantingDate_forEXP
    Exit Function
Error1:     MsgBox Err.Description & " in RemoveDaysAfterPlantingToTheOrder."
End Function

Public Sub FillOUTGraphData_RealDates1()
Dim VariableCount As Integer
Dim myrsTest As Recordset
Dim myrsGraphData  As Recordset
Dim myrsOUTvsEXP  As Recordset
Dim myrsInfo  As Recordset
Dim myrsDateTemp  As Recordset
Dim myrsGraphSelect  As Recordset
Dim fff As Integer
    
    On Error GoTo Error1
    Set myrsGraphData = dbXbuild.OpenRecordset("Graph_Data")
    Set myrsOUTvsEXP = dbXbuild.OpenRecordset("Select * From Exp_OUT Where Axis = 'Y'")
    VariableCount = 0
    ReDim FirstPlantDate(2)
    ReDim LastPlantDate(2)

    
    With myrsOUTvsEXP
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Set myrsTest = dbXbuild.OpenRecordset("Select * From [" & _
                .Fields("OUT_File").Value & "_File_Info] Where [RunNumber] = " _
                & .Fields("Run").Value & " And [TRNO] = " _
                & .Fields("TRNO").Value)
            If Not (myrsTest.EOF And myrsTest.BOF) Then
                If CDAYexists = False Then

                    Set myrsInfo = dbXbuild.OpenRecordset("Select * From [" & _
                        .Fields("OUT_File").Value & "_OUT] Where [RunNumber] = " _
                        & .Fields("Run").Value & " Order by DAS")
                    
                Else
                    Set myrsInfo = dbXbuild.OpenRecordset("Select * From [" & _
                        .Fields("OUT_File").Value & "_OUT] Where [RunNumber] = " _
                        & .Fields("Run").Value & " Order by CDAY")
                
                End If
                
                    If Not (myrsInfo.EOF And myrsInfo.BOF) Then
                        myrsInfo.MoveFirst
                        Do While Not (myrsInfo.EOF)
                            myrsGraphData.AddNew
                            myrsGraphData.Fields(.Fields("Variable").Value & _
                                "(" & .Fields("OUT_File") & ") Run " & .Fields("RUN")).Value = _
                                myrsInfo.Fields(.Fields("Variable")).Value
                                
                                myrsGraphData.Fields("TheOrder").Value = myrsInfo.Fields("Date").Value
                                If FirstPlantDate(1) = "12:00:00 AM" Then FirstPlantDate(1) = TheDate(myrsInfo.Fields("Date").Value)
                                LastPlantDate(1) = TheDate(myrsInfo.Fields("Date").Value)

                            myrsGraphData.Update
                            myrsInfo.MoveNext
                        Loop
                    End If
                    .MoveLast
                    
                End If
                VariableCount = .RecordCount
            End If
        End With
        myrsGraphData.MoveLast
       ' LastPlantDate(1) = TheDate(myrsGraphData.Fields("TheOrder").Value)
        
        'Add the rest of out variables/data adjusting the dates
        ReDim Preserve FirstPlantDate(VariableCount + 1)
        ReDim Preserve LastPlantDate(VariableCount + 1)
        Dim m As Integer
        m = 1
        If VariableCount > 1 Then
            With myrsOUTvsEXP
                .MoveFirst
                .MoveNext
                Set myrsTest = dbXbuild.OpenRecordset("Select * From [" & _
                    .Fields("OUT_File").Value & "_File_Info] Where [RunNumber] = " _
                    & .Fields("Run").Value & " And [TRNO] = " _
                    & .Fields("TRNO").Value)
                    frmWait.prgLoad.Visible = True
                    If Not (myrsTest.EOF And myrsTest.BOF) Then
                        Do While Not .EOF
                            'm = m + 1
                            frmWait.prgLoad.Value = .PercentPosition
                            
                            If CDAYexists = False Then
                                Set myrsInfo = dbXbuild.OpenRecordset("Select * From [" & _
                                    .Fields("OUT_File").Value & "_OUT] Where [RunNumber] = " _
                                    & .Fields("Run").Value & " Order By DAS")
                            Else
                                Set myrsInfo = dbXbuild.OpenRecordset("Select * From [" & _
                                    .Fields("OUT_File").Value & "_OUT] Where [RunNumber] = " _
                                    & .Fields("Run").Value & " Order By CDAY")
                            
                            End If
                            
                            
                            If Not (myrsInfo.EOF And myrsInfo.BOF) Then
                                myrsGraphData.MoveFirst
                                myrsInfo.MoveFirst
                                Do While Not (myrsInfo.EOF)
                                    
                                   Set myrsDateTemp = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [TheOrder] = '" _
                                        & myrsInfo.Fields("DATE") & "'")
                                        If FirstPlantDate(m + 1) = "12:00:00 AM" Then FirstPlantDate(m + 1) = TheDate(myrsInfo.Fields("DATE"))
                                            
                                            LastPlantDate(m + 1) = TheDate(myrsInfo.Fields("DATE"))
                                       ' End If
                                    If myrsDateTemp.EOF And myrsDateTemp.BOF Then
                                        myrsGraphData.AddNew
                                        myrsGraphData.Fields(.Fields("Variable").Value & _
                                            "(" & .Fields("OUT_File") & ") Run " & .Fields("RUN")).Value = _
                                            myrsInfo.Fields(.Fields("Variable")).Value
                                        myrsGraphData.Fields("TheOrder").Value = _
                                            myrsInfo.Fields("DATE").Value
                                        myrsGraphData.Update
                                    Else
                                        Set myrsGraphSelect = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [TheOrder] = '" _
                                            & myrsInfo.Fields("Date") & "'")
                                        myrsGraphSelect.Edit
                                        myrsGraphSelect.Fields(.Fields("Variable").Value & _
                                            "(" & .Fields("OUT_File") & ") Run " & .Fields("RUN")).Value = _
                                            myrsInfo.Fields(.Fields("Variable")).Value
                                        myrsGraphSelect.Update
                                    End If
                                    myrsInfo.MoveNext
                                Loop
                                myrsInfo.MovePrevious
                               ' LastPlantDate(m + 1) = TheDate(myrsInfo.Fields("Date").Value)
                                m = m + 1
                            End If
                            
                            .MoveNext
                        Loop
                        frmWait.prgLoad.Value = 100
                    End If
                    myrsTest.Close
                End With
            End If

    Err.Clear
    On Error Resume Next
    myrsTest.Close
    myrsGraphData.Close
    myrsOUTvsEXP.Close
    myrsInfo.Close
    myrsDateTemp.Close
    myrsGraphSelect.Close
    Exit Sub
Error1:
    Screen.MousePointer = vbDefault
    'MsgBox Err.Description
    ErrorFound = True
      '  Unload frmWait2
    Unload frmWait3
Unload frmWait2
Unload frmWait
    'MsgBox Err.Description & " in FillOUTGraphData_RealDates/frmSelection."
End Sub

Public Sub FillOUT_X_GraphData_RealDates1()
Dim myrsTest As Recordset
Dim myrsGraphData As Recordset
Dim myrsOUTvsEXP As Recordset
Dim myrsInfo As Recordset
Dim myrsDateTemp As Recordset
Dim myrsGraphSelect As Recordset

    On Error GoTo Error1
    Set myrsGraphData = dbXbuild.OpenRecordset("Graph_Data")
    Set myrsOUTvsEXP = dbXbuild.OpenRecordset("Select * From Exp_OUT Where Axis = 'X'")
    With myrsOUTvsEXP
        If Not (.EOF And .BOF) Then
           DoEvents
         Else
            myrsOUTvsEXP.Close
            myrsGraphData.Close
            Exit Sub
         End If
        .MoveFirst
        Set myrsTest = dbXbuild.OpenRecordset("Select * From [" & _
            .Fields("OUT_File").Value & "_File_Info] Where [RunNumber] = " _
            & .Fields("Run").Value & " And [TRNO] = " _
            & .Fields("TRNO").Value)
        If Not (myrsTest.EOF And myrsTest.BOF) Then
            Do While Not .EOF
                Set myrsInfo = dbXbuild.OpenRecordset("Select * From [" & _
                    .Fields("OUT_File").Value & "_OUT] Where [RunNumber] = " _
                    & .Fields("Run").Value)
                If Not (myrsInfo.EOF And myrsInfo.BOF) Then
                    myrsGraphData.MoveFirst
                    myrsInfo.MoveFirst
                    Do While Not (myrsInfo.EOF)
                        Set myrsDateTemp = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [TheOrder] = '" _
                            & myrsInfo.Fields("Date") & "'")
                        If myrsDateTemp.EOF And myrsDateTemp.BOF Then
                            myrsGraphData.AddNew
                            myrsGraphData.Fields("X_Axis_" & .Fields("Variable").Value & _
                                "(" & .Fields("OUT_File") & ") Run " & .Fields("RUN")).Value = _
                                myrsInfo.Fields(.Fields("Variable")).Value
                            myrsGraphData.Fields("TheOrder").Value = myrsInfo.Fields("Date").Value
                            myrsGraphData.Update
                        Else
                            Set myrsGraphSelect = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [TheOrder] = '" _
                                & myrsInfo.Fields("Date") & "'")
                            myrsGraphSelect.Edit
                            myrsGraphSelect.Fields("X_Axis_" & .Fields("Variable").Value & _
                                "(" & .Fields("OUT_File") & ") Run " & .Fields("RUN")).Value = _
                                myrsInfo.Fields(.Fields("Variable")).Value
                            myrsGraphSelect.Update
                        End If
                            myrsInfo.MoveNext
                    Loop
                End If
                .MoveNext
            Loop
        End If
    End With
    
    Err.Clear
    On Error Resume Next
    myrsTest.Close
    myrsGraphData.Close
    myrsOUTvsEXP.Close
    myrsInfo.Close
    myrsDateTemp.Close
    myrsGraphSelect.Close
    Exit Sub
Error1:
Screen.MousePointer = vbDefault
'MsgBox Err.Description & " in FillOUT_X_GraphData_RealDate/frmSelection."
   ' MsgBox "Error with experimental data."
    ErrorFound = True
    Unload frmWait2
    Unload frmWait
Unload frmWait3
End Sub

Public Sub AddExpDataToGraphData_YearDate1()
Dim myrsOUTvsEXP As Recordset
Dim myrsDateTemp As Recordset
Dim myrsGraphSelect As Recordset
Dim myrsGraphData As Recordset
Dim myrsOUTvsEXP_Sel As Recordset
Dim ExpTableName As String
Dim myrsExp As Recordset
Dim myrsPlantingDay As Recordset
Dim myDaysAfterPlanting As String
Dim myrstempData As Recordset

    On Error GoTo Error1
    Set myrsGraphData = dbXbuild.OpenRecordset("Graph_Data")
    myrsGraphData.MoveFirst
    Set myrsOUTvsEXP_Sel = dbXbuild.OpenRecordset("Select * From [Exp_OUT] Where " & _
    "[ExpCheck] = 'YES' And [Axis] = 'X'")
    With myrsOUTvsEXP_Sel
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do While Not .EOF
                ExpTableName = .Fields("ExperimentID").Value & _
                    " " & .Fields("CropID").Value & "T"
                Set myrsExp = dbXbuild.OpenRecordset("Select * From [" & _
                    ExpTableName & "] Where [TRNO] = " & .Fields("TRNO") & " And " & _
                    .Fields("Variable") & " Is Not Null")
                If Not (myrsExp.EOF And myrsExp.BOF) Then
                    myrsGraphData.MoveFirst
                    myrsExp.MoveFirst
                    Do While Not (myrsExp.EOF)
                            Set myrsDateTemp = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [TheOrder] = '" _
                                & myrsExp.Fields("TheOrder") & "' and " & _
                                "[X_Exp_Axis_" & .Fields("Variable").Value & _
                                "(" & ExpTableName & ") TRT " & .Fields("TRNO").Value & _
                                "/" & .Fields("Run").Value & "] Is Null")
                        If myrsDateTemp.EOF And myrsDateTemp.BOF Then
                            myrsGraphData.AddNew
                            myrsGraphData.Fields("X_Exp_Axis_" & .Fields("Variable").Value & _
                                "(" & ExpTableName & ") TRT " & .Fields("TRNO").Value & "/" & .Fields("Run").Value) = _
                            myrsExp.Fields(.Fields("Variable")).Value
                            myrsGraphData.Fields("TheOrder") = myrsExp.Fields("TheOrder")
                            myrsGraphData.Update
                        Else
                            myrsDateTemp.Edit
                            myrsDateTemp.Fields("X_Exp_Axis_" & .Fields("Variable").Value & _
                                "(" & ExpTableName & ") TRT " & .Fields("TRNO").Value & "/" & .Fields("Run").Value) = _
                                myrsExp.Fields(.Fields("Variable")).Value
                            myrsDateTemp.Update
                        End If
                        myrsExp.MoveNext
                    Loop
                End If
                .MoveNext
            Loop
        End If
    End With
    ''
    Set myrsOUTvsEXP_Sel = dbXbuild.OpenRecordset("Select * From [Exp_OUT] Where " & _
    "[ExpCheck] = 'YES' And [Axis] = 'Y'")
    With myrsOUTvsEXP_Sel
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do While Not .EOF
                ExpTableName = .Fields("ExperimentID").Value & _
                    " " & .Fields("CropID").Value & "T"
                Set myrsExp = dbXbuild.OpenRecordset("Select * From [" & _
                    ExpTableName & "] Where [TRNO] = " & .Fields("TRNO") & " And [" & _
                    .Fields("Variable").Value & "] Is Not Null")
                If Not (myrsExp.EOF And myrsExp.BOF) Then
                    myrsGraphData.MoveFirst
                    myrsExp.MoveFirst
                    Do While Not (myrsExp.EOF)
                            Set myrsDateTemp = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [TheOrder] = '" _
                                & myrsExp.Fields("TheOrder").Value & "' And [" & .Fields("Variable").Value & _
                                "(" & ExpTableName & ") TRT " & .Fields("TRNO").Value & "/" & _
                                .Fields("Run").Value & "] Is Null")
                        If myrsDateTemp.EOF And myrsDateTemp.BOF Then
                            myrsGraphData.AddNew
                            myrsGraphData.Fields(.Fields("Variable").Value & _
                                "(" & ExpTableName & ") TRT " & .Fields("TRNO").Value & "/" & _
                                .Fields("Run").Value).Value = _
                                myrsExp.Fields(.Fields("Variable")).Value
                            myrsGraphData.Fields("TheOrder").Value = myrsExp.Fields("TheOrder")
                            myrsGraphData.Update
                        Else
                            myrsDateTemp.Edit
                            myrsDateTemp.Fields(.Fields("Variable").Value & _
                                "(" & ExpTableName & ") TRT " & .Fields("TRNO").Value & "/" & _
                                .Fields("Run").Value) = myrsExp.Fields(.Fields("Variable")).Value
                            myrsDateTemp.Update
                        End If
                        myrsExp.MoveNext
                    Loop
                End If
                .MoveNext
            Loop
        End If
    End With
   ' dbXbuild.Execute "Delete * From [Graph_Data] Where [TheOrder] Is Null"
    On Error Resume Next
    Err.Clear
    myrsOUTvsEXP.Close
    myrsDateTemp.Close
    myrsGraphSelect.Close
    myrsGraphData.Close
    myrsExp.Close
    myrsOUTvsEXP_Sel.Close
    
    Exit Sub
Error1:
Screen.MousePointer = vbDefault
'MsgBox Err.Description & " in AddExpDataToGraphData_YearDate/frmSelection."
    'MsgBox "Error with experimental data."
    ErrorFound = True
    Unload frmWait2
    Unload frmWait3
Unload frmWait
End Sub


Public Sub FindExcelDataOUT_Time()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim prmNumberOfVariables As Integer
Dim myrsOUT As Recordset
Dim myrs As Recordset
Dim mySQLString As String
Dim prmNumberOfRecords As Integer
    On Error GoTo Error1
    ExcelDataExist = True
    'Set myrs = dbXbuild.OpenRecordset("Graph_Data")
    If ShowRealDates = 0 Then
        Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where TheOrder Is Not Null Order By Date")
        'Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Order By Date")
    ElseIf ShowRealDates = 1 Then
        Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where TheOrder Is Not Null Order By RealDate")
    ElseIf ShowRealDates = 2 Then
        Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where TheOrder Is Not Null Order By Date")
    End If
    With myrs
        If Not (.EOF And .BOF) Then
            .MoveLast
            prmNumberOfRecords = .RecordCount
            ReDim ExcelVariable(.Fields.Count + 1)
            ReDim textVariables(.Fields.Count + 1)
            If NumberExpVariables_Y > 0 Then
                prmNumberOfVariables = .Fields.Count - 1
            Else
                prmNumberOfVariables = .Fields.Count - 2
            End If
            mySQLString = ""
            For i = 2 To NumberOUTVariables + 1
                ExcelVariable(i) = RemoveSymbols(SelectedVariable(.Fields(i + 1).Name))
                textVariables(i) = RemoveSymbols(.Fields(i + 1).Name)
                mySQLString = mySQLString & "[" & .Fields(i + 1).Name & "] Is Not Null Or "
            Next i
        End If
    End With
    mySQLString = Mid(mySQLString, 1, Len(mySQLString) - 4)
    
    If ShowRealDates = 0 Then
        Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where (" & mySQLString & _
            ") And TheOrder Is Not Null Order By Date")
        'Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Order By Date")
        ExcelVariable(1) = "Day"
        textVariables(1) = "Day"
    ElseIf ShowRealDates = 1 Then
        Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where (" & mySQLString & _
            ") And TheOrder Is Not Null Order By RealDate")
        'Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Order By RealDate")
        ExcelVariable(1) = "Date"
        textVariables(1) = "Date"
    ElseIf ShowRealDates = 2 Then
        Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Order By Date")
        ExcelVariable(1) = "Day"
        textVariables(1) = "Day"
    End If
    Dim TheOnlyOUTFileld As String
    Dim myrsTime As Recordset
    With myrs
        TheOnlyOUTFileld = .Fields(3).Name
        ReDim ExcelDataPlot(1 To prmNumberOfRecords + 1, 1 To NumberOUTVariables + NumberExpVariables_Y + 2)
        i = 1
            
            '''''''''''''''''''''''''
                For i = 2 To Excel_Number_Y_OUT_Variables + 1
                    .MoveFirst
                    j = 1
                    Do While Not .EOF
                        Dim nn
                        nn = .Fields(i + 1).Name
                        
                       ' If IsNull(.Fields(i + 1).Value) = False Then
                            ExcelDataPlot(j, i) = .Fields(i + 1).Value
                            j = j + 1
                      '  End If
                        .MoveNext
                    Loop
                Next i
                Dim MyString As String
                MyString = ""
                For i = 2 To Excel_Number_Y_OUT_Variables + 2
                    MyString = MyString & "Or [" & .Fields(i + 1).Name & "] is not Null "
                Next i
                
                MyString = Trim(Mid(MyString, 3))
                If Excel_Number_Y_OUT_Variables = 0 Then
                    MyString = "[" & TheOnlyOUTFileld & "] is not Null "
                End If

                For i = Excel_Number_Y_OUT_Variables + 2 To .Fields.Count - 2
                    .MoveFirst
                    j = 1
                    Do While Not .EOF
                        ExcelDataPlot(j, i) = .Fields(i + 1).Value
                        j = j + 1
                        .MoveNext
                    Loop
                Next i

                
                myrs.MoveFirst
                j = 1
                Do While Not myrs.EOF
                    If ShowRealDates = 1 Then
                        ExcelDataPlot(j, 1) = myrs.Fields("RealDate").Value
                    ElseIf ShowRealDates = 0 Then
                        ExcelDataPlot(j, 1) = myrs.Fields("Date").Value
                    ElseIf ShowRealDates = 2 Then
                        ExcelDataPlot(j, 1) = myrs.Fields("Date").Value
                    End If
                    j = j + 1
                    myrs.MoveNext
                Loop
                myrs.Close
            
            
        End With
            '''''''''''''''''''


    'Experimental Variables
        If NumberExpVariables_Y > 0 Then
            If NumberOUTVariables > 0 Then
            If ShowRealDates = 0 Then
                ExcelVariable(NumberOUTVariables + 2) = "Day"
                'ReDim Preserve textVariables(NumberOUTVariables + 2)
                textVariables(NumberOUTVariables + 2) = "Day"
            ElseIf ShowRealDates = 1 Then
                ExcelVariable(NumberOUTVariables + 2) = "Date"
                textVariables(NumberOUTVariables + 2) = "Date"
            ElseIf ShowRealDates = 2 Then
                ExcelVariable(NumberOUTVariables + 2) = "Day"
                textVariables(NumberOUTVariables + 2) = "Day"
            End If
            End If
         End If
        Dim ExpVariableName() As String
        Dim mb As Integer
        ReDim ExpVariableName(1)
        mb = 1
        Set myrs = dbXbuild.OpenRecordset("Graph_Data")
        With myrs
            If Not (.EOF And .BOF) Then
                mySQLString = ""
                For i = NumberOUTVariables + 2 To .Fields.Count - 2
                    ExcelVariable(i + 1) = RemoveSymbols(SelectedVariable(.Fields(i + 1).Name))
                    'ReDim Preserve textVariables(i + 1)
                    textVariables(i + 1) = RemoveSymbols(.Fields(i + 1).Name)
                    mySQLString = mySQLString & "[" & .Fields(i + 1).Name & "] Is Not Null Or "
                    ReDim Preserve ExpVariableName(mb + 1)
                    ExpVariableName(mb) = .Fields(i + 1).Name
                    mb = mb + 1
                Next i
            End If
        End With
    
    Dim TotalNumberExpNames As Integer
    TotalNumberExpNames = mb - 1
    Dim myrsExp As Recordset
    If TotalNumberExpNames > 0 Then
        For mb = 1 To TotalNumberExpNames
            Set myrsExp = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [" & _
            ExpVariableName(mb) & "] Is Not Null")
            With myrsExp
                If .EOF And .BOF Then
                    'MsgBox "Experimental data for" & " " & ExpVariableName(mb) & " " & "is missing.", vbExclamation
                    
                End If
                .Close
            End With
        Next mb
    End If
    If mySQLString <> "" Then
        mySQLString = Mid(mySQLString, 1, Len(mySQLString) - 4)
      
        If ShowRealDates = 0 Then
            Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where (" & mySQLString & _
                ") And TheOrder Is Not Null Order By Date")
        ElseIf ShowRealDates = 1 Then
            Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where (" & mySQLString & _
            ") And TheOrder Is Not Null Order By RealDate")
        ElseIf ShowRealDates = 2 Then
            Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where " & mySQLString & " Order By Date Where")
        End If
     
     Else
        If ShowRealDates = 0 Then
            Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where TheOrder Is Not Null Order By Date")
        ElseIf ShowRealDates = 1 Then
            Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where TheOrder Is Not Null Order By RealDate")
        ElseIf ShowRealDates = 2 Then
            Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Order By Date Where")
        End If
     
     End If
        Dim ccc
        If myrs.EOF And myrs.BOF Then
            If ShowRealDates = 0 Then
                Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where TheOrder Is Not Null Order By Date")
            ElseIf ShowRealDates = 1 Then
                Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where TheOrder Is Not Null Order By RealDate")
            ElseIf ShowRealDates = 2 Then
                Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Order By Date Where")
            End If
        End If
        
        myrs.MoveLast
        ccc = myrs.RecordCount
        Dim m As Integer
        With myrs
            For i = NumberOUTVariables + 2 To .Fields.Count - 1
                .MoveFirst
                j = 1
                Do While Not .EOF
                    ExcelDataPlot(j, i) = .Fields(i).Value
                    Dim mm
                    mm = .Fields(i).Name
                    If j = 85 Then
                    Dim bv
                    bv = 4
                    End If
                    j = j + 1
                    .MoveNext
                Loop
                For m = j To prmNumberOfRecords
                    ExcelDataPlot(m, i) = Null
                Next m
            Next i
            
            .MoveFirst
            j = 1
            Do While Not .EOF
                If ShowRealDates = 1 Then
                    ExcelDataPlot(j, NumberOUTVariables + 2) = .Fields(2).Value
                ElseIf ShowRealDates = 0 Then
                    ExcelDataPlot(j, NumberOUTVariables + 2) = .Fields(1).Value
                ElseIf ShowRealDates = 2 Then
                    ExcelDataPlot(j, NumberOUTVariables + 2) = .Fields(1).Value
                End If
                j = j + 1
                .MoveNext
            Loop
        End With

'''''''''''''''''''''''
        Dim New_Variable() As String
        Open App.path & "\" & "Graph_Data.txt" For Output As #5 'Len = 800
        ReDim New_Variable(prmNumberOfVariables)
        For i = 1 To prmNumberOfVariables
            New_Variable(i) = textVariables(i)
            If Len(New_Variable(i)) > 4 Then
                New_Variable(i) = Trim(Mid(New_Variable(i), 1, InStr(New_Variable(i), "(") - 1)) & _
                    "(" & Trim(Mid(New_Variable(i), InStr(New_Variable(i), ")") + 1)) & ")"

                If Len(New_Variable(i)) = 11 Then
                    New_Variable(i) = New_Variable(i) & "  "
                End If
                If Len(New_Variable(i)) = 12 Then
                    New_Variable(i) = New_Variable(i) & " "
                End If
                
                If Len(New_Variable(i)) = 13 Then
                    New_Variable(i) = New_Variable(i) & ""
                End If
            ElseIf Len(New_Variable(i)) = 3 Then
                New_Variable(i) = New_Variable(i) & "          "
            ElseIf Len(New_Variable(i)) = 4 Then
                New_Variable(i) = New_Variable(i) & "         "
            End If
        
        Next i
        Dim MyLableString As String
        MyLableString = ""
        
        
        For i = 1 To prmNumberOfVariables
            MyLableString = MyLableString & " " & New_Variable(i)
        Next i

        Print #5, ""
        Print #5, "File(s): " & CropName
        Print #5, ""
        Print #5, MyLableString
        
        For j = 1 To ExcelRecNumber
            For i = 1 To prmNumberOfVariables
                Print #5, IIf(IsNull(ExcelDataPlot(j, i)) = True, " ", ExcelDataPlot(j, i)),
            Next i
            Print #5, ""
        Next j
        Close 5

''''''''''''''''''''''''
    Err.Clear
    On Error Resume Next
    myrs.Close
    myrsOUT.Close
    Exit Sub
Error1:
   ' If Dir(App.path & "\" & "Graph_Data.txt") <> "" Then Kill App.path & "\" & "Graph_Data.txt"
    ExcelDataExist = False
    MsgBox Err.Description & " in FindExcelDataOUT_Time." & Err.Number
Exit Sub
NoData:
    ExcelDataExist = False

End Sub



Public Sub FillNullDataInOut0()
Dim myNumberOfVariables As Integer
Dim i As Integer
Dim j As Integer
Dim a()
Dim d()
Dim NumberOfMyrs_Records As Integer
Dim FirstEvementExists As Boolean
Dim ai_Next
Dim i_First As Integer
Dim i_Second As Integer
Dim k As Integer
Dim delay
Dim g As Integer
Dim prmNumberExpVariables As Integer
Dim NumberOUTVariables
Dim bb As Integer
Dim cc As Integer
Dim myrs As Recordset
    On Error GoTo Error1
    Set myrs = dbXbuild.OpenRecordset("Select * from [Exp_OUT] " & _
        "Where [ExpCheck] = 'Yes'")
    With myrs
        If Not (.EOF And .BOF) Then
            .MoveLast
            prmNumberExpVariables = .RecordCount
        Else
            prmNumberExpVariables = 0
        End If
        .Close
    End With
    If ShowRealDates = 1 Then
        Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Order By RealDate")
        '''''''''''''''''''''''''''
        With myrs
            If Not (.EOF And .BOF) Then
                .MoveLast
                NumberOfMyrs_Records = .RecordCount
            Else
                Exit Sub
            End If
        End With
        myNumberOfVariables = myrs.Fields.Count - 1
        
        NumberOUTVariables = myNumberOfVariables - prmNumberExpVariables
        For j = 3 To NumberOUTVariables
            ReDim a(NumberOfMyrs_Records + 10)
            ReDim d(NumberOfMyrs_Records + 10)
            With myrs
                .MoveFirst
                i = 1
                Do While Not (.EOF)
                    a(i) = .Fields(j).Value
                    d(i) = .Fields(2).Value
                    i = i + 1
                    .MoveNext
                Loop
            End With
            For i = 1 To NumberOfMyrs_Records - 1
                If DateDiff("d", Format(d(i), "mm/dd/yyyy"), Format(d(i + 1), "mm/dd/yyyy")) = 0 Then
                    If IsNull(a(i)) = False Then
                        a(i + 1) = a(i)
                    End If
                End If
            Next i
            For i = NumberOfMyrs_Records To 2 Step -1
                If DateDiff("d", Format(d(i), "mm/dd/yyyy"), Format(d(i - 1), "mm/dd/yyyy")) = 0 Then
                    If IsNull(a(i)) = False Then
                        a(i - 1) = a(i)
                    End If
                End If
            Next i
            Dim mm As Integer
            FirstEvementExists = False
            
                
            For mm = 1 To NumberOfMyrs_Records - 2
                If IsNull(a(mm)) = False Then
                    Exit For
                End If
            Next mm
           
            With myrs
                .MoveFirst
                i = 1
                Do While Not .EOF
                    .Edit
                    If IsNull(.Fields(j).Value) = True Then
                        .Fields(j).Value = a(i)
                    End If
                    .Update
                    i = i + 1
                    .MoveNext
                Loop
            End With
        Next j
        ''''''''''''''''''''''''''''
    ElseIf ShowRealDates = 0 Or ShowRealDates = 2 Then
        Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Order By Date")
        With myrs
            If Not (.EOF And .BOF) Then
                .MoveLast
                NumberOfMyrs_Records = .RecordCount
                '.MoveFirst
            Else
                Exit Sub
            End If
        End With
        myNumberOfVariables = myrs.Fields.Count - 1
        
        NumberOUTVariables = myNumberOfVariables - prmNumberExpVariables
        For j = 3 To NumberOUTVariables
            ReDim a(NumberOfMyrs_Records + 10)
            ReDim d(NumberOfMyrs_Records + 10)
            With myrs
                .MoveFirst
                i = 1
                Do While Not (.EOF)
                    a(i) = .Fields(j).Value
                    d(i) = .Fields(1).Value
                    i = i + 1
                    .MoveNext
                Loop
            End With
            For i = 1 To NumberOfMyrs_Records - 1
                If d(i) = d(i + 1) Then
                    If IsNull(a(i)) = False Then
                        a(i + 1) = a(i)
                    End If
                End If
            Next i
            For i = NumberOfMyrs_Records To 2 Step -1
                If d(i) = d(i - 1) Then
                    If IsNull(a(i)) = False Then
                        a(i - 1) = a(i)
                    End If
                End If
            Next i
            For mm = 1 To NumberOfMyrs_Records - 2
                If IsNull(a(mm)) = False Then
                    Exit For
                End If
            Next mm
           
            
            
            With myrs
                .MoveFirst
                i = 1
                Do While Not .EOF
                    .Edit
                    If IsNull(.Fields(j).Value) = True Then
                        .Fields(j).Value = a(i)
                    End If
                    .Update
                    i = i + 1
                    .MoveNext
                Loop
            End With
        Next j
    End If
    Exit Sub
Error1:    MsgBox Err.Description & " in FillNullDataInOut0."
End Sub




Public Sub FindExcelData()
Dim myrs As Recordset
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim prmNumberOfVariables As Integer
Dim myrsOUT As Recordset
    
    On Error GoTo Error1
    ExcelDataExist = True
    If FileType <> "SUM" Then
        If ShowRealDates = 0 Then
            Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where TheOrder IS Not Null Order By Date")
        ElseIf ShowRealDates = 1 Then
            Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where TheOrder IS Not Null Order By RealDate")
        ElseIf ShowRealDates = 2 Then
            Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where TheOrder IS Not Null Order By Date")
        End If

    
    Else
        Set myrs = dbXbuild.OpenRecordset("Graph_Data")
    End If
    
    With myrs
        If Not (.EOF And .BOF) Then
            prmNumberOfVariables = .Fields.Count - 2
            Excel_Number_X_OUT_Variables = (.Fields.Count - 3) / 2
            Excel_Number_Y_OUT_Variables = (.Fields.Count - 3) / 2
            ExcelRecNumber = .RecordCount
            ReDim ExcelVariable(.Fields.Count - 2)
            ReDim textVariables(.Fields.Count - 2)
                For i = 1 To .Fields.Count - 3
                    ExcelVariable(i) = RemoveSymbols(SelectedVariable(.Fields(i + 2).Name))
                    textVariables(i) = RemoveSymbols(.Fields(i + 2).Name)
                'OxanaOxana
                Next i
            .MoveLast
            ReDim ExcelDataPlot(1 To .RecordCount + 1, 1 To .Fields.Count - 1)
            
            
            If ExpData_vs_Simulated = 0 Then
                i = 2
                For i = 2 To .Fields.Count - 2
                    .MoveFirst
                    j = 1
                    Do While Not .EOF
                        ExcelDataPlot(j, i) = Round(.Fields(i + 1).Value, 4)
                       ' If .Fields(i + 1).Value = 25387 Then
                         '   Dim mmm
                           ' mmm = 4
                       ' End If
                        
                        j = j + 1
                        .MoveNext
                    Loop
                Next i
                .MoveFirst
                j = 1
                Do While Not .EOF
                    If ShowRealDates = 1 Then
                        ExcelDataPlot(j, 1) = Round(.Fields(2).Value, 4)
                    ElseIf ShowRealDates = 0 Then
                        ExcelDataPlot(j, 1) = Round(.Fields(1).Value, 4)
                        If j = 57 Then
                           ' mmm = 6
                        End If
                    ElseIf ShowRealDates = 2 Then
                        ExcelDataPlot(j, 1) = Round(.Fields(1).Value, 4)
                    End If
                    j = j + 1
                    .MoveNext
                Loop
            
            Dim mm
            mm = 4
            Else
                j = 1
                .MoveFirst
                Do While Not .EOF
                    If ShowRealDates = 1 Then
                        ExcelDataPlot(j, 1) = Round(.Fields(2).Value, 4)
                    ElseIf ShowRealDates = 0 Then
                        ExcelDataPlot(j, 1) = Round(.Fields(1).Value, 4)
                    End If
                    j = j + 1
                    .MoveNext
                Loop
                i = 2
                For i = 2 To .Fields.Count - 2
                    .MoveFirst
                    j = 1
                    Do While Not .EOF
                        ExcelDataPlot(j, i) = Round(.Fields(i + 1).Value, 4)
                        j = j + 1
                        .MoveNext
                    Loop
                Next i
            End If
        Else
            GoTo NoData
        End If
        If ShowX_Axis <> 1 And ExpData_vs_Simulated = 0 And FileType = "OUT" Then
            Set myrsOUT = dbXbuild.OpenRecordset("Select * From [Exp_OUT] Where " & _
            "[Axis] = 'X' ")
            
            
            With myrsOUT
                If Not (.EOF And .BOF) Then
                    .MoveLast
                    Excel_Number_X_OUT_Variables = .RecordCount
                    .MoveFirst
                Else
                    GoTo NoData
                End If
                .Close
            End With
            
             Set myrsOUT = dbXbuild.OpenRecordset("Select * From [Exp_OUT] Where " & _
            "[Axis] = 'X' And [ExpCheck] = 'Yes'  ")
            
            
            With myrsOUT
                If Not (.EOF And .BOF) Then
                    .MoveLast
                    Excel_Number_X_Exp_Variables = .RecordCount
                End If
                .Close
            End With
           
            
            
            Set myrsOUT = dbXbuild.OpenRecordset("Select * From [Exp_OUT] Where " & _
            "[Axis] = 'Y' And [ExpCheck] = 'Yes' ")
            With myrsOUT
                If Not (.EOF And .BOF) Then
                    .MoveLast
                If Excel_Number_X_Exp_Variables <> 0 Then
                    Excel_Number_Y_Exp_Variables = .RecordCount
                Else
                    Excel_Number_Y_Exp_Variables = 0
                End If
                End If
                .Close
            End With
            Set myrsOUT = dbXbuild.OpenRecordset("Select * From [Exp_OUT] Where " & _
            "[Axis] = 'Y' ")
            With myrsOUT
                If Not (.EOF And .BOF) Then
                    .MoveLast
                    Excel_Number_Y_OUT_Variables = .RecordCount
                Else
                    GoTo NoData
                End If
                .Close
            End With
        End If
        If ExpData_vs_Simulated = 1 Then
            Set myrsOUT = dbXbuild.OpenRecordset("Exp_OUT")
            With myrsOUT
                If Not (.EOF And .BOF) Then
                    .MoveLast
                    Excel_Number_Y_OUT_Variables = .RecordCount
                    Excel_Number_X_OUT_Variables = Excel_Number_Y_OUT_Variables
                Else
                    GoTo NoData
                End If
            End With
        End If
    End With
    If ShowExperimentalData = 0 Then
        Excel_Number_X_Exp_Variables = 0
        Excel_Number_Y_Exp_Variables = 0
    End If
   ' Exit Sub
    
  Dim New_Variable() As String
    If ShowX_Axis <> 1 Or FileType = "T-file" Or FileType = "SUM" Then
        Open App.path & "\" & "Graph_Data.txt" For Output As #4 'Len = 800
            'Print #4, "Variables:"
            'Print #4, "Var1 = " & "Date"
        ReDim New_Variable(prmNumberOfVariables)
        For i = 1 To prmNumberOfVariables
            'Print #4, "Var" & (i) & " = " & Replace(ExcelVariable(i - 1), "X_Exp_Axis_", "")
            New_Variable(i) = Replace(textVariables(i), "X_Exp_Axis_", "")
            New_Variable(i) = Replace(textVariables(i), "X_Axis_", "")
            
            '''''''''''''''''
            
            If Len(New_Variable(i)) > 4 Then
                New_Variable(i) = Trim(Mid(New_Variable(i), 1, InStr(New_Variable(i), "(") - 1)) & _
                    "(" & Trim(Mid(New_Variable(i), InStr(New_Variable(i), ")") + 1)) & ")"

                If Len(New_Variable(i)) = 11 Then
                    New_Variable(i) = New_Variable(i) & "  "
                End If
                If Len(New_Variable(i)) = 12 Then
                    New_Variable(i) = New_Variable(i) & " "
                End If
                
                If Len(New_Variable(i)) = 13 Then
                    New_Variable(i) = New_Variable(i) & ""
                End If
            ElseIf Len(New_Variable(i)) = 3 Then
                New_Variable(i) = New_Variable(i) & "          "
            ElseIf Len(New_Variable(i)) = 4 Then
                New_Variable(i) = New_Variable(i) & "         "
            End If
        
        Next i
        Dim MyLableString As String
        
        
        If FileType = "SUM" Then
            MyLableString = ""
        Else
            MyLableString = " Date        "

        End If
        
        For i = 1 To prmNumberOfVariables
            MyLableString = MyLableString & " " & New_Variable(i)
        Next i
     
        
        Print #4, ""
        
        Print #4, "File(s): " & CropName
        Print #4, ""

        Print #4, MyLableString
            
        
        For j = 1 To ExcelRecNumber
            If FileType = "SUM" Then
                For i = 2 To prmNumberOfVariables
                    Print #4, IIf(IsNull(ExcelDataPlot(j, i)) = True, " ", ExcelDataPlot(j, i)),
                Next i
            Else
                For i = 1 To prmNumberOfVariables
                    Print #4, IIf(IsNull(ExcelDataPlot(j, i)) = True, " ", ExcelDataPlot(j, i)),
                Next i
            End If
            Print #4, ""
        Next j
        Close 4
    End If




    Err.Clear
    On Error Resume Next
    myrs.Close
    myrsOUT.Close
    Exit Sub
Error1:
    If Dir(App.path & "\" & "Graph_Data.txt") <> "" Then Kill App.path & "\" & "Graph_Data.txt"
    ExcelDataExist = False
    MsgBox Err.Description & " in FindExcelData."
Exit Sub
NoData:
    ExcelDataExist = False
End Sub

Public Sub MakePlotData1()
Dim myrs As Recordset
Dim TheLastDay As Integer
Dim TheLastDate As Date
Dim TheFirstDay As Integer
Dim TheFirstDate As Date
Dim myrsSelect As Recordset
Dim i As Integer
Dim DifferenceInLastDayAndFirstDay As Integer
Dim NextDay As Date
    '============
    On Error GoTo Error1
        If ShowRealDates = 1 Then
            Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Order By RealDate")
            With myrs
                If Not (.EOF And .BOF) Then
                    .MoveFirst
                    TheFirstDate = .Fields("RealDate").Value
                    .MoveLast
                    TheLastDate = .Fields("RealDate").Value
                End If
            End With
        
            DifferenceInLastDayAndFirstDay = DateDiff("d", Format(TheFirstDate, "mm/dd/yyyy"), Format(TheLastDate, "mm/dd/yyyy"))
            For i = 1 To DifferenceInLastDayAndFirstDay
                frmWait.prgLoad = Round((100 / DifferenceInLastDayAndFirstDay) * i)
                NextDay = DateAdd("d", i, Format(TheFirstDate, "mm/dd/yyyy"))
                Set myrsSelect = dbXbuild.OpenRecordset("Select * From Graph_Data where RealDate = #" _
                    & NextDay & "#")
                If myrsSelect.EOF And myrsSelect.BOF Then
                    With myrs
                        .AddNew
                        .Fields("RealDate") = NextDay
                        .Update
                    End With
                End If
            Next i
        
        ElseIf ShowRealDates = 0 Or ShowRealDates = 2 Then
            Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Order By Date")
            With myrs
                If Not (.EOF And .BOF) Then
                    .MoveFirst
                    Do While Not .EOF
                        If .Fields("Date").Value = 0 Then
                            .MoveNext
                        Else
                            Exit Do
                        End If
                    Loop
                    TheFirstDay = .Fields("Date").Value
                    .MoveLast
                    TheLastDay = .Fields("Date").Value
                End If
            End With
            For i = TheFirstDay To TheLastDay
                Set myrsSelect = dbXbuild.OpenRecordset("Select * From Graph_Data where Date = " & i)
                If myrsSelect.EOF And myrsSelect.BOF Then
                    With myrs
                        .AddNew
                        .Fields("Date") = i
                        .Update
                    End With
                End If
            Next i
        End If
       'With myrs
    Exit Sub
Error1:     MsgBox Err.Description & " in MakePlotData1."
End Sub



Public Sub FillNullDataInOut1()
Dim myrs As Recordset
Dim myNumberOfVariables As Integer
Dim i As Integer
Dim j As Integer
Dim a()
Dim d()
Dim NumberOfMyrs_Records As Integer
Dim FirstEvementExists As Boolean
Dim ai_Next
Dim i_First As Integer
Dim i_Second As Integer
Dim k As Integer
Dim delay
Dim g As Integer
Dim prmNumberExpVariables As Integer
Dim NumberOUTVariables
Dim bb As Integer
Dim cc As Integer
    Set myrs = dbXbuild.OpenRecordset("Select * from [Exp_OUT] " & _
        "Where [ExpCheck] = 'Yes'")
    With myrs
        If Not (.EOF And .BOF) Then
            .MoveLast
            prmNumberExpVariables = .RecordCount
        Else
            prmNumberExpVariables = 0
        End If
        .Close
    End With
    If ShowRealDates = 1 Then
        Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Order By RealDate")
        With myrs
            If Not (.EOF And .BOF) Then
                .MoveLast
                NumberOfMyrs_Records = .RecordCount
            Else
                Exit Sub
            End If
        End With
        myNumberOfVariables = myrs.Fields.Count - 1
        
        NumberOUTVariables = myNumberOfVariables - prmNumberExpVariables
        For j = 3 To NumberOUTVariables
            ReDim a(NumberOfMyrs_Records + 10)
            ReDim d(NumberOfMyrs_Records + 10)
            With myrs
                .MoveFirst
                i = 1
                Do While Not (.EOF)
                    If FirstPlantDate(j - 2) = .Fields(2).Value Then Exit Do
                    .MoveNext
                Loop

                Do While Not (.EOF)
                    a(i) = .Fields(j).Value
                    d(i) = .Fields(2).Value
                    i = i + 1
                    If LastPlantDate(j - 2) = .Fields(2).Value Then Exit Do
                    .MoveNext
                Loop
            End With
            Dim mm As Integer
           
            
            For i = mm To NumberOfMyrs_Records - 2
                For cc = 1 To NumberOfMyrs_Records - 2
                    i_First = i
                    If IsNull(a(i + cc)) = False Then
                        i_Second = i + cc
                        Exit For
                    End If
                Next cc
                If i_Second - i_First > 1 Then
                    delay = (a(i_Second) - a(i_First)) / (i_Second - i_First)
                    bb = 1
                    For g = (i_First + 1) To (i_Second - 1)
                        a(g) = a(i_First) + bb * delay
                        bb = bb + 1
                    Next g
                End If
                i = i + cc - 1
            Next i
            
            myrs.MoveFirst
            Do While Not (myrs.EOF)
                If FirstPlantDate(j - 2) = myrs.Fields(2).Value Then Exit Do
                myrs.MoveNext
            Loop

            '''''
            Do While Not (myrs.EOF)
            
            
                If IsNull(myrs.Fields(j).Value) = False Then Exit Do
                myrs.MoveNext
            Loop
            '''''

            
            With myrs
                '.MoveFirst
                i = 1
                Do While Not .EOF
                    .Edit
                    If IsNull(.Fields(j).Value) = True Then
                        .Fields(j).Value = a(i)
                    End If
                    .Update
                    i = i + 1
                    .MoveNext
                Loop
            End With
        Next j
        ''''''''''''''''''''''''''''
    ElseIf ShowRealDates = 0 Or ShowRealDates = 2 Then
        Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Where Date = 0")
        Dim TextToChange As String
        
        With myrs
            If Not (.EOF And .BOF) Then
                .MoveLast
                TextToChange = IIf(IsNull(!TheOrder) = True, "", !TheOrder)
                .Edit
                !TheOrder = "TTT"
                .Update
            End If
            .Close
        End With
        dbXbuild.Execute ("Delete * From Graph_Data Where Date = 0 and TheOrder <> 'TTT'")
        Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Where TheOrder = 'TTT'")
        With myrs
            If Not (.EOF And .BOF) Then
                .MoveLast
                .Edit
                !TheOrder = TextToChange
                .Update
            End If
            .Close
        End With
        
        Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Order By Date")
        With myrs
            If Not (.EOF And .BOF) Then
                .MoveLast
                NumberOfMyrs_Records = .RecordCount
            Else
                Exit Sub
            End If
        End With
        myNumberOfVariables = myrs.Fields.Count - 1
        
        NumberOUTVariables = myNumberOfVariables - prmNumberExpVariables
        For j = 3 To NumberOUTVariables
            ReDim a(NumberOfMyrs_Records + 10)
            ReDim d(NumberOfMyrs_Records + 10)
            '''''''''''''''
            With myrs
                .MoveFirst
                i = 1
                Do While Not (.EOF)
                    If FirstPlantDay(j - 2) = .Fields(1).Value Then Exit Do
                    .MoveNext
                Loop
                Dim bv
                bv = .Fields(j).Name
                Do While Not (.EOF)
                    a(i) = .Fields(j).Value
                    d(i) = .Fields(1).Value
                    i = i + 1
                    If LastPlantDay(j - 2) = .Fields(1).Value Then Exit Do
                    .MoveNext
                Loop
            End With
            '''''''''''''''''''
            For i = mm To NumberOfMyrs_Records - 2
                For cc = 1 To NumberOfMyrs_Records - 2
                    i_First = i
                    If IsNull(a(i + cc)) = False Then
                        i_Second = i + cc
                        Exit For
                    End If
                Next cc
                If i_Second - i_First > 1 Then
                    delay = (a(i_Second) - a(i_First)) / (i_Second - i_First)
                    bb = 1
                    For g = (i_First + 1) To (i_Second - 1)
                        a(g) = a(i_First) + bb * delay
                        bb = bb + 1
                    Next g
                End If
                i = i + cc - 1
                If i > NumberOfMyrs_Records - 2 Then
                    Exit For
                End If
            Next i
            
            myrs.MoveFirst
            Do While Not (myrs.EOF)
                If FirstPlantDay(j - 2) = myrs.Fields(1).Value Then Exit Do
                myrs.MoveNext
            Loop
            '''''
            Do While Not (myrs.EOF)
            
            
                If IsNull(myrs.Fields(j).Value) = False Then Exit Do
                myrs.MoveNext
            Loop
            '''''
            
            With myrs
                i = 1
                Do While Not .EOF
                    .Edit
                   ' Debug.Print i & "-" & .Fields(j).Value & "-" & a(i)
                    If IsNull(.Fields(j).Value) = True And IsNull(a(i)) = False Then
                        .Fields(j).Value = a(i)
                    
                    End If
                    
                    .Update
                    i = i + 1
                    .MoveNext
                Loop
            End With
        Next j
    End If
    Exit Sub
Error1:    MsgBox Err.Description & " in FillNullDataInOut."
    
End Sub



Public Function True_ExpExist(my_prmTableOutName As String, my_VariableID As String) As Boolean
Dim myrs As Recordset
Dim myrsExp As Recordset
Dim OutFileShort As String
Dim i As Integer
Dim k As Integer

    True_ExpExist = False
    
    OutFileShort = Replace(my_prmTableOutName, "_List", "")
    Set myrs = dbXbuild.OpenRecordset("Select Distinct ExperimentID, CropID From [" & OutFileShort & "_File_Info]")
    With myrs
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do While Not .EOF
                Err.Clear
                On Error GoTo Error1
                For k = 1 To NumberT_Tables
                    
                    If ArrayT_Tables(k) = !ExperimentID & "." & !CropID & "T" Then
                
                        Set myrsExp = dbXbuild.OpenRecordset(!ExperimentID & " " & !CropID & "T")
                        If Not (myrsExp.EOF And myrsExp.BOF) Then
                            For i = 0 To myrsExp.Fields.Count - 1
                                If my_VariableID = myrsExp.Fields(i).Name Then
                                    True_ExpExist = True
                                    ExpExists = True
                                    Exit Function
                                End If
                            Next i
                        Else
                            True_ExpExist = False
                        End If
                    End If
                Next k
                .MoveNext
            Loop
        Else
            True_ExpExist = False
            Exit Function
        End If
    End With
    
    Exit Function
Error1:
    True_ExpExist = False
   ' ExpExists = False
   ' MsgBox Err.Description & Err.Number & "    in True_ExpExist/frmSelection."
End Function

Public Sub Add_DAP_Collumn_To_OUT()
Dim myrs2 As Recordset
Dim myrs3 As Recordset
Dim myrs As Recordset
Dim TableName() As String
Dim TableName_YesDAP() As String
Dim TableName_NoDAP() As String

Dim Number_YesDAP_Tables As Integer
Dim Number_NoDAP_Tables As Integer
Dim i As Integer
Dim j As Integer
Dim TableWithDAP As String
    
    On Error GoTo Error1
    
    Number_YesDAP_Tables = 0
    Number_NoDAP_Tables = 0
    ReDim TableName(NumberOutTables)
    For i = 1 To NumberOutTables
        TableName(i) = Replace(OUTTableNames(i), ".", "_")
    Next i
    For i = 1 To NumberOutTables
        Set myrs = dbXbuild.OpenRecordset("Select * From " & _
            TableName(i) & " Where DAP Is Not Null")
        If Not (myrs.EOF And myrs.BOF) Then
            Number_YesDAP_Tables = Number_YesDAP_Tables + 1
            ReDim Preserve TableName_YesDAP(Number_YesDAP_Tables)
            TableName_YesDAP(Number_YesDAP_Tables) = TableName(i)
         Else
            Number_NoDAP_Tables = Number_NoDAP_Tables + 1
            ReDim Preserve TableName_NoDAP(Number_NoDAP_Tables)
            TableName_NoDAP(Number_NoDAP_Tables) = TableName(i)
        End If
    Next i
        
    If Number_YesDAP_Tables > 0 Then
        TableWithDAP = TableName_YesDAP(1)
        DAPexists = True
    Else
        DAPexists = False
        Exit Sub
    End If
        
    If Number_NoDAP_Tables > 0 Then
        For i = 1 To Number_NoDAP_Tables
           ' If CDAYexists = True Then
            '    Set myrs2 = dbXbuild.OpenRecordset("Select * From " & TableName_NoDAP(i) & " Order by CDAY")
           ' Else
                Set myrs2 = dbXbuild.OpenRecordset("Select * From " & TableName_NoDAP(i) & " Order by DAS")
          '  End If
          '  If Not (myrs2.EOF And myrs2.BOF) Then
             '   myrs2.MoveFirst
              '  Do While Not myrs2.EOF
                  '  myrs2.Edit
                    'If CDAYexists = True Then
                      '  Set myrs3 = dbXbuild.OpenRecordset("Select * From [" & _
                            TableWithDAP & "] Where [DATE] = '" & _
                            myrs2.Fields("DATE").Value & _
                            "' And [RunNumber] = " & _
                            myrs2.Fields("RunNumber").Value & _
                            " Order By CDAY")
                   ' Else
                      '  Set myrs3 = dbXbuild.OpenRecordset("Select * From [" & _
                            TableWithDAP & "] Where [DATE] = '" & _
                            myrs2.Fields("DATE").Value & _
                            "' And [RunNumber] = " & _
                            myrs2.Fields("RunNumber").Value & _
                            " Order By DAS")
                    
                   ' End If
                    'If CDAYexists = True Then
                      '  Set myrs3 = dbXbuild.OpenRecordset("Select * From [" & _
                       '     TableWithDAP & "] Where [RunNumber] = " & _
                        '    myrs2.Fields("RunNumber").Value & _
                          '  " And CDAY <> 0" & _
                           ' " Order By CDAY")
                   ' Else
                        Set myrs3 = dbXbuild.OpenRecordset("Select * From [" & _
                            TableWithDAP & "] Where [RunNumber] = " & _
                            myrs2.Fields("RunNumber").Value & _
                            " And DAS <> 0 and DAP <> 0" & _
                            " Order By DAS")
                    
                   ' End If
                    If Not (myrs3.EOF And myrs3.BOF) Then
                        Dim DifferenceInDays As Long
                        With myrs3
                            .MoveFirst
                            DifferenceInDays = !DAS - !DAP
                        End With
                    End If
                  '  If Not (myrs3.EOF And myrs3.BOF) Then
                  '      myrs2.Fields("DAP").Value = myrs3.Fields("DAP").Value
                  '  Else
                  '      myrs2.Fields("DAP").Value = 0
                   ' End If
                   ' myrs2.Update
                   ' myrs2.MoveNext
               ' Loop
               ' myrs3.Close
            'End If
            
            With myrs2
                If Not (.EOF And .BOF) Then
                    .MoveFirst
                    Do While Not .EOF
                        .Edit
                        !DAP = IIf((!DAS - DifferenceInDays) < 0, 0, !DAS - DifferenceInDays)
                        .Update
                        .MoveNext
                    Loop
                End If
                .Close
            End With
        
        Next i
        myrs2.Close
    End If
    
    Exit Sub
Error1:
   ' MsgBox Err.Description & " in Add_DAP_Collumn_To_OUT/frmSelection."

End Sub

Public Function AnyDataExists() As Boolean
Dim myrs As Recordset

Dim NumberDataFields As Integer
Dim NumberDataRecords As Double
Dim i As Integer
Dim DataFieldName() As String
Dim NumberNullValues As Double
Dim TotalNullFields As Integer
    On Error GoTo Error1
    
    TotalNullFields = 5
    Set myrs = dbXbuild.OpenRecordset("Graph_Data")
    With myrs
        If Not (.EOF And .BOF) Then
            .MoveLast
            NumberDataRecords = .RecordCount
            NumberDataFields = .Fields.Count - 3
            ReDim DataFieldName(NumberDataFields)
            For i = 1 To NumberDataFields
                DataFieldName(i) = .Fields(i + 2).Name
            Next i
            .Close
        Else
            AnyDataExists = False
            .Close
            Exit Function
        End If
    End With
    
    If NumberDataFields > 0 Then
        For i = 1 To NumberDataFields
            Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] " & _
                "Where isnull([" & DataFieldName(i) & "]) = True")
            
            With myrs
                If Not (.EOF And .BOF) Then
                    .MoveLast
                    NumberNullValues = .RecordCount
                    If NumberNullValues = NumberDataRecords Then
                        NumberNullValues = 0
                        .Edit
                        '.Fields(DataFieldName(i)).Value = 0
                        .Update
                    Else
                        NumberNullValues = 1
                    End If
                Else
                    NumberNullValues = 1
                End If
            End With
            TotalNullFields = TotalNullFields + NumberNullValues
        Next i
        If TotalNullFields > 5 Then
            AnyDataExists = True
        Else
            AnyDataExists = False
        End If
        myrs.Close
    Else
        AnyDataExists = False
        myrs.Close
        Exit Function
    End If
    Exit Function
Error1:
    MsgBox Err.Description & " in AnyDataExists/frmSelection."
End Function
