VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmStatistic 
   Caption         =   "Statistic"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10845
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   10845
   Tag             =   "1078"
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Orientation     =   2
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
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
      Left            =   5520
      TabIndex        =   14
      Tag             =   "1010"
      Top             =   2400
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Left            =   6480
      TabIndex        =   13
      Tag             =   "1058"
      Top             =   2400
      Width           =   885
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmStatistic.frx":0000
      Height          =   2175
      Left            =   0
      OleObjectBlob   =   "frmStatistic.frx":0014
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label Label2 
      Caption         =   "Note: Mean values were rounded to the nearest digit."
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
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Tag             =   "1035"
      Top             =   2280
      Width           =   4575
   End
   Begin VB.Label lblToolPrint 
      Caption         =   "Label2"
      Height          =   135
      Left            =   8160
      TabIndex        =   17
      Tag             =   "1077"
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lbltoolClose 
      Caption         =   "Label1"
      Height          =   135
      Left            =   0
      TabIndex        =   16
      Tag             =   "1014"
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   12
      Left            =   4920
      TabIndex        =   15
      Tag             =   "1133"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   11
      Left            =   2760
      TabIndex        =   12
      Tag             =   "1132"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   10
      Left            =   840
      TabIndex        =   11
      Tag             =   "1131"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   9
      Left            =   840
      TabIndex        =   10
      Tag             =   "1130"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   8
      Left            =   840
      TabIndex        =   9
      Tag             =   "1129"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   7
      Left            =   840
      TabIndex        =   8
      Tag             =   "1128"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   6
      Left            =   840
      TabIndex        =   7
      Tag             =   "1127"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   6
      Tag             =   "1126"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   5
      Tag             =   "1125"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   4
      Tag             =   "1124"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   3
      Tag             =   "1123"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Tag             =   "1122"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Tag             =   "1121"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmStatistic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo Error1
   
    With CommonDialog1
        .Flags = &H4
        .CancelError = True
        On Error GoTo ErrHandler
        .PrinterDefault = True
        CommonDialog1.FontName = "Courier New"
        CommonDialog1.FontSize = 6
        CommonDialog1.FontBold = False
        .ShowPrinter
    End With
    
    Printer.Orientation = 2
    Printer.Print " "
    
    With Data1.Recordset
        Err.Clear
        On Error Resume Next
        Printer.FontName = "Courier New"
        Err.Clear
        On Error GoTo Error1
        Printer.FontSize = 14
        Printer.FontBold = True
        Printer.Print "                              Statistic"
        Printer.Print " "
        Printer.Print " "
        Printer.FontSize = 10
        Printer.FontBold = True
        
        Printer.Print Spc(26); Spc(16); Spc(16); "Mean"; Spc(16); _
            Spc(11); "Std.Dev."; Spc(18)
        
        Printer.Print Spc(10); "Variable Name"; Spc(13); Spc(6); "Observed"; Spc(6); "Simulated"; Spc(5); _
         "Ratio"; Spc(6); "Observed"; Spc(6); "Simulated"; Spc(5); _
         "r-Square"; Spc(7)
        Printer.FontBold = False
        
        .MoveFirst
        Err.Clear
        On Error Resume Next
        Do While Not .EOF
            Printer.Print Spc(10); .Fields(0).Value; Spc(31 - Len(.Fields(0).Value)); _
            Str(.Fields(1).Value); Spc(14 - Len(Str(.Fields(1).Value))); _
            Str(.Fields(2).Value); Spc(14 - Len(Str(.Fields(2).Value))); _
            Str(.Fields(3).Value); Spc(14 - Len(Str(.Fields(3).Value))); _
            Str(.Fields(4).Value); Spc(14 - Len(Str(.Fields(4).Value))); _
            Str(.Fields(5).Value); Spc(14 - Len(Str(.Fields(5).Value))); _
            Str(.Fields(6).Value); Spc(14 - Len(Str(.Fields(6).Value)))
            .MoveNext
        Loop
        Printer.Print " "
        Printer.Print " "
        Printer.FontBold = True
        Printer.Print Spc(10); "Variable Name"; Spc(13); Spc(6); "Mean Diff."; Spc(4); "Mean Abs.Diff"; Spc(1); _
         "RMSE"; Spc(10); "d-Stat."; Spc(7); "Used Obs."; Spc(5); "Total Obs."
         
         Printer.FontBold = False
        .MoveFirst
        Do While Not .EOF
            Printer.Print Spc(10); .Fields(0).Value; Spc(31 - Len(.Fields(0).Value)); _
            Str(.Fields(7).Value); Spc(14 - Len(Str(.Fields(7).Value))); _
            Str(.Fields(8).Value); Spc(14 - Len(Str(.Fields(8).Value))); _
            Str(.Fields(9).Value); Spc(14 - Len(Str(.Fields(9).Value))); _
            Str(.Fields(10).Value); Spc(14 - Len(Str(.Fields(10).Value))); _
            Str(.Fields(11).Value); Spc(14 - Len(Str(.Fields(11).Value))); _
            Str(.Fields(12).Value); Spc(14 - Len(Str(.Fields(12).Value)))
            .MoveNext
        Loop
    
    End With
    Printer.EndDoc
                                          
                                          
    Exit Sub
Error1:     MsgBox Err.Description
    Exit Sub
ErrHandler:
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
Dim myAppPath As String
    On Error GoTo Error1
    Me.Top = frmGraph.MSChart1.Top
    Me.Width = 9450
    Me.Height = 3150
    LoadResStrings Me
  '  If gsglXFactor <> 1 Or gsglYFactor <> 1 Then
        Call SetDeviceIndependentWindow(Me, 1)
  '  End If

    If Right$(App.path, 1) = "\" Then
        myAppPath = App.path
    Else
        myAppPath = App.path & "\"
    End If
    With Data1
        .DatabaseName = myAppPath & "GraphCreateTEMP.mdb"
        .RecordSource = "Statistic_Calculated"
        .Refresh
    End With
    Me.DBGrid1.Refresh
    
    ColumnNames
    
    With cmdPrint
        .Top = Me.Height - (2 + 1 / 4) * .Height
        .Left = Me.Width - (1 + 1 / 2) * .Width - cmdOK.Width
        .ToolTipText = lblToolPrint.Caption
    End With
    With cmdOK
        .Top = Me.Height - (2 + 1 / 4) * .Height
        .Left = Me.Width - (1 + 1 / 4) * .Width
        .ToolTipText = lbltoolClose.Caption
    End With
    Unload frmDocument
    
    Exit Sub
Error1: MsgBox Err.Description & " in Form_Load/frmStatistic."
End Sub


Public Sub ColumnNames()
Dim col() As Column
Dim NumberOfColumns As Integer
Dim i As Integer
    On Error GoTo Error1
    NumberOfColumns = 13
    ReDim Preserve col(NumberOfColumns)
    For i = 0 To NumberOfColumns - 1
        Set col(i) = DBGrid1.Columns(i)
        col(i).Caption = Trim(Label1(i).Caption)
        col(i).WrapText = True
        col(i).Width = 600 * gsglXFactor
        col(i).Alignment = 0
    Next i
    Set col(0) = DBGrid1.Columns(0)
    col(0).Width = 1800 * gsglXFactor
    DBGrid1.Width = 9005 * gsglXFactor
    Exit Sub
Error1: MsgBox Err.Description & " in ColumnNames/frmStatistic."
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 15
        cmdOK_Click
    Case 16
        cmdPrint_Click
    End Select
End Sub

