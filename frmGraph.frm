VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmGraph 
   AutoRedraw      =   -1  'True
   Caption         =   "Graph"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9330
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9330
   Tag             =   "1057"
   Begin VB.CommandButton cmdExportToFile 
      Caption         =   "Export to file"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   9
      Tag             =   "1036"
      Top             =   3960
      Width           =   885
   End
   Begin VB.CommandButton cmdStatistic 
      Caption         =   "Statistic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   4
      Tag             =   "1116"
      Top             =   3000
      Width           =   885
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
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Tag             =   "1010"
      Top             =   3480
      Width           =   885
   End
   Begin VB.CommandButton cmdBack 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Tag             =   "1050"
      Top             =   4920
      Width           =   885
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export to Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   1
      Tag             =   "1064"
      Top             =   4440
      Width           =   885
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   6015
      Left            =   120
      OleObjectBlob   =   "frmGraph.frx":0000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   7695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8280
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Orientation     =   2
   End
   Begin VB.Label lblToolExportFile 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   10
      Tag             =   "1037"
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbltoolBack 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Tag             =   "1018"
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblToolExcel 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      Tag             =   "1017"
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblToolPrint 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   6
      Tag             =   "1016"
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblToolStatistic 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   5
      Tag             =   "1015"
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   6255
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7935
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NumberOfRecords As Integer
Dim NumberOfVariables As Integer
Dim TotalNumberDatPlot As Integer
Dim NewVariableName() As String
Dim ColorExpMatch() As Integer
Dim ColorOUTMatch() As Integer
Dim NumberMatchingPairs As Integer
Dim ArrowVariableY() As String
Dim ArrowVariableX() As String
Dim ArrowVariableX_Exp() As String
Dim ArrowVariableY_Exp() As String
Dim NumberYVariables As Integer
Dim NumberXVariables As Integer
Dim PlotTitle As String
Dim XAxisTitle As String
Dim EnoughData As Boolean
Dim maxXvalue
Dim minXvalue
Dim maxYvalue
Dim minYvalue
Dim NewName As String


Private Sub cmdExport_Click()
    CreateExcelApp
End Sub

Private Sub cmdBack_Click()
    If ExpData_vs_Simulated = 1 Then
        KeepSim_vs_Obs = True
    Else
        KeepSim_vs_Obs = False
    End If
    frmSelection.Show
    Unload Me
End Sub

Private Sub cmdExportToFile_Click()
  '  WriteDataToFile
    frmDataExport.Show
End Sub

Private Sub cmdPrint_Click()
Dim MSChart1_Width
Dim MSChart1_Height
Dim MSChart1_Top
Dim MSChart1_Left
Dim Form_Hight
    
   
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
    Printer.Print " "
    Printer.Print " "
    
    Me.BackColor = &HFFFFFF
   ' MSChart1.Left = 2000
   ' MSChart1.Height = 7000
   ' MSChart1.Width = 9000

    MSChart1_Width = MSChart1.Width
    MSChart1_Height = MSChart1.Height
    MSChart1_Top = MSChart1.Top
    MSChart1_Left = MSChart1.Left
    MSChart1.Top = 500
    Form_Hight = Me.Height
    Me.cmdBack.Visible = False
    Me.cmdExport.Visible = False
    Me.cmdPrint.Visible = False
    Me.cmdStatistic.Visible = False
    Me.cmdExportToFile.Visible = False
    'Me.Label1.Visible = False
    
    MSChart1.Left = 2000
    MSChart1.Height = 9000
    MSChart1.Width = 11000
    Shape1.Visible = True
    Shape1.Top = MSChart1.Top - 5
    Shape1.Left = MSChart1.Left - 5
    Shape1.Width = MSChart1.Width + 10
    Shape1.Height = MSChart1.Height + 10

    Me.PrintForm
    Me.Height = Form_Hight
    MSChart1.Width = MSChart1_Width
    MSChart1.Height = MSChart1_Height
    MSChart1.Top = MSChart1_Top
    MSChart1.Left = MSChart1_Left
    Me.BackColor = &H8000000F
    Shape1.Visible = False
    'Me.MSChart1.Width = 11000
    Me.cmdBack.Visible = True
    Me.cmdExport.Visible = True
    Me.cmdPrint.Visible = True
    Me.cmdStatistic.Visible = True
    Me.cmdExportToFile.Visible = True
    'Me.Label1.Visible = True
ErrHandler:
End Sub

Private Sub cmdStatistic_Click()
    If EnoughData = False Then
        Exit Sub
    Else
        frmStatistic.Show 1
    End If
 
    'frmStatistic.Show 1
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
    Case 2
        cmdBack_Click
    Case 19
       cmdStatistic_Click
    Case 5
        cmdExport_Click
    Case 6
        cmdPrint_Click
    Case 16
        cmdExportToFile_Click
    End Select
End Sub

Private Sub Form_Load()
'Scatter plot will be shown on the screen.
'All OUT data will be connected with a line.
'All corresponding exp. data will be shown in matching color marks.

Dim i As Integer
Dim FinalDataPlot()
Dim j As Integer
Dim ColorRed As Integer
Dim ColorGreen As Integer
Dim ColorBlue As Integer
Dim gg
Dim nm As Integer
Dim StringCut As String
Dim myrsDATA As Recordset
Dim myrsMydata As Recordset
Dim FieldName As String
Dim h As Integer
Dim bb As Integer
Dim aa As Integer
'Dim StartValue() As Single
    Me.Top = 0
    Me.Left = 0
    Me.Width = 9450
    Me.Height = 6400
    LoadResStrings Me
    Screen.MousePointer = vbHourglass
    Unload frmDocument
    
   ' If gsglXFactor <> 1 Or gsglYFactor <> 1 Then
        Call SetDeviceIndependentWindow(Me, 1)
   ' End If
    
    If FindExcelInstance = False Then
        ExcelDataExist = False
    End If
  
  '  If FileType = "SUM" Then
    
       ' cmdExportToFile.Enabled = False
   ' End If
    If ShowStatistic = 1 Then
        cmdStatistic.Enabled = True
    Else
        cmdStatistic.Enabled = False
    End If

    If ExcelDataExist = False Then
        cmdExport.Enabled = False
    Else
        If DisableExcelButton = True Then
            cmdExport.Enabled = False
        Else
            cmdExport.Enabled = True
        End If
    End If
    cmdStatistic.ToolTipText = lblToolStatistic.Caption
    cmdExport.ToolTipText = lblToolExcel.Caption
    cmdPrint.ToolTipText = lblToolPrint.Caption
    cmdBack.ToolTipText = lbltoolBack.Caption
    cmdExportToFile.ToolTipText = lblToolExportFile & " " & App.path & "\" & "Graph_Data[].txt"
    
    On Error GoTo Error1

    NumberMatchingPairs = 0
   ' MakeDataArray
  '  If FileType <> "SUM" Then
      '  MakePlotData1
  '  End If
    
    If FileType = "OUT" Then
        frmWait.prgLoad.Value = 0
        FindMatchingPairs
        
        FillNullDataInOut0
        frmWait.prgLoad.Value = 30
        
        EnoughData = True
        If ShowStatistic = 1 Then
            InStatistic = True
            CalculateStatistic
        End If
    End If
        If ExpData_vs_Simulated = 1 Then
             FillNullDataInOut_expVSsim
        End If
    If FileType = "SUM" Then
        FindMatchingPairs
        EnoughData = True
        If ShowStatistic = 1 Then
            CalculateStatistic
        End If
    End If
    MakeDataAdjusted
    
    MakeDataAdjusted_Remove_Multiplier
    frmWait.prgLoad.Value = 50
    If FileType <> "SUM" Then
        If ShowRealDates = 1 Then
            Set myrsDATA = dbXbuild.OpenRecordset("Select * From Adjusted_Graph_Data Order By RealDate")
        ElseIf ShowRealDates = 0 Then
            Set myrsDATA = dbXbuild.OpenRecordset("Select * From Adjusted_Graph_Data Order By Date")
        ElseIf ShowRealDates = 2 Then
            Set myrsDATA = dbXbuild.OpenRecordset("Select * From Adjusted_Graph_Data Order By Date")
        End If
    Else
        Set myrsDATA = dbXbuild.OpenRecordset("Adjusted_Graph_Data")
        ShowX_Axis = 0
    End If
    With myrsDATA
        .MoveLast
        TotalNumberDatPlot = .RecordCount
        NumberOfVariables = .Fields.Count - 3
    End With
    If ShowX_Axis = 1 Then
        PlotTitle = "Time series"
        ReDim Preserve FinalDataPlot(1 To TotalNumberDatPlot, 1 To NumberOfVariables * 2)
        With myrsDATA
            If Not (.EOF And .BOF) Then
                DoEvents
            Else
                Unload frmWait
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            .MoveFirst
            i = 0
            For j = 1 To NumberOfVariables
                .MoveFirst
                i = 0
                Do While Not .EOF
                    If ShowRealDates = 1 Then
                        gg = Format(!realdate, "mm/dd/yyyy")
                        i = i + 1
                        FinalDataPlot(i, 2 * j - 1) = DateDiff("d", "1/1/" & Year(gg), Format(gg, "mm/dd/yyyy")) + 2 + _
                        DateDiff("d", "1/1/1900", "1/1/" & Year(gg))
                    ElseIf ShowRealDates = 0 Or ShowRealDates = 2 Then
                        i = i + 1
                        FinalDataPlot(i, 2 * j - 1) = !Date
                    End If
                    .MoveNext
                Loop
            Next j
            For j = 1 To NumberOfVariables
                frmWait.prgLoad.Value = 50 + Round((50 / NumberOfVariables) * j)
                FieldName = .Fields(j + 2).Name
                i = 0
                .MoveFirst
                Do While Not .EOF
                    i = i + 1
                    FinalDataPlot(i, 2 * j) = Round(.Fields(j + 2).Value, 4)
                    .MoveNext
                Loop
            Next j
        End With
        minXvalue = FinalDataPlot(1, 1)
        maxXvalue = FinalDataPlot(TotalNumberDatPlot, 1)
        Chart_SetUp
        With MSChart1
            .ChartData = FinalDataPlot
            .ColumnCount = NumberOfVariables * 2
            For i = 1 To NumberOfVariables * 2
                .Column = i
                If ShowRealDates = 1 Then
                    .Plot.Axis(VtChAxisIdX).Labels(1).Format = "d mmm yyyy"
                    .Plot.Axis(VtChAxisIdX).AxisTitle = "Date"
                End If
                If ShowRealDates = 0 Then
                    .Plot.Axis(VtChAxisIdX).Labels(1).Format = "###"
                    If CDAYexists = False And DAPexists = False Then
                       .Plot.Axis(VtChAxisIdX).AxisTitle = "Days after Start of Simulation "
                    Else
                        .Plot.Axis(VtChAxisIdX).AxisTitle = "Days after Planting"
                    End If
                End If
                If ShowRealDates = 2 Then
                    .Plot.Axis(VtChAxisIdX).Labels(1).Format = "###"
                    .Plot.Axis(VtChAxisIdX).AxisTitle = "Days of a Year"
                End If
            Next i
            
            If FileType = "OUT" Then
                For i = 1 To NumberOfVariables
                    
                    If NewVariableName(i) <> "" Then
                        
                        With .Plot.SeriesCollection.Item(2 * i - 1)
                            NewName = RemoveSymbols(SelectedVariable(NewVariableName(i)))
                            NewName = RunName(NewName)
                            .LegendText = NewName
                            ColorRed = myColorRed(i)
                            ColorGreen = myColorGreen(i)
                            ColorBlue = myColorBlue(i)
                            'Debug.Print i
                            If prmShowLine = 1 Then
                                .DataPoints(-1).Brush.FillColor.Set ColorRed, ColorGreen, ColorBlue
                                .SeriesMarker.Show = False
                                .ShowLine = True
                            Else
                                .DataPoints(-1).Brush.FillColor.Set 252, 252, 252
                                .SeriesMarker.Show = True
                                .ShowLine = False
                                .SeriesMarker.Auto = False
                                With .DataPoints.Item(-1).Marker
                                    .Style = ArrayMarkerStyle(i)
                                    
                                    .FillColor.Set ColorRed, ColorGreen, ColorBlue
                                    If Marker_Small = 1 Then
                                        .Size = 100
                                    Else
                                        .Size = 100 * Marker_Size
                                    End If
                                    With .Pen.VtColor
                                        .Blue = ColorBlue
                                        .Red = ColorRed
                                        .Green = ColorGreen
                                    End With
                                End With
                            ''
                            End If
                        End With
                    End If
                Next i
                For i = NumberOfVariables - NumberExpVariables + 1 To NumberOfVariables
                    If i > 0 Then
                        If NewVariableName(i) <> "" Then
                            NewName = RemoveSymbols(SelectedVariable(NewVariableName(i)))
                            NewName = RunName(NewName)
                            .Plot.SeriesCollection.Item(2 * i - 1).LegendText = NewName
                        End If
                    End If
                Next i
            Else
                For i = 1 To NumberOfVariables
                    With .Plot.SeriesCollection.Item(2 * i - 1)
                            NewName = RemoveSymbols(SelectedVariable(NewVariableName(i)))
                            NewName = RunName(NewName)
                            .LegendText = NewName
                        ColorRed = myColorRed(i)
                        ColorGreen = myColorGreen(i)
                        ColorBlue = myColorBlue(i)
                        .DataPoints(-1).Brush.FillColor.Set 252, 252, 252
                        .SeriesMarker.Auto = False
                        With .DataPoints.Item(-1).Marker
                            .Style = ArrayMarkerStyle(i)
                            
                            .FillColor.Set ColorRed, ColorGreen, ColorBlue
                            With .Pen.VtColor
                                .Blue = ColorBlue
                                .Red = ColorRed
                                .Green = ColorGreen
                            End With
                        End With
                                    If Marker_Small = 1 Then
                                        .DataPoints.Item(-1).Marker.Size = 100
                                    Else
                                        .DataPoints.Item(-1).Marker.Size = 100 * Marker_Size
                                    End If
                        .SeriesMarker.Show = True
                        .ShowLine = False
                    End With
                Next i
            End If
            
            If NumberMatchingPairs > 0 Then
                For i = 1 To NumberMatchingPairs
                    If prmShowLine = 1 Then
                        With .Plot.SeriesCollection(2 * ColorOUTMatch(i) - 1).DataPoints(-1).Brush.FillColor
                            ColorRed = .Red
                            ColorGreen = .Green
                            ColorBlue = .Blue
                        End With
                    Else
                        With .Plot.SeriesCollection(2 * ColorOUTMatch(i) - 1).DataPoints.Item(-1).Marker.FillColor
                            ColorRed = .Red
                            ColorGreen = .Green
                            ColorBlue = .Blue
                        End With
                    End If
                    
                    With .Plot.SeriesCollection.Item(2 * ColorExpMatch(i) - 1)
                        .SeriesMarker.Auto = False
                        .SeriesMarker.Show = True
                        .ShowLine = False
                        With .DataPoints.Item(-1)
                            With .Marker
                                .Style = ArrayMarkerStyle(i)
                                .FillColor.Set ColorRed, ColorGreen, ColorBlue
                                .Pen.VtColor.Blue = ColorBlue
                                .Pen.VtColor.Red = ColorRed
                                .Pen.VtColor.Green = ColorGreen
                                    If Marker_Small = 1 Then
                                        .Size = 100
                                    Else
                                        .Size = 100 * Marker_Size
                                    End If
                            End With
                            .Brush.FillColor.Set 252, 252, 252
                        End With
                    End With
                    With .Plot.SeriesCollection.Item(2 * ColorOUTMatch(i) - 1)
                        .SeriesMarker.Show = False
                        .ShowLine = True
                            If prmShowLine = 1 Then
                                .SeriesMarker.Show = False
                                .ShowLine = True
                            Else
                                .SeriesMarker.Show = True
                                .ShowLine = False
                            ''
                                .SeriesMarker.Auto = False
                                With .DataPoints.Item(-1).Marker
                                    If Marker_Small = 1 Then
                                        .Size = 100
                                    Else
                                        .Size = 100 * Marker_Size
                                    End If
                                    '.Style = ArrayMarkerStyle(i)
                                    .Style = ArrayMarkerStyle(i + 2)
                                    .FillColor.Set ColorRed, ColorGreen, ColorBlue
                                    With .Pen.VtColor
                                        .Blue = ColorBlue
                                        .Red = ColorRed
                                        .Green = ColorGreen
                                    End With
                                End With
                            ''
                            End If
                    '''
                    End With
                Next i
            End If
        End With
    
    Else 'if X vs Y
        PlotTitle = "Scatter Plot"
        If ExpData_vs_Simulated = 1 Then
            NumberExpVariables_X = 0
            PlotTitle = "Experimental Data vs. Simulated Data"
        
        Else
            PlotTitle = "Scatter Plot"
        End If
        If FileType = "SUM" Then
            'PlotTitle = "Evaluated Data"
            PlotTitle = NameSUM_File
            
        End If
        
        FindPairsXY
        If FileType = "SUM" Then
            
            XAxisTitle = "Simulated Data"
            
        Else
            If NumberXVariables <> 0 Then
                XAxisTitle = RemoveSymbols(SelectedVariable(Trim(Mid(ArrowVariableX(1), 1, InStr(ArrowVariableX(1), "Run") - 1))))
            Else
                XAxisTitle = RemoveSymbols(SelectedVariable(Trim(Mid(ArrowVariableX_Exp(1), 1, InStr(ArrowVariableX_Exp(1), "TRT") - 1))))
            End If
        End If
        'oxanaOxana
        If ExpData_vs_Simulated = 1 Then
            PlotTitle = XAxisTitle
            XAxisTitle = "Simulated"
            
        End If
        
        If FileType <> "SUM" Then
            If ShowRealDates = 1 Then
                Set myrsMydata = dbXbuild.OpenRecordset("Select * From Adjusted_Graph_Data Order by RealDate")
            ElseIf ShowRealDates = 0 Then
                Set myrsMydata = dbXbuild.OpenRecordset("Select * From Adjusted_Graph_Data Order by Date")
            ElseIf ShowRealDates = 2 Then
                Set myrsMydata = dbXbuild.OpenRecordset("Select * From Adjusted_Graph_Data Order by Date")
            End If
        Else
            Set myrsMydata = dbXbuild.OpenRecordset("Adjusted_Graph_Data")
        End If
        
        FindMaxValueScale
        Dim One_to_One_Index As Integer
        
        myrsMydata.MoveLast
        TotalNumberDatPlot = myrsMydata.RecordCount
        NumberOfVariables = myrsMydata.Fields.Count - 3
        myrsMydata.MoveFirst
        ReDim Preserve FinalDataPlot(1 To TotalNumberDatPlot, 1 To (NumberOfVariables - NumberXVariables - NumberExpVariables_X) * 2)
        
        If ExpData_vs_Simulated = 1 Or FileType = "SUM" Then
            ReDim Preserve FinalDataPlot(1 To TotalNumberDatPlot, 1 To (NumberOfVariables - NumberXVariables - NumberExpVariables_X + 1) * 2)
        End If
        With myrsMydata
            If Not (.EOF And .BOF) Then
                DoEvents
            Else
                Unload frmWait
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            .MoveFirst
            i = 1
            One_to_One_Index = NumberXVariables * NumberYVariables + 1
            Dim One_to_One_Value
            Dim One_to_One_Value2
            One_to_One_Value2 = MinAxisValue
            
            Do While Not .EOF
                bb = 1
                aa = 1
                For h = 1 To NumberXVariables
                    
                    For j = 1 To NumberYVariables
                        If ArrowVariableY(h, j) <> "" Then
                            FinalDataPlot(i, 2 * bb - 1) = .Fields(ArrowVariableX(h)).Value
                            FinalDataPlot(i, 2 * bb) = .Fields(ArrowVariableY(h, j)).Value
                            If ExpData_vs_Simulated = 1 Or FileType = "SUM" Then
                                If IsNull(.Fields(ArrowVariableX(h)).Value) = False _
                                    And IsNull(.Fields(ArrowVariableY(h, j)).Value) = False Then
                                    
                                    If .Fields(ArrowVariableX(h)).Value > .Fields(ArrowVariableY(h, j)).Value Then
                                        One_to_One_Value = .Fields(ArrowVariableX(h)).Value
                                    Else
                                        One_to_One_Value = .Fields(ArrowVariableY(h, j)).Value
                                    End If
                                ElseIf IsNull(.Fields(ArrowVariableX(h)).Value) = False Then
                                    One_to_One_Value = .Fields(ArrowVariableX(h)).Value
                                ElseIf IsNull(.Fields(ArrowVariableY(h, j)).Value) = False Then
                                    One_to_One_Value = .Fields(ArrowVariableY(h, j)).Value
                                End If
                                If One_to_One_Value2 < One_to_One_Value Then
                                    One_to_One_Value2 = One_to_One_Value
                                    
                                End If
                              ' Debug.Print One_to_One_Value2
                               ' FinalDataPlot(i, 2 * One_to_One_Index - 1) = One_to_One_Value
                               ' FinalDataPlot(i, 2 * One_to_One_Index) = One_to_One_Value
                            End If

                            bb = bb + 1
                            
                        End If
                    Next j
                Next h
                If ExpData_vs_Simulated = 1 Or FileType = "SUM" Then
                    FinalDataPlot(i, 2 * One_to_One_Index - 1) = One_to_One_Value
                    FinalDataPlot(i, 2 * One_to_One_Index) = One_to_One_Value
                End If
                
                If NumberExpVariables_X > 0 Then
                    For h = 1 To NumberExpVariables_X
                        For j = 1 To NumberExpVariables_Y / NumberExpVariables_X
                            If ArrowVariableY_Exp(h, j) <> "" Then
                                FinalDataPlot(i, 2 * (bb + aa - 1) - 1) = .Fields(ArrowVariableX_Exp(h)).Value
                                FinalDataPlot(i, 2 * (bb + aa - 1)) = .Fields(ArrowVariableY_Exp(h, j)).Value
                                aa = aa + 1
                            End If
                        Next j
                    Next h
                End If
                i = i + 1
                .MoveNext
            Loop
           ' MinAxisValue = 0
            If ExpData_vs_Simulated = 1 Or FileType = "SUM" Then
                FinalDataPlot(1, 2 * One_to_One_Index - 1) = MinAxisValue
                FinalDataPlot(1, 2 * One_to_One_Index) = MinAxisValue
            End If
            If ExpData_vs_Simulated = 1 Or FileType = "SUM" Then
                FinalDataPlot(i - 1, 2 * One_to_One_Index - 1) = MaxAxisValue
                FinalDataPlot(i - 1, 2 * One_to_One_Index) = MaxAxisValue
            End If

            
            
            .Close
        End With
        Chart_SetUp
        Dim mySeriesCollectionNumber As Integer
        
        With MSChart1
            .ChartData = FinalDataPlot
            .ColumnCount = .Plot.SeriesCollection.Count
            If ExpData_vs_Simulated = 1 Or FileType = "SUM" Then
            .Plot.Axis(VtChAxisIdX).Intersection.Auto = False
            .Plot.Axis(VtChAxisIdY).Intersection.Auto = False
            
           ' .Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
           ' .Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
         '  If .Plot.Axis(VtChAxisIdX).ValueScale.Maximum > .Plot.Axis(VtChAxisIdY).ValueScale.Maximum Then
         '       MaxAxisValue = .Plot.Axis(VtChAxisIdX).ValueScale.Maximum
          ' Else
          '      MaxAxisValue = .Plot.Axis(VtChAxisIdY).ValueScale.Maximum
          ' End If
            .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinAxisValue
           ' .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxAxisValue
           ' .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxAxisValue
            .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinAxisValue
           ' .Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = 7
           '.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 7
                .Plot.UniformAxis = True
            Else
          
          '  .Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
           ' .Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
           ' .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxAxisValue
           ' .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinAxisValue
           ' .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxAxisValue
           ' .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinAxisValue
            '.Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = 7
           '.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 7
           ' .Plot.UniformAxis = True
            End If
            
              mySeriesCollectionNumber = .Plot.SeriesCollection.Count
              For i = 1 To .Plot.SeriesCollection.Count - (NumberExpVariables_Y) * 2
                If ShowX_Axis <> 1 Then prmShowLine = 0
                With .Plot.SeriesCollection.Item(i)
                    If prmShowLine = 1 Then
                        .SeriesMarker.Show = False
                        .ShowLine = True
                        .Pen.Style = LineStyle(i)
                        .Pen.Width = 40
                        .DataPoints.Item(-1).Brush.FillColor.Set myColorRed(i), myColorGreen(i), myColorBlue(i)
                    Else
                        .SeriesMarker.Show = True
                        .ShowLine = False
                        .SeriesMarker.Auto = False
                        With .DataPoints.Item(-1).Marker
                                    If Marker_Small = 1 Then
                                        .Size = 100
                                    Else
                                        .Size = 100 * Marker_Size
                                    End If
                            .Style = ArrayMarkerStyle(i)
                            .FillColor.Set myColorRed(i), myColorGreen(i), myColorBlue(i)
                            With .Pen.VtColor
                                .Blue = myColorBlue(i)
                                .Red = myColorRed(i)
                                .Green = myColorGreen(i)
                            End With
                        End With
                        .ShowLine = False
                    End If
                End With
            Next i
            
            If ExpData_vs_Simulated = 1 Then
                With .Plot.SeriesCollection.Item(2 * One_to_One_Index - 1)
                    .SeriesMarker.Show = False
                    .ShowLine = True
                    .Pen.Style = VtPenStyleSolid
                    .Pen.Width = 1
                    .LegendText = "1x1"
                    .Pen.VtColor.Blue = myColorBlue(4)
                    .Pen.VtColor.Red = myColorRed(4)
                    .Pen.VtColor.Green = myColorGreen(4)
                End With
                With .Plot.SeriesCollection.Item(2 * One_to_One_Index)
                    .SeriesMarker.Show = False
                    .ShowLine = True
                    .Pen.Style = VtPenStyleSolid
                    .Pen.Width = 1
                    .Pen.VtColor.Blue = myColorBlue(4)
                    .Pen.VtColor.Red = myColorRed(4)
                    .Pen.VtColor.Green = myColorGreen(4)
                End With
            ElseIf FileType = "SUM" Then
                With .Plot.SeriesCollection.Item(2 * One_to_One_Index - 1)
                    .SeriesMarker.Show = False
                    .ShowLine = False
                    .Pen.Style = VtPenStyleNull
                    .Pen.Width = 0
                    .LegendText = ""
                    .Pen.VtColor.Blue = myColorBlue(4)
                    .Pen.VtColor.Red = myColorRed(4)
                    .Pen.VtColor.Green = myColorGreen(4)
                End With
                With .Plot.SeriesCollection.Item(2 * One_to_One_Index)
                    .SeriesMarker.Show = False
                    .ShowLine = True
                    .Pen.Style = VtPenStyleNull
                    .Pen.Width = 0
                    .Pen.VtColor.Blue = myColorBlue(4)
                    .Pen.VtColor.Red = myColorRed(4)
                    .Pen.VtColor.Green = myColorGreen(4)
                End With
            End If
           
            
            i = 0
            For i = .Plot.SeriesCollection.Count - (NumberExpVariables_Y) * 2 To .Plot.SeriesCollection.Count
                If i > 0 Then
                    With .Plot.SeriesCollection.Item(i)
                        .SeriesMarker.Show = True
                        .ShowLine = False
                        .SeriesMarker.Auto = False
                        With .DataPoints.Item(-1).Marker
                                    If Marker_Small = 1 Then
                                        .Size = 100
                                    Else
                                        .Size = 100 * Marker_Size
                                    End If
                            .Style = ArrayMarkerStyle(i)
                            .FillColor.Set myColorRed(i), myColorGreen(i), myColorBlue(i)
                            With .Pen.VtColor
                                .Blue = myColorBlue(i)
                                .Red = myColorRed(i)
                                .Green = myColorGreen(i)
                            End With
                        End With
                    End With
                Else
                    Exit For
                End If
            Next i
            i = 1
                
                For h = 1 To NumberXVariables
                    ColorRed = myColorRed(h)
                    ColorGreen = myColorGreen(h)
                    ColorBlue = myColorBlue(h)
                    For j = 1 To NumberYVariables
                        With .Plot.SeriesCollection.Item(2 * i - 1)
                            NewName = RemoveSymbols(SelectedVariable(ArrowVariableY(h, j)))
                            'OxanaOxana
                            If FileType <> "SUM" Then
                                NewName = RunName(NewName)
                            End If
                            
                            If ExpData_vs_Simulated = 1 Then
                                NewName = Mid(NewName, InStr(NewName, ")") + 1)
                            End If
                        
                            
                            .LegendText = NewName
                            
                            .SeriesMarker.Auto = False
                            With .DataPoints.Item(-1).Marker
                                    If Marker_Small = 1 Then
                                        .Size = 100
                                    Else
                                        .Size = 100 * Marker_Size
                                    End If
                                .FillColor.Set ColorRed, ColorGreen, ColorBlue
                                With .Pen.VtColor
                                    .Blue = ColorBlue
                                    .Red = ColorRed
                                    .Green = ColorGreen
                                End With
                            End With
                            i = i + 1
                            .DataPoints.Item(-1).Brush.FillColor.Set ColorRed, ColorGreen, ColorBlue
                            .DataPoints.Item(-1).Brush.FillColor.Set 252, 252, 252
                        End With
                    Next j
                
                Next h
                
                If NumberXVariables > 0 Then
                    For h = 1 To NumberXVariables
                        ColorRed = myColorRed(h)
                        ColorGreen = myColorGreen(h)
                        ColorBlue = myColorBlue(h)
                        If NumberExpVariables_X <> 0 Then
                            For j = 1 To NumberExpVariables_Y / NumberExpVariables_X
                                If 2 * i - 1 < mySeriesCollectionNumber + 1 Then
                                With .Plot.SeriesCollection.Item(2 * i - 1)
                                    NewName = RemoveSymbols(SelectedVariable(ArrowVariableY_Exp(h, j)))
                                    If FileType <> "SUM" Then
                                        NewName = RunName(NewName)
                                    End If
                                    .LegendText = NewName
                                    .SeriesMarker.Auto = False
                                    With .DataPoints.Item(-1).Marker
                                        .FillColor.Set ColorRed, ColorGreen, ColorBlue
                                    If Marker_Small = 1 Then
                                        .Size = 100
                                    Else
                                        .Size = 100 * Marker_Size
                                    End If
                                        '.Style = ArrayMarkerStyle(h + i + 5)
                                        .Style = ArrayMarkerStyle(i + 1)
                                        With .Pen.VtColor
                                            .Blue = ColorBlue
                                            .Red = ColorRed
                                            .Green = ColorGreen
                                        End With
                                    End With
                                    i = i + 1
                                    .DataPoints.Item(-1).Brush.FillColor.Set 252, 252, 252
                                End With
                                End If
                            Next j
                        End If
                    Next h
                Else
                    If NumberExpVariables_X <> 0 Then
                        For h = 1 To NumberExpVariables_X
                            ColorRed = myColorRed(h)
                            ColorGreen = myColorGreen(h)
                            ColorBlue = myColorBlue(h)
                            For j = 1 To NumberExpVariables_Y / NumberExpVariables_X
                                With .Plot.SeriesCollection.Item(2 * i - 1)
                                    .SeriesMarker.Show = True
                                    .SeriesMarker.Auto = False
                                    .ShowLine = False
                                    NewName = RemoveSymbols(SelectedVariable(ArrowVariableY_Exp(h, j)))
                                    NewName = RunName(NewName)
                                    .LegendText = NewName
                                    .SeriesMarker.Auto = False
                                    With .DataPoints.Item(-1).Marker
                                        .FillColor.Set ColorRed, ColorGreen, ColorBlue
                                    If Marker_Small = 1 Then
                                        .Size = 100
                                    Else
                                        .Size = 100 * Marker_Size
                                    End If
                                        .Style = ArrayMarkerStyle(i + 1)
                                        '.Style = ArrayMarkerStyle(h + i + 5)
                                        With .Pen.VtColor
                                            .Blue = ColorBlue
                                            .Red = ColorRed
                                            .Green = ColorGreen
                                            End With
                                        End With
                                        i = i + 1
                                        .DataPoints.Item(-1).Brush.FillColor.Set 252, 252, 252
                                    End With
                                Next j
                            Next h
                        End If
                End If
            
           
           
           
           End With
        End If 'if X vs Y

    
    MSChart1.Height = Me.Height - cmdBack.Height / 4 - cmdBack.Height
    
    Me.MSChart1.Top = 10 * gsglXFactor
   ' Shape1.Top = 0
   ' Shape1.Left = 0
    MSChart1.Left = 10 * gsglXFactor
   ' Shape1.Width = MSChart1.Width + 20 * gsglXFactor
   ' Shape1.Height = MSChart1.Height + 20 * gsglXFactor

    'MSChart1.Title.Font.Name = "Trebuchet MS"
    MSChart1.Title.VtFont.Size = 16
    '''''''''''''''
    
    '''''''''''''''
    Unload frmWait
    Screen.MousePointer = vbDefault
    Err.Clear
    On Error Resume Next
    myrsDATA.Close
    myrsMydata.Close
    Unload frmDocument
Exit Sub
Error1:
    Unload frmWait
    Screen.MousePointer = vbDefault
    MsgBox Err.Description & " in frmGraph_Load."

End Sub


Public Sub MakeDataArray()
Dim myrs As Recordset
Dim i As Integer
Dim j As Integer
Dim FieldName As String
Dim StartValue()  As Single
    
    On Error GoTo Error1
     
    If ShowRealDates = 0 Then
        
        Set myrs = dbXbuild.OpenRecordset("Select * From [Graph_Data] Where [Date] = 0")
        With myrs
            If Not (.EOF And .BOF) Then
                
                .MoveLast
                .Edit
                !Date = 1
                .Update
            'find the increment here
            If .RecordCount > 1 Then
                .MovePrevious
            End If
                For j = 1 To .Fields.Count - 3
                    ReDim Preserve StartValue(j + 1)
                    FieldName = .Fields(j + 2).Name
                    If InStr(FieldName, "Run") <> 0 _
                        And UCase(Mid(FieldName, 4, 1)) = "C" Then
                        StartValue(j) = Round(.Fields(j + 2).Value, 4)
                    Else
                        StartValue(j) = 0
                    End If
                Next j
            Else
                Exit Sub
            End If
            .Close
        End With
        dbXbuild.Execute "Delete * From Graph_Data Where [Date] = 0"
        Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data")
        With myrs
            .MoveFirst
            .Edit
            !Date = 0
            .Update
            For j = 1 To .Fields.Count - 3
                .MoveFirst
                Do While Not .EOF
                    .Edit
                    .Fields(j + 2).Value = .Fields(j + 2).Value - StartValue(j)
                    .Update
                    .MoveNext
                Loop
            Next j
        End With
    End If
    
    Err.Clear
    On Error Resume Next
    myrs.Close
    Exit Sub
Error1:     MsgBox Err.Description & " in MakeDataArray."
    
End Sub

'Public Sub MakePlotData2()
'This sub makes X -axis even
'Dim NewDate() As Date
'Dim NewDate2()
'Dim i As Integer
'Dim j As Integer
'Dim theSmallestDifference As Integer
'Dim prmRange As Integer
'Dim k As Integer
'Dim prmDate()
'Dim m As Integer
'Dim myrsDATA As Recordset
'Dim myrs As Recordset
'Dim nn As Integer
'Dim ArrayNumber()
'Dim ArrayDate() As Date
'Dim ArrayDate2() As Integer
'Dim PlotData()
'Dim mmm
'    On Error GoTo Error1
 '       If ShowRealDates = 1 Then
 '           Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Order By RealDate")
 '       ElseIf ShowRealDates = 0 Then
 '           Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Order By Date")
 '       ElseIf ShowRealDates = 2 Then
 '           Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Order By Date")
 '       End If
 '      With myrs
 '           If Not (.EOF And .BOF) Then
 '               .MoveLast
 '               NumberOfRecords = .RecordCount
 '               NumberOfVariables = .Fields.Count - 3
 '           End If
 '           ReDim ArrayNumber(1 To NumberOfVariables + 1, 1 To NumberOfRecords + 1)
 '           ReDim ArrayDate(1 To NumberOfRecords + 1)
 '           ReDim ArrayDate2(1 To NumberOfRecords + 1)
 '           .MoveFirst
 '           i = 1
  '          If FileType <> "OUT" Then
 '               If ShowRealDates = 0 Then
 '                   ShowRealDates = 1
 '               End If
 '           End If
 '           Do While Not .EOF
 '               If ShowRealDates = 1 Then
 '                   ArrayDate(i) = .Fields(2).Value
 '               ElseIf ShowRealDates = 0 Then
 '                   ArrayDate2(i) = .Fields(1).Value
 '               ElseIf ShowRealDates = 2 Then
 '                   ArrayDate2(i) = .Fields(1).Value
 '               End If
 '               i = i + 1
 '               .MoveNext
 '           Loop
 '           .MoveFirst
 '           i = 2
 '           j = 1
 '       For i = 1 To NumberOfVariables
 '           Do While Not .EOF
 '               ArrayNumber(i, j) = .Fields(i + 2).Value
 '               j = j + 1
'                .MoveNext
 '           Loop
'            j = 1
 '           .MoveFirst
'        Next i
 '   End With
    
  '  If FileType <> "OUT" Then
  '      If ShowRealDates = 0 Then
  '          ShowRealDates = 1
  '      End If
  '  End If
  '  If ShowRealDates = 1 Then
  '      If FileType = "OUT" Then
  '          prmRange = DateDiff("d", Format(ArrayDate(1), "mm/dd/yyyy"), Format(ArrayDate(NumberOfRecords), "mm/dd/yyyy"))
   '         theSmallestDifference = 1
   '         j = 1
   '         i = 2
   '         k = 2
   '         m = 1
   '         ReDim NewDate(2)
    '        ReDim prmDate(2)
   '         ReDim PlotData(1 To NumberOfRecords + prmRange, 1 To NumberOfVariables)
   '         prmDate(1) = ArrayDate(1)
   '         NewDate(1) = ArrayDate(1)
   '         For j = 1 To NumberOfVariables
   '             PlotData(1, j) = ArrayNumber(j, 1)
   '         Next j
   '         i = 2
   '         Do While DateDiff("d", Format(ArrayDate(NumberOfRecords), "mm/dd/yyyy"), Format(prmDate(m), "mm/dd/yyyy")) < 0
   '             ReDim Preserve prmDate(i + 1)
   '             ReDim Preserve NewDate(m + 2)
   '             prmDate(i) = DateAdd("d", theSmallestDifference, Format(prmDate(i - 1), "mm/dd/yyyy"))
   '             m = m + 1
   '             nn = DateDiff("d", Format(prmDate(i), "mm/dd/yyyy"), Format(ArrayDate(k), "mm/dd/yyyy"))
   '             If DateDiff("d", Format(prmDate(i), "mm/dd/yyyy"), Format(ArrayDate(k), "mm/dd/yyyy")) = 0 Then
   '                 For j = 1 To NumberOfVariables
  '                      NewDate(m) = ArrayDate(k)
  '                      PlotData(m, j) = ArrayNumber(j, k)
  '                  Next j
  '                  k = k + 1
  '                  If ArrayDate(k - 1) = ArrayDate(k) Then
  '                      i = i - 1
   '                 End If
  '              ElseIf DateDiff("d", Format(prmDate(i), "mm/dd/yyyy"), Format(ArrayDate(k), "mm/dd/yyyy")) > 0 Then
  '                  For j = 1 To NumberOfVariables
  '                      NewDate(m) = prmDate(i)
 '                       PlotData(m, j) = Null
 '                   Next j
 '               Else
 '                   For j = 1 To NumberOfVariables
 '                       NewDate(m) = prmDate(i)
 '                       PlotData(m, j) = Null
 '                   Next j
 '                   k = k + 1
 '               End If
 '               i = i + 1
 '           Loop
  '          TotalNumberDatPlot = m
 '           For j = 1 To NumberOfVariables
  '              NewDate(TotalNumberDatPlot) = ArrayDate(NumberOfRecords)
  '              PlotData(TotalNumberDatPlot, j) = ArrayNumber(j, NumberOfRecords)
  '          Next j
  '
  '      Else 'T-file
  '          NumberExpVariables = NumberOfVariables
  '          ReDim NewDate(1 To NumberOfRecords)
  '          TotalNumberDatPlot = NumberOfRecords
   '         ReDim PlotData(1 To NumberOfRecords, 1 To NumberOfVariables)
  '          For j = 1 To NumberOfVariables
  '              For i = 1 To NumberOfRecords
  '                  NewDate(i) = ArrayDate(i)
  '                  PlotData(i, j) = ArrayNumber(j, i)
  '              Next i
  '          Next j
   '     End If 'If FileType = "OUT" Then
  '  End If 'If ShowRealDates = 1 Then
    
  '  If ShowRealDates = 0 Or ShowRealDates = 2 Then 'Days after Planting
 '      ' prmRange = ArrayDate2(NumberOfRecords) - ArrayDate2(1)
 '       theSmallestDifference = 1
 '       j = 1
 '       i = 2
 '       k = 2
 '       m = 1
 '       ReDim NewDate2(2)
 '       ReDim prmDate(2)
 '       ReDim PlotData(1 To 1000, 1 To NumberOfVariables)
 '       prmDate(1) = ArrayDate2(1)
 '       NewDate2(1) = ArrayDate2(1)
 '       For j = 1 To NumberOfVariables
 '           PlotData(1, j) = ArrayNumber(j, 1)
 '       Next j
'        i = 2
 '       'err.clear
 '       Do While ArrayDate2(NumberOfRecords) - prmDate(m) >= 0
'            ReDim Preserve prmDate(m + 1)
  '          ReDim Preserve NewDate2(m + 2)
 '           prmDate(i) = prmDate(i - 1) + theSmallestDifference
 '           m = m + 1
 '           nn = ArrayDate2(k) - prmDate(i)
 '           If ArrayDate2(k) - prmDate(i) = 0 Then
 '               For j = 1 To NumberOfVariables
 '                   NewDate2(m) = ArrayDate2(k)
 '                   PlotData(m, j) = ArrayNumber(j, k)
 '               Next j
 '               k = k + 1
'                If ArrayDate2(k - 1) = ArrayDate2(k) Then
'                    i = i - 1
'                End If
'            ElseIf ArrayDate2(k) - prmDate(i) > 0 Then
 '               For j = 1 To NumberOfVariables
 '                   NewDate2(m) = prmDate(i)
 '                   PlotData(m, j) = Null
 '               Next j
 '           Else
 '               For j = 1 To NumberOfVariables
 '                   NewDate2(m) = prmDate(i)
 '                   PlotData(m, j) = Null
'                Next j
 '               k = k + 1
 '           End If
 '           i = i + 1
 '       Loop
 '       TotalNumberDatPlot = m
 '   End If 'ShowRealDates <> 1 Then 'Days after Planting

'     dbXbuild.Execute "Delete * From Graph_Data"
'     Set myrsDATA = dbXbuild.OpenRecordset("Graph_Data")
'     With myrsDATA
'        For i = 1 To TotalNumberDatPlot
'            .AddNew
'            '!TheOrder = i
 '           If ShowRealDates = 1 Then
 '               !realdate = NewDate(i)
 '           ElseIf ShowRealDates = 0 Or ShowRealDates = 2 Then
 '               !Date = NewDate2(i)
 '           End If
 '           .Update
 '       Next i
  '      For j = 1 To NumberOfVariables
 '           .MoveFirst
  '          For i = 1 To TotalNumberDatPlot
 '               .Edit
 '               .Fields(2 + j).Value = PlotData(i, j)
 '               .Update
 '               If Not .EOF Then
 '                   .MoveNext
 '               End If
 '           Next i
 '       Next j
 '   End With
 '   Err.Clear
 '   On Error Resume Next
 '   myrsDATA.Close
 '   Exit Sub
'Error1:     MsgBox Err.Description & i & " in MakePlotData."
'End Sub


Public Sub MakeDataAdjusted()
'This sub will adjust scales for plotting data
Dim myrs As Recordset
Dim myrs2 As Recordset
Dim i As Integer
Dim mySQL As String
Dim arrMaxValues()
Dim MaxValue As Single
Dim myFraction As Single
Dim j
Dim k As Integer
Dim myFraction2
Dim X As Integer
Dim prmPlotData()
Dim FieldNumber As Integer
Dim RecordNumber As Integer
Dim prmVariableName() As String
Dim prmNumberOfVariables As Integer
Dim prmTotalNumberDatPlot As Single
Dim Variable_RemovedXAxis1 As String
Dim Variable_RemovedXAxis2 As String
If FileType <> "SUM" Then
    If ShowRealDates = 1 Then
        Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Order by RealDate")
    ElseIf ShowRealDates = 0 Then
        Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Order by Date")
    ElseIf ShowRealDates = 2 Then
        Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Order by Date")
    End If
Else
    Set myrs = dbXbuild.OpenRecordset("Graph_Data")
End If

On Error Resume Next
dbXbuild.Execute "Drop Table [Adjusted_Graph_Data]"
Err.Clear
mySQL = ""
On Error GoTo Error1
    With myrs
        If Not (.EOF And .BOF) Then
            .MoveLast
            ReDim prmPlotData(1 To .RecordCount, 1 To .Fields.Count - 3)
            ReDim prmVariableName(1 To .Fields.Count - 3)
            ReDim NewVariableName(1 To .Fields.Count - 3)
            prmNumberOfVariables = .Fields.Count - 3
            prmTotalNumberDatPlot = .RecordCount
            For i = 1 To prmNumberOfVariables
            'For i = 1 To NumberOfVariables
                mySQL = mySQL & "Max([" & .Fields(i + 2).Name & "]) As [Max_" & .Fields(i + 2).Name & "], "
            Next i
            If mySQL <> "" Then
               mySQL = Mid(mySQL, 1, Len(mySQL) - 2)
            End If
            
            For FieldNumber = 1 To .Fields.Count - 3
                    .MoveFirst
                    RecordNumber = 1
                    Do While Not .EOF
                        prmPlotData(RecordNumber, FieldNumber) = .Fields(FieldNumber + 2).Value
                        'Debug.Print RecordNumber, FieldNumber, prmPlotData(RecordNumber, FieldNumber)
                        .MoveNext
                        RecordNumber = RecordNumber + 1
                    Loop
                prmVariableName(FieldNumber) = .Fields(FieldNumber + 2).Name
            Next FieldNumber
        Else
            Exit Sub
        End If
        
    End With
    Set myrs2 = dbXbuild.OpenRecordset("Select " & _
        mySQL & " From [Graph_Data];")
        myrs2.MoveLast
    Dim arrMaxValuesNAMES() As String
    Dim arrMaxValuesNAMES_New As String
    ReDim arrMaxValues(1 To prmNumberOfVariables)
    ReDim arrMaxValuesNAMES(1 To prmNumberOfVariables + 1)
    ReDim arrMaxValues_New(1 To prmNumberOfVariables)
    Dim AB_arrMaxValues
    Dim AB_arrMaxValuesNAMES
    ReDim AB_arrMaxValues(1 To prmNumberOfVariables + 1)
    ReDim AB_arrMaxValuesNAMES(1 To prmNumberOfVariables + 1)
    
    If FileType <> "SUM" Then
        With myrs2
            For i = 1 To prmNumberOfVariables
                arrMaxValues(i) = .Fields(i - 1).Value
                arrMaxValuesNAMES(i) = Replace(Mid(.Fields(i - 1).Name, 1, _
                    InStr(.Fields(i - 1).Name, "(") - 1), "Max_", "")
            Next i
            .Close
        End With
    Else
        With myrs2
            For i = 1 To prmNumberOfVariables
                arrMaxValues(i) = .Fields(i - 1).Value
                arrMaxValuesNAMES(i) = Replace(.Fields(i - 1).Name, "Max_X_Axis_", "")
                arrMaxValuesNAMES(i) = Replace(arrMaxValuesNAMES(i), "Max_", "")
            Next i
            .Close
        End With
    End If
    
    myrs.Close
    Dim bb As Integer
    Dim aa As String
    aa = arrMaxValuesNAMES(1)
    bb = 1
    Dim nn As Integer
    For nn = 1 To prmNumberOfVariables
        If Mid(arrMaxValuesNAMES(nn), 1, 4) <> "sss_" Then
            For i = 1 To prmNumberOfVariables
                If Mid(arrMaxValuesNAMES(i), 1, 4) <> "sss_" Then
                    If arrMaxValuesNAMES(i) = aa Then
                        AB_arrMaxValuesNAMES(bb) = aa
                        AB_arrMaxValues(bb) = arrMaxValues(i)
                        arrMaxValuesNAMES(i) = "sss_" & arrMaxValuesNAMES(i)
                        bb = bb + 1
                    End If
                End If
            Next i
            aa = arrMaxValuesNAMES(nn + 1)
        End If
    Next nn
    For i = 1 To prmNumberOfVariables
        arrMaxValuesNAMES(i) = Replace(arrMaxValuesNAMES(i), "sss_", "")
    Next i
   Dim NameToCompare() As String
   Dim dd As Integer
   Dim dd2 As Integer
    dd = 1
    arrMaxValuesNAMES_New = AB_arrMaxValuesNAMES(1)
    For i = 1 To prmNumberOfVariables
        arrMaxValues_New(dd) = -1111
        If arrMaxValuesNAMES_New <> "" Then
            Do While arrMaxValuesNAMES_New = AB_arrMaxValuesNAMES(i)
                If AB_arrMaxValues(i) > arrMaxValues_New(dd) Then
                    arrMaxValues_New(dd) = AB_arrMaxValues(i)
                    ReDim Preserve NameToCompare(dd + 1)
                    dd2 = dd + 1
                    NameToCompare(dd) = arrMaxValuesNAMES_New
                End If
                i = i + 1
            Loop
            
        Else
            Exit For
        End If
        If i >= prmNumberOfVariables Then Exit For
        arrMaxValuesNAMES_New = AB_arrMaxValuesNAMES(i + 1)
        dd = dd + 1
    Next i
    
    
    MaxValue = -1111
   ' For i = 1 To prmNumberOfVariables
    '    If arrMaxValues(i) > MaxValue Then
     '      MaxValue = arrMaxValues(i)
      '  End If
        
    'Next i
    
    For i = 1 To prmNumberOfVariables
        If arrMaxValues_New(i) > MaxValue Then
           MaxValue = arrMaxValues_New(i)
        End If
        
    Next i
    For bb = 1 To dd2
        
        For i = 1 To prmNumberOfVariables
            
            If arrMaxValuesNAMES(i) = NameToCompare(bb) Then
                arrMaxValues(i) = arrMaxValues_New(bb)
            End If
        Next i
    Next bb
    
    
    Dim h As Integer
    j = 1
    Dim ArrayMultiplier()
    For i = 1 To prmNumberOfVariables
        ReDim Preserve ArrayMultiplier(i + 1)
        If IsNull(arrMaxValues(i)) = False Then
            If arrMaxValues(i) <> 0 Then
                myFraction = Round(MaxValue / arrMaxValues(i))
            Else
                myFraction = 1
            End If
            For h = 1 To prmNumberOfVariables
                
                If NewVariableName(h) <> "" Then
                    Variable_RemovedXAxis1 = Replace(NewVariableName(h), "X_Axis_", "")
                    Variable_RemovedXAxis2 = Replace(prmVariableName(i), "X_Axis_", "")
                    If InStr(1, Variable_RemovedXAxis1, " x ") <> 0 Then
                        Variable_RemovedXAxis1 = Trim(Mid(Variable_RemovedXAxis1, 1, _
                            InStr(1, Variable_RemovedXAxis1, " x ")))
                    End If
                    If InStr(1, Variable_RemovedXAxis2, " x ") <> 0 Then
                        Variable_RemovedXAxis2 = Trim(Mid(Variable_RemovedXAxis2, 1, _
                            InStr(1, Variable_RemovedXAxis2, " x ")))
                    End If
                    If h <> i Then
                        If FileType <> "SUM" Then
                            If Mid(Variable_RemovedXAxis1, 1, InStr(1, Variable_RemovedXAxis1, "(") - 1) = _
                                Mid(Variable_RemovedXAxis2, 1, InStr(1, Variable_RemovedXAxis2, "(") - 1) Then
                                If InStr(1, NewVariableName(h), " x ") <> 0 Then
                                    myFraction = Val(Mid(NewVariableName(h), InStr(1, NewVariableName(h), " x ") + 3))
                                Else
                                    myFraction = 1
                                End If
                                Exit For
                            End If
                        Else
                            If Mid(Variable_RemovedXAxis1, 1, Len(Variable_RemovedXAxis1) - 1) = _
                                Mid(Variable_RemovedXAxis2, 1, Len(Variable_RemovedXAxis2) - 1) Then
                                If InStr(1, NewVariableName(h), " x ") <> 0 Then
                                    myFraction = Val(Mid(NewVariableName(h), InStr(1, NewVariableName(h), " x ") + 3))
                                Else
                                    myFraction = 1
                                End If
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next h
            If myFraction >= 10 Then
                myFraction2 = myFraction
                j = 1
                
                Do While myFraction2 >= 10
                    j = 10 * j
                    myFraction2 = Round(myFraction / j)
                Loop
                
                For k = 1 To prmTotalNumberDatPlot
                    prmPlotData(k, i) = prmPlotData(k, i) * j
                Next k
                NewVariableName(i) = prmVariableName(i) & " x " & j
            Else
                For k = 1 To prmTotalNumberDatPlot
                    prmPlotData(k, i) = prmPlotData(k, i)
                Next k
                NewVariableName(i) = prmVariableName(i)
            End If
        Else
            NewVariableName(i) = prmVariableName(i)
        End If
        ArrayMultiplier(i) = j
       
       
       ' Debug.Print ArrayMultiplier(i)
    Next i
    
    Dim MinValue
    MinValue = 999999999
    For i = 1 To prmNumberOfVariables
        If ArrayMultiplier(i) < MinValue Then
           MinValue = ArrayMultiplier(i)
        End If
    Next i

   
   'Finding the min and max values
    minYvalue = 999999999
    maxYvalue = -999999999
    minXvalue = 999999999
    maxXvalue = -999999999
   
   If ShowX_Axis = 1 Then
        'find Y min and max
        For i = 1 To prmNumberOfVariables
            For k = 1 To prmTotalNumberDatPlot
                If prmPlotData(k, i) < minYvalue Then
                   minYvalue = prmPlotData(k, i)
                End If
                If prmPlotData(k, i) > maxYvalue Then
                   maxYvalue = prmPlotData(k, i)
                End If
            Next k
        Next i
    Else
        'find Y min and max
        For i = 2 To prmNumberOfVariables
            For k = 1 To prmTotalNumberDatPlot
                If prmPlotData(k, i) < minXvalue Then
                    minXvalue = prmPlotData(k, i)
                End If
                If prmPlotData(k, i) > maxXvalue Then
                    maxXvalue = prmPlotData(k, i)
                End If
            Next k
        Next i
        'find X min and max
        For k = 1 To prmTotalNumberDatPlot
            If prmPlotData(k, 1) < minYvalue Then
                minYvalue = prmPlotData(k, 1)
            End If
            If prmPlotData(k, 1) > maxYvalue Then
                maxYvalue = prmPlotData(k, 1)
            End If
        Next k
    End If
    
     
     dbXbuild.Execute "Create Table [Adjusted_Graph_Data] ([TheOrder] Text, " & _
     "[Date] Single, [RealDate] Date)"
    For i = 1 To prmNumberOfVariables
        dbXbuild.Execute "Alter Table [Adjusted_Graph_Data] Add Column [" & NewVariableName(i) & "] " & _
        "Single"
    Next i
    If FileType <> "SUM" Then
        If ShowRealDates = 1 Then
            Set myrs2 = dbXbuild.OpenRecordset("Select * From Graph_Data Order by RealDate")
        ElseIf ShowRealDates = 0 Then
            Set myrs2 = dbXbuild.OpenRecordset("Select * From Graph_Data Order by Date")
        ElseIf ShowRealDates = 2 Then
            Set myrs2 = dbXbuild.OpenRecordset("Select * From Graph_Data Order by Date")
        End If
    Else
        Set myrs2 = dbXbuild.OpenRecordset("Graph_Data")
    End If
    Set myrs = dbXbuild.OpenRecordset("Adjusted_Graph_Data")
    If myrs2.EOF And myrs2.BOF Then
        myrs2.Close
        Exit Sub
    End If
    k = 1
    myrs2.MoveFirst
    Do While Not myrs2.EOF
        With myrs
            .AddNew
            If ShowRealDates = 1 Then
                !realdate = myrs2!realdate
            ElseIf ShowRealDates = 0 Then
                !Date = myrs2!Date
            ElseIf ShowRealDates = 2 Then
                !Date = myrs2!Date
            End If
            For i = 1 To prmNumberOfVariables
                .Fields(i + 2).Value = prmPlotData(k, i)
            Next i
            .Update
        End With
        k = k + 1
        myrs2.MoveNext
    Loop
    Err.Clear
    On Error Resume Next
    myrs2.Close
    myrs.Close
    Exit Sub
Error1:
    MsgBox Err.Description & bb & " in MakeDataAdjusted"
End Sub


Public Sub FillNullDataInOut2()
'This sub adds values to make the line continuous
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim X As Integer
Dim h As Integer
Dim ValueDifference
Dim n As Integer
Dim TotalNumberDatPlot1 As Integer
Dim TheLastDateInArray()
Dim myrs As Recordset
Dim prmNumberExpVariables  As Integer
Dim myPlotData()
Dim myExpPlotData()
Dim myNewDate()
Dim meNewDate2()
Dim myNumberOfVariables
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
    ElseIf ShowRealDates = 0 Or ShowRealDates = 2 Then
        Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Order By Date")
    End If
    myNumberOfVariables = myrs.Fields.Count - 3
    With myrs
        If Not (.EOF And .BOF) Then
            .MoveLast
            TotalNumberDatPlot1 = .RecordCount
            ReDim myPlotData(TotalNumberDatPlot1 + 1, myNumberOfVariables - prmNumberExpVariables + 1)
            ReDim myExpPlotData(TotalNumberDatPlot1 + 1, prmNumberExpVariables + 1)
            ReDim myNewDate(TotalNumberDatPlot1 + 1)
            ReDim myNewDate2(TotalNumberDatPlot1 + 1)
            ReDim TheLastDateInArray(myNumberOfVariables - prmNumberExpVariables + 1)
            i = 1
            .MoveFirst
            Do While Not .EOF
                If ShowRealDates = 1 Then
                    myNewDate(i) = !realdate
                ElseIf ShowRealDates = 0 Then
                    myNewDate2(i) = !Date
                ElseIf ShowRealDates = 2 Then
                    myNewDate2(i) = !Date
                End If
                i = i + 1
                .MoveNext
            Loop
            For j = 1 To myNumberOfVariables - prmNumberExpVariables
                i = 1
                .MoveFirst
                Do While Not .EOF
                    myPlotData(i, j) = .Fields(j + 2).Value
                    i = i + 1
                    .MoveNext
                Loop
                TheLastDateInArray(j) = MyLastDate(.Fields(j + 2).Name)
            
            Next j
            For j = 1 To prmNumberExpVariables
                i = 1
                .MoveFirst
                Do While Not .EOF
                    myExpPlotData(i, j) = .Fields(myNumberOfVariables - prmNumberExpVariables + 2 + j).Value
                    i = i + 1
                    .MoveNext
                Loop
            Next j
        End If
        .Close
    End With
    For j = 1 To myNumberOfVariables - prmNumberExpVariables
        n = 2
        Do While IsNull(myPlotData(n, j)) = True
            n = n + 1
        Loop
        If n > 2 Then
            n = n - 1
        End If
        For i = n To TotalNumberDatPlot1
            If ShowRealDates = 1 Then
                If DateDiff("d", Format(myNewDate(i), "mm/dd/yyyy"), Format(TheLastDateInArray(j), "mm/dd/yyyy")) < 0 Then
                    Exit For
                End If
            ElseIf ShowRealDates = 0 Or ShowRealDates = 2 Then
                If -myNewDate2(i) + Val(TheLastDateInArray(j)) < 0 Then
                    Exit For
                End If
            End If

            k = 1
            If IsNull(myPlotData(i, j)) = True Then
                Do While IsNull(myPlotData(i + k, j)) = True
                    k = k + 1
                Loop
                If IsNull(myPlotData(i - 1, j)) = True And IsNull(myPlotData(i + k, j)) = False Then
                    myPlotData(i - 1, j) = myPlotData(i + k, j)
                End If
                
                If myPlotData(i - 1, j) <= myPlotData(i + k, j) Then
                    ValueDifference = (myPlotData(i + k, j) - myPlotData(i - 1, j)) / (k + 1)
                    For X = 1 To k
                        myPlotData(i + X - 1, j) = myPlotData(i - 1, j) + X * ValueDifference
                    Next X
                Else
                    ValueDifference = (myPlotData(i - 1, j) - myPlotData(i + k, j)) / (k + 1)
                    For X = 1 To k
                        myPlotData(i + X - 1, j) = myPlotData(i - 1, j) - X * ValueDifference
                    Next X
                End If
            End If
            i = i + k - 1
        Next i
    Next j
    
    dbXbuild.Execute "Delete * From Graph_Data"
    Set myrs = dbXbuild.OpenRecordset("Graph_Data")
    With myrs
        For i = 1 To TotalNumberDatPlot
            .AddNew
            If ShowRealDates = 1 Then
                !realdate = myNewDate(i)
            ElseIf ShowRealDates = 0 Or ShowRealDates = 2 Then
                !Date = myNewDate2(i)
            End If
            .Update
        Next i
        .Close
    End With
    If ShowRealDates = 1 Then
        Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Order by RealDate")
    ElseIf ShowRealDates = 0 Then
        Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Order by Date")
    ElseIf ShowRealDates = 2 Then
        Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Order by Date")
    End If
    With myrs
        .MoveFirst
        For j = 1 To myNumberOfVariables - prmNumberExpVariables
            .MoveFirst
            For i = 1 To TotalNumberDatPlot
                .Edit
                .Fields(2 + j).Value = myPlotData(i, j)
                .Update
                .MoveNext
            Next i
        Next j
        For j = 1 To prmNumberExpVariables
            .MoveFirst
            For i = 1 To TotalNumberDatPlot
                .Edit
                .Fields(myNumberOfVariables - prmNumberExpVariables + 2 + j).Value = myExpPlotData(i, j)
                .Update
                .MoveNext
            Next i
        Next j
    Err.Clear
    On Error Resume Next
    myrs.Close
    End With
    Exit Sub
Error1:    MsgBox Err.Description & " in FillNullDataInOut."
End Sub


Public Sub FindMatchingPairs()
'This sub find the experimental variables that corresponds to out - to match color
Dim myrs As Recordset
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim NumberXVariables As Integer
Dim prmNumberOfVariables As Integer
Dim prmVariableName() As String
    
    On Error GoTo Error1
   ' If Show_Sim = 0 Then
   '     NumberMatchingPairs = 0
   '     ShowStatistic = 0
    '    prmShowLine = 0
    '    Exit Sub
    'End If
    
    If ExpData_vs_Simulated = 1 Or FileType = "SUM" Then
        Set myrs = dbXbuild.OpenRecordset("Graph_Data")
        With myrs
            'prmNumberOfVariables = .Fields.Count - 3
            NumberMatchingPairs = (.Fields.Count - 3) / 2
            ReDim ColorExpMatch(1 To NumberMatchingPairs)
            ReDim ColorOUTMatch(1 To NumberMatchingPairs)
            For i = 1 To NumberMatchingPairs
                ColorOUTMatch(i) = i
                ColorExpMatch(i) = i + NumberMatchingPairs
            Next i
        End With
        Exit Sub
    End If
    
    
    Set myrs = dbXbuild.OpenRecordset("Select * From [Exp_OUT] Where [Axis] = 'X'")
    If Not (myrs.EOF And myrs.BOF) Then
        myrs.MoveLast
        NumberXVariables = myrs.RecordCount
    Else
        NumberXVariables = 0
    End If
    
    
    myrs.Close
    Set myrs = dbXbuild.OpenRecordset("Graph_Data")
        With myrs
        prmNumberOfVariables = .Fields.Count - 3 - NumberXVariables
        ReDim prmVariableName(1 To prmNumberOfVariables)
        For i = 1 To prmNumberOfVariables
            prmVariableName(i) = .Fields(2 + NumberXVariables + i).Name
        Next i
        End With
    myrs.Close
    
    Set myrs = dbXbuild.OpenRecordset("Select * From [Exp_OUT] Where [ExpCheck] = 'YES' And [Axis] = 'Y'")
    With myrs
        If Not (.EOF And .BOF) Then
            .MoveLast
            NumberMatchingPairs = .RecordCount
            ReDim ColorExpMatch(1 To NumberMatchingPairs)
            ReDim ColorOUTMatch(1 To NumberMatchingPairs)
            .MoveFirst
            j = 1
            Do While Not .EOF
                For i = 1 To prmNumberOfVariables
                   If prmVariableName(i) = !Variable & "(" & !OUT_File & _
                        ") Run " & !Run Then
                        If ShowX_Axis = 1 Or ExpData_vs_Simulated = 1 Then
                            ColorOUTMatch(j) = i
                        Else
                            ColorOUTMatch(j) = i + NumberXVariables + 2
                        End If
                   End If
                   If prmVariableName(i) = !Variable & "(" & !ExperimentID & " " & _
                        !CropID & "T) TRT " & !TRNO & "/" & !Run Then
                        If ShowX_Axis = 1 Or ExpData_vs_Simulated = 1 Then
                            ColorExpMatch(j) = i
                        Else
                            ColorExpMatch(j) = i + NumberXVariables + 2
                        End If
                   End If
                
                Next i
                j = j + 1
                .MoveNext
            Loop
        Else
            NumberMatchingPairs = 0
            Exit Sub
        End If
        .Close
    End With
    Err.Clear
    On Error Resume Next
    myrs.Close
    Exit Sub
Error1:
    MsgBox Err.Description & " in FindMatchingPairs."

End Sub


Public Sub CreateExcelApp()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oSheet As Excel.Worksheet
Dim oSheetSTAT As Excel.Worksheet
Dim oRng As Excel.Range
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim bbb As String


    On Error GoTo Err_Handler
    Set oXL = CreateObject("Excel.Application")
    oXL.Visible = True
    Set oWB = oXL.Workbooks.Add
    Set oSheet = oWB.ActiveSheet
    
If ShowX_Axis <> 1 Or FileType = "T-file" Then
        If ShowRealDates = 1 Then
            oSheet.Cells(1, 1).Value = "Date"
         'vvv
        ElseIf ShowRealDates = 0 Then
            If CDAYexists = False And DAPexists = False Then
               oSheet.Cells(1, 1).Value = "Days after Start of Simulation"
            Else
                oSheet.Cells(1, 1).Value = "Days after Planting"
            End If
        ElseIf ShowRealDates = 2 Then
            oSheet.Cells(1, 1).Value = "Day of year"
        
        End If
        
        If FileType = "SUM" Then
            oSheet.Cells(1, 1).Value = " "
            oSheet.Range(oSheet.Cells(1, 1), oSheet.Cells(1, NumberOfVariables / 2 + 1)).Font.ColorIndex = 46
            oSheet.Range(oSheet.Cells(1, NumberOfVariables / 2 + 2), oSheet.Cells(1, NumberOfVariables + 2)).Font.ColorIndex = 41
        End If

    
    If FileType <> "SUM" Then
        For i = 1 To NumberOfVariables
            bbb = ExcelVariable(i)
            
            oSheet.Cells(1, i + 1).Value = RunName(bbb)
        Next i
    Else
        For i = 1 To NumberOfVariables
            bbb = ExcelVariable(i)
            oSheet.Cells(1, i + 1).Value = bbb
        Next i
    End If
    
    With oSheet.Range(oSheet.Cells(1, 1), oSheet.Cells(1, NumberOfVariables + 1))
        .Font.Bold = True
        .VerticalAlignment = xlVAlignCenter
    End With
    
    oSheet.Range(oSheet.Cells(2, 1), oSheet.Cells(ExcelRecNumber + 1, NumberOfVariables + 1)).Value = ExcelDataPlot
    
    Set oRng = oSheet.Range(oSheet.Cells(2, 1), oSheet.Cells(1, NumberOfVariables + 1))
    If FileType <> "SUM" Then
        oSheet.Range(oSheet.Cells(2, 1), oSheet.Cells(ExcelRecNumber + 1, 1)).AutoFormat
    End If
    oSheet.Range(oSheet.Cells(2, 1), oSheet.Cells(ExcelRecNumber + 1, 1)).Cells.Font.Bold = False
  Else
        If NumberExpVariables_Y > 0 Then
           ' For i = 1 To NumberOfVariables + 2
            For i = 1 To NumberOfVariables + 2
                bbb = ExcelVariable(i)
                'oSheet.Cells(1, i).Value = ExcelVariable(i)
                oSheet.Cells(1, i).Value = RunName(bbb)
            Next i
            With oSheet.Range(oSheet.Cells(1, 1), oSheet.Cells(1, NumberOfVariables + 2))
                .Font.Bold = True
                .VerticalAlignment = xlVAlignCenter
            End With
            oSheet.Range(oSheet.Cells(2, 1), oSheet.Cells(ExcelRecNumber + 1, NumberOfVariables + 2)).Value = ExcelDataPlot
            If FileType <> "SUM" Then
                oSheet.Range(oSheet.Cells(2, 1), oSheet.Cells(ExcelRecNumber + 1, 1)).AutoFormat
            End If
            oSheet.Range(oSheet.Cells(2, 2 + NumberExpVariables_Y), oSheet.Cells(ExcelRecNumber + 1, 2 + NumberExpVariables_Y)).AutoFormat
            With oSheet.Range(oSheet.Cells(2, 1), oSheet.Cells(2, NumberOfVariables + 2))
                .Font.Bold = False
            End With
        Else
            For i = 1 To NumberOfVariables + 1
                bbb = ExcelVariable(i)
                
                oSheet.Cells(1, i).Value = ExcelVariable(i)
            Next i
            With oSheet.Range(oSheet.Cells(1, 1), oSheet.Cells(1, NumberOfVariables + 2))
                .Font.Bold = True
                .VerticalAlignment = xlVAlignCenter
            End With
            oSheet.Range(oSheet.Cells(2, 1), oSheet.Cells(ExcelRecNumber + 1, NumberOfVariables + 1)).Value = ExcelDataPlot
            If FileType <> "SUM" Then
                oSheet.Range(oSheet.Cells(1, 1), oSheet.Cells(ExcelRecNumber + 1, 1)).AutoFormat
            End If
            With oSheet.Range(oSheet.Cells(2, 1), oSheet.Cells(2, NumberOfVariables + 2))
                .Font.Bold = False
            End With
        End If
  End If
    'remove column is sum
   ' oXL.Worksheets(1).Name = "Data"
    oSheet.Name = "Data"
    Call CreateExcelChart(oSheet)
    

    
    oXL.Worksheets(1).Name = "Data"
    
    
    If ShowStatistic = 1 Then
        Set oSheetSTAT = oWB.Worksheets.Add
        'oXL.Worksheets(2).Name = "Statistic"
        oSheetSTAT.Name = "Statistic"
        oXL.Visible = True
        
        'Set ObjWs2 = oXL.Worksheets.Add
        'Set oSheetSTAT = oWB.Worksheets(2)
        Call WriteStatistic(oSheetSTAT)
    End If
    
    

Exit_CreateChart:
      Set oRng = Nothing
      Set oSheet = Nothing
      Set oWB = Nothing
      Set oXL = Nothing
      Exit Sub
   
Err_Handler:
      MsgBox Err.Description & " in Create Excel Application.", vbCritical
     ' MsgBox "Cannot export to Excel: requires MS Office 2000 or later."
      Resume Exit_CreateChart
 
End Sub

  Private Sub CreateExcelChart(oWS As Excel.Worksheet)
    Dim oChart As Excel.Chart

    ' Add a Chart for the selected data
      On Error GoTo Err_Handler
     
      Set oChart = oWS.Parent.Charts.Add
    
     
    
    If ShowX_Axis = 1 Then
        Call Time_Chart_Create(oChart, oWS)
    Else 'X vs Y
        If ExpData_vs_Simulated = 0 Then
            Call XY_Chart_Create(oChart, oWS)
        Else
            Call ExpData_vs_Simulated_Chart_Create(oChart, oWS)
        End If
    End If
        
Exit_CreateChart:
      Set oChart = Nothing
      Exit Sub
   
Err_Handler:
      MsgBox Err.Description & " in CreateExcelChart.", vbCritical
      Resume Exit_CreateChart

End Sub


Public Sub FindPairsXY()
Dim myrsDatAPlot As Recordset
Dim myrsOUT As Recordset
Dim prmRuns() As Integer
Dim i As Integer
Dim ArrowRun As Integer
Dim k As Integer
Dim h As Integer
Dim kk As Integer
Dim prmFieldName As String
Dim TheStringLength As Integer
Dim prmTreatment As Integer
Dim prmExpName As String
    On Error GoTo Error1
    Set myrsDatAPlot = dbXbuild.OpenRecordset("Adjusted_Graph_Data")
    If FileType = "SUM" Then
        With myrsDatAPlot
            NumberYVariables = 1
            NumberXVariables = (.Fields.Count - 3) / 2
        NumberExpVariables_Y = 0
        NumberExpVariables_X = 0
        
        ReDim ArrowVariableX(NumberXVariables + 1)
        ReDim ArrowVariableY(NumberXVariables + 1, NumberXVariables + 1)

        
        k = 1
        For h = 1 To NumberXVariables
            ArrowVariableX(h) = .Fields(h + 2).Name
            ArrowVariableY(h, k) = .Fields(2 + h + NumberXVariables).Name
        Next h
        
        ''''''''''''''
        
         .Close
        End With

        Exit Sub
    End If
    
    If FileType = "OUT" And ExpData_vs_Simulated = 0 Then
        Set myrsOUT = dbXbuild.OpenRecordset("Select * From [Exp_OUT] Where " & _
            "[Axis] = 'X'")
        With myrsOUT
            If Not (.EOF And .BOF) Then
                .MoveLast
                NumberXVariables = .RecordCount
                ReDim prmRuns(1 To NumberXVariables)
                .MoveFirst
                i = 1
                Do While Not .EOF
                    prmRuns(i) = !Run
                    i = i + 1
                    .MoveNext
                Loop
            End If
        End With
        ReDim ArrowVariableX(1 To NumberXVariables)
        ReDim ArrowVariableY(1 To NumberXVariables, 1 To NumberOfVariables - NumberExpVariables)
        
        With myrsDatAPlot
            For h = 1 To NumberXVariables
                ArrowVariableX(h) = .Fields(h + 2).Name
                If InStr(ArrowVariableX(h), " X ") <> 0 Then
                    TheStringLength = InStr(ArrowVariableX(h), " X ") - _
                        InStr(ArrowVariableX(h), "Run") + 3
                Else
                    TheStringLength = Len(ArrowVariableX(h)) - InStr(ArrowVariableX(h), "Run") + 3
                End If
                ArrowRun = Val(Mid(ArrowVariableX(h), InStr(ArrowVariableX(h), "Run") + 3, _
                    TheStringLength))
                kk = 1
           
            
                For k = 3 + NumberXVariables To NumberOfVariables - NumberExpVariables + 2
                    prmFieldName = .Fields(k).Name
                    If InStr(prmFieldName, " X ") <> 0 Then
                        TheStringLength = InStr(prmFieldName, " X ") - _
                            InStr(prmFieldName, "Run") + 3
                    Else
                        TheStringLength = Len(ArrowVariableX(h)) - InStr(ArrowVariableX(h), "Run") + 3
                    End If
                
                    If ArrowRun = Val(Mid(prmFieldName, InStr(prmFieldName, "Run") + 3, TheStringLength)) Then
                        'ReDim Preserve ArrowVariableY(1 To NumberXVariables, 1 To NumberOfVariables - NumberExpVariables)
                        ArrowVariableY(h, kk) = prmFieldName
                        kk = kk + 1
                    End If
                Next k
            Next h
        End With
        NumberYVariables = kk - 1
        kk = 1
    End If
    If NumberExpVariables_X > 0 Then
        ReDim ArrowVariableX_Exp(1 To NumberExpVariables_X)
        ReDim ArrowVariableY_Exp(1 To NumberExpVariables_X, 1 To NumberExpVariables_Y)
        'ReDim ArrowVariableY_Exp(NumberExpVariables_X + 1, NumberExpVariables_Y + 1)
        With myrsDatAPlot
            For h = 1 To NumberExpVariables_X
                ArrowVariableX_Exp(h) = .Fields(2 + NumberOfVariables - NumberExpVariables_X - NumberExpVariables_Y + h).Name
                If InStr(ArrowVariableX_Exp(h), " X ") <> 0 Then
                    TheStringLength = InStr(ArrowVariableX_Exp(h), " X ") - _
                        InStr(ArrowVariableX_Exp(h), " TRT ") + 4
                Else
                    TheStringLength = Len(ArrowVariableX_Exp(h)) - _
                        InStr(ArrowVariableX_Exp(h), " TRT ") + 4
                End If
                prmTreatment = Val(Mid(ArrowVariableX_Exp(h), _
                    InStr(ArrowVariableX_Exp(h), " TRT ") + 4, TheStringLength))
                kk = 1
                For k = 1 To NumberExpVariables_Y
                    prmExpName = .Fields(2 + NumberOfVariables - _
                        NumberExpVariables_Y + k).Name
                    If InStr(prmExpName, " X ") <> 0 Then
                        TheStringLength = InStr(prmExpName, " X ") - _
                            InStr(prmExpName, " TRT ") + 4
                    Else
                        TheStringLength = Len(prmExpName) - _
                            InStr(prmExpName, " TRT ") + 4
                    End If
                    If prmTreatment = Val(Mid(prmExpName, InStr(prmExpName, " TRT ") + 4, TheStringLength)) Then
                        ArrowVariableY_Exp(h, kk) = prmExpName
                        'ReDim Preserve ArrowVariableY(1 To NumberXVariables, 1 To NumberOfVariables - NumberExpVariables + NumberExpVariables_Y / NumberExpVariables_X)
                        'ArrowVariableY(h, NumberYVariables + kk) = prmExpName
                        kk = kk + 1
                    End If
                Next k
            Next h
        End With
    ElseIf FileType = "OUT" And ExpData_vs_Simulated = 1 Then
        Set myrsOUT = dbXbuild.OpenRecordset("Select * From [Exp_OUT]")
        With myrsOUT
            If Not (.EOF And .BOF) Then
                .MoveLast
                NumberXVariables = .RecordCount
                ReDim prmRuns(1 To NumberXVariables)
                .MoveFirst
                i = 1
                Do While Not .EOF
                    prmRuns(i) = !Run
                    i = i + 1
                    .MoveNext
                Loop
            End If
        End With
        ReDim ArrowVariableX(NumberXVariables + 1)
        ReDim ArrowVariableY(NumberXVariables + 1, NumberXVariables + 1)
        
        With myrsDatAPlot
            k = 1
            For h = 1 To NumberXVariables
                ArrowVariableX(h) = .Fields(h + 2).Name
                ArrowVariableY(h, k) = .Fields(2 + h + NumberXVariables).Name
            Next h
        NumberYVariables = 1
        NumberExpVariables_Y = 0
        NumberExpVariables_X = 0
        End With
    End If
    
    Err.Clear
    On Error Resume Next
    myrsDatAPlot.Close
    myrsOUT.Close
        
    Exit Sub
Error1:     MsgBox Err.Description & " in FindPairsXY/frmGraph."
    End Sub


Public Function myColorRed(ColorIndex As Integer) As Integer
Dim nm As Integer
    nm = IIf((ColorIndex + 15) Mod 15 = 0, 15, (ColorIndex + 15) Mod 15)
    Select Case nm
        Case 1
            myColorRed = 255
        Case 10
            myColorRed = 0
        Case 3
            myColorRed = 0
        Case 4
            myColorRed = 64
        Case 5
            myColorRed = 0
        Case 6
            myColorRed = 128
        Case 7
            myColorRed = 0
        Case 8
            myColorRed = 0
        Case 9
            myColorRed = 255
        Case 2
            myColorRed = 255
        Case 11
            myColorRed = 255
        Case 12
            myColorRed = 255
        Case 13
            myColorRed = 255
        Case 14
            myColorRed = 0
        Case 15
            myColorRed = 0
    End Select
End Function

Public Function myColorBlue(ColorIndex As Integer) As Integer
Dim nm As Integer
    nm = IIf((ColorIndex + 15) Mod 15 = 0, 15, (ColorIndex + 15) Mod 15)
    Select Case nm
        Case 1
            myColorBlue = 0
        Case 10
            myColorBlue = 255
        Case 3
            myColorBlue = 0
        Case 4
            myColorBlue = 64
        Case 5
            myColorBlue = 255
        Case 6
            myColorBlue = 0
        Case 7
            myColorBlue = 0
        Case 8
            myColorBlue = 128
        Case 9
            myColorBlue = 64
        Case 2
            myColorBlue = 0
        Case 11
            myColorBlue = 0
        Case 12
            myColorBlue = 125
        Case 13
            myColorBlue = 125
        Case 14
            myColorBlue = 255
        Case 15
            myColorBlue = 64
    End Select
End Function

Public Function myColorGreen(ColorIndex As Integer) As Integer
Dim nm As Integer
    nm = IIf((ColorIndex + 15) Mod 15 = 0, 15, (ColorIndex + 15) Mod 15)
    Select Case nm
        Case 1
            myColorGreen = 0
        Case 10
            myColorGreen = 255
        Case 3
            myColorGreen = 255
        Case 4
            myColorGreen = 64
        Case 5
            myColorGreen = 64
        Case 6
            myColorGreen = 0
        Case 7
            myColorGreen = 128
        Case 8
            myColorGreen = 0
        Case 9
            myColorGreen = 64
        Case 2
            myColorGreen = 255
        Case 11
            myColorGreen = 125
        Case 12
            myColorGreen = 64
        Case 13
            myColorGreen = 125
        Case 14
            myColorGreen = 125
        Case 15
            myColorGreen = 255
    End Select
End Function



Public Sub Chart_SetUp()
    On Error GoTo Error1
    With MSChart1
        .ToDefaults
        .Left = 0
        .ChartType = VtChChartType2dXY
        .Title = PlotTitle
        .ShowLegend = True
        .Legend.Location.LocationType = VtChLocationTypeBottom
        With .Plot
            .UniformAxis = False
            .Axis(VtChAxisIdX).AxisTitle = XAxisTitle
            
            If ExpData_vs_Simulated = 1 Then
                .Axis(VtChAxisIdY).AxisTitle = "Experimental"
                .Axis(VtChAxisIdY).AxisTitle.TextLayout.VertAlignment = VtVerticalAlignmentCenter
                .Axis(VtChAxisIdY).AxisTitle.TextLayout.Orientation = VtOrientationUp
                
            End If
            
            If Show_Grid = 1 Then
                .Axis(VtChAxisIdX).AxisGrid.MajorPen.Style = VtPenStyleSolid
                .Axis(VtChAxisIdY).AxisGrid.MajorPen.Style = VtPenStyleSolid
            Else
                .Axis(VtChAxisIdX).AxisGrid.MajorPen.Style = VtPenStyleNull
                .Axis(VtChAxisIdY).AxisGrid.MajorPen.Style = VtPenStyleNull
                
            End If
            If ExpData_vs_Simulated = 1 Or FileType = "SUM" Then
                .Axis(VtChAxisIdX).AxisGrid.MajorPen.Style = VtPenStyleDotted
                .Axis(VtChAxisIdY).AxisGrid.MajorPen.Style = VtPenStyleDotted
            
            End If
            .Wall.Pen.Style = VtPenStyleNull
        End With
        With .Backdrop.Fill
            .Brush.FillColor.Set 255, 255, 255
            .Style = VtFillStyleBrush
        End With
    End With
    
    
    Exit Sub
Error1: MsgBox Err.Description & " in Chart_SetUp/frmGraph."
End Sub

Public Function ArrayMarkerStyle(StyleIndex As Integer)
Dim nm As Integer
    nm = IIf((StyleIndex + 11) Mod 11 = 0, 11, (StyleIndex + 11) Mod 11)
    Select Case nm
        Case 1
            ArrayMarkerStyle = VtMarkerStyleFilledSquare
        Case 2
            ArrayMarkerStyle = VtMarkerStyleFilledDiamond
        Case 3
            ArrayMarkerStyle = VtMarkerStyleFilledUpTriangle
        Case 4
            ArrayMarkerStyle = VtMarkerStyleFilledDownTriangle
        Case 5
            ArrayMarkerStyle = VtMarkerStylePlus
        Case 6
            ArrayMarkerStyle = VtMarkerStyleX
        Case 7
            ArrayMarkerStyle = VtMarkerStyleStar
        Case 8
            ArrayMarkerStyle = VtMarkerStyleSquare
        Case 9
            ArrayMarkerStyle = VtMarkerStyleDiamond
        Case 10
            ArrayMarkerStyle = VtMarkerStyleUpTriangle
        Case 11
            ArrayMarkerStyle = VtMarkerStyleDownTriangle
    End Select
End Function

Public Function LineStyle(StyleIndex As Integer)
Dim nm As Integer
    nm = IIf((StyleIndex + 8) Mod 8 = 0, 8, (StyleIndex + 8) Mod 8)
    Select Case nm
    Case 1
        LineStyle = VtPenStyleSolid
    Case 2
        LineStyle = VtPenStyleDotted
    Case 3
        LineStyle = VtPenStyleDashed
    Case 4
        LineStyle = VtPenStyleDashDot
    Case 5
        LineStyle = VtPenStyleDashDotDot
    Case 6
        LineStyle = VtPenStyleDitted
    Case 7
        LineStyle = VtPenStyleDashDit
    Case 8
        LineStyle = VtPenStyleDashDitDit
    End Select
End Function


Public Sub XY_Chart_Create(prmChart As Chart, prmoWS As Excel.Worksheet)
Dim MyColorIndex As Integer
Dim nn As Integer
Dim i As Integer
Dim bb As Integer
Dim MarksColorMatch() As Integer
Dim StartForExp As Integer

    On Error GoTo Err_Handler
    nn = 1
    With prmChart.PlotArea.Interior
        .ColorIndex = 2
        .PatternColorIndex = 1
        .Pattern = xlSolid
    End With

    
    With prmChart
        .ChartType = xlXYScatter
        If FileType = "SUM" And Excel_Number_Y_OUT_Variables = 1 Then
            .SetSourceData Source:=prmoWS.Range("B2:C" & ExcelRecNumber + 1), PlotBy _
                :=xlColumns
        Else
            .SetSourceData Source:=prmoWS.Range("B2:C" & ExcelRecNumber + 1), PlotBy _
                :=xlColumns
        End If
        'For i = 1 To Excel_Number_Y_OUT_Variables + Excel_Number_Y_Exp_Variables - 1
         If FileType <> "SUM" Then
            For i = 1 To Excel_Number_Y_OUT_Variables + Excel_Number_Y_Exp_Variables - 1
                .SeriesCollection.NewSeries
            Next i
        Else
            For i = 1 To Excel_Number_Y_OUT_Variables + Excel_Number_Y_Exp_Variables - 1
                .SeriesCollection.NewSeries
            Next i
        End If
        ReDim MarksColorMatch(1 To Excel_Number_X_OUT_Variables)
        For bb = 1 To Excel_Number_X_OUT_Variables
            MyColorIndex = Round(Rnd() * 10)
            If FileType = "SUM" Then
                MyColorIndex = 1
            End If
            
            MarksColorMatch(bb) = MyColorIndex
            For i = 1 To Excel_Number_Y_OUT_Variables / Excel_Number_X_OUT_Variables
                .SeriesCollection(nn).XValues = "=Data!R2C" & (bb + 1) & _
                    ":R" & ExcelRecNumber + 1 & "C" & (bb + 1)
                .SeriesCollection(nn).Values = "=Data!R2C" & _
                    (1 + Excel_Number_X_OUT_Variables + _
                    (Excel_Number_Y_OUT_Variables / Excel_Number_X_OUT_Variables) * (bb - 1) + i) & _
                    ":R" & ExcelRecNumber + 1 & "C" & _
                    (1 + Excel_Number_X_OUT_Variables + _
                    (Excel_Number_Y_OUT_Variables / Excel_Number_X_OUT_Variables) * (bb - 1) + i)
                With .SeriesCollection(nn)
                    .Name = prmoWS.Cells(1, _
                        1 + Excel_Number_X_OUT_Variables + _
                        (Excel_Number_Y_OUT_Variables / Excel_Number_X_OUT_Variables) * (bb - 1) + i)
                    If FileType = "SUM" Then
                        .MarkerBackgroundColorIndex = xlNone
                    Else
                        .MarkerBackgroundColorIndex = MarksColorMatch(bb)
                    End If
                    .MarkerForegroundColorIndex = MarksColorMatch(bb)

                End With
                nn = nn + 1
            Next i
        Next bb
        bb = 1
        If Excel_Number_X_Exp_Variables > 0 Then
            StartForExp = 1 + Excel_Number_X_OUT_Variables + Excel_Number_Y_OUT_Variables
            For bb = 1 To Excel_Number_X_Exp_Variables
                For i = 1 To Excel_Number_Y_Exp_Variables / Excel_Number_X_Exp_Variables
                    .SeriesCollection(nn).Values = "=Data!R2C" & _
                        (StartForExp + Excel_Number_X_Exp_Variables + _
                        (Excel_Number_Y_Exp_Variables / Excel_Number_X_Exp_Variables) * (bb - 1) + i) & _
                        ":R" & ExcelRecNumber + 1 & "C" & _
                        (StartForExp + Excel_Number_X_Exp_Variables + _
                        (Excel_Number_Y_Exp_Variables / Excel_Number_X_Exp_Variables) * (bb - 1) + i)
                    
                     .SeriesCollection(nn).XValues = "=Data!R2C" & (bb + StartForExp) & _
                        ":R" & ExcelRecNumber + 1 & "C" & (bb + StartForExp)

                    With .SeriesCollection(nn)
                        .Name = prmoWS.Cells(1, _
                            (StartForExp + Excel_Number_X_Exp_Variables + _
                            (Excel_Number_Y_Exp_Variables / Excel_Number_X_Exp_Variables) * (bb - 1) + i))
                        .MarkerBackgroundColorIndex = MarksColorMatch(bb)
                        .MarkerForegroundColorIndex = MarksColorMatch(bb)

                    End With
                    nn = nn + 1
                Next i
            Next bb
        End If
        .HasTitle = True
        If FileType = "SUM" Then
            .ChartTitle.Characters.Text = NameSUM_File
        Else
            .ChartTitle.Characters.Text = "Scatter Plot"
        End If
        With .Axes(xlCategory, xlPrimary)
            .HasTitle = True
            If FileType = "SUM" Then
                .AxisTitle.Characters.Text = "Simulated Data"
                .AxisTitle.Font.ColorIndex = 3
                .TickLabels.Font.ColorIndex = 3
            Else
                .AxisTitle.Characters.Text = Trim(Mid(ExcelVariable(1), 1, InStr(ExcelVariable(1), " Run ")))
            End If
        End With
        With .Axes(xlValue, xlPrimary)
            '.HasTitle = False
            If FileType = "SUM" Then
                .HasTitle = True
                .AxisTitle.Characters.Text = "Observed Data"
                .AxisTitle.Font.ColorIndex = 5
                .TickLabels.Font.ColorIndex = 5
            Else
                .HasTitle = False
            End If
        End With
        .HasLegend = True
        With .Legend
            .Position = xlBottom
            .Border.Weight = xlHairline
            .Border.LineStyle = xlNone
        End With
        .Axes(xlValue).MajorGridlines.Delete
        With .PlotArea.Border
            .Weight = xlThin
            .LineStyle = xlNone
        End With
    End With

Exit_CreateChart:
      Set prmChart = Nothing
      Exit Sub
Err_Handler:
      MsgBox Err.Description & " in CreateExcxelApp.", vbCritical
      Resume Exit_CreateChart
End Sub

Public Sub ExpData_vs_Simulated_Chart_Create(prmChart As Chart, prmoWS As Excel.Worksheet)
Dim MyColorIndex As Integer
Dim nn As Integer
Dim i As Integer
Dim bb As Integer

    On Error GoTo Err_Handler
    nn = 1
    With prmChart.PlotArea.Interior
        .ColorIndex = 2
        .PatternColorIndex = 1
        .Pattern = xlSolid
    End With

    With prmChart
        .ChartType = xlXYScatter
        .SetSourceData Source:=prmoWS.Range("B2:C" & ExcelRecNumber + 1), PlotBy _
            :=xlColumns
        For i = 1 To Excel_Number_Y_OUT_Variables - 1
            .SeriesCollection.NewSeries
        Next i
            For i = 1 To Excel_Number_Y_OUT_Variables
                .SeriesCollection(nn).XValues = "=Data!R2C" & (i + 1) & _
                    ":R" & ExcelRecNumber + 1 & "C" & (i + 1)
                .SeriesCollection(nn).Values = "=Data!R2C" & _
                    (1 + i + Excel_Number_Y_OUT_Variables) & _
                    ":R" & ExcelRecNumber + 1 & "C" & _
                    (1 + i + Excel_Number_Y_OUT_Variables)
                With .SeriesCollection(nn)
                    .Name = prmoWS.Cells(1, _
                        1 + i + Excel_Number_Y_OUT_Variables)
                '    If Marker_Small = 1 Then
                 '       .MarkerSize = 3
                 '   Else
                 '       .MarkerSize = 8
                  '  End If
                End With
                nn = nn + 1
            Next i
        .HasTitle = True
        .ChartTitle.Characters.Text = "Experimental Data vs. Simulated Data"
        With .Axes(xlCategory, xlPrimary)
            .HasTitle = True
           ReDim ExcelVariable(2)
           ReDim textVariables(2)
           ExcelVariable(1) = prmoWS.Cells(1, 2)
           textVariables(1) = prmoWS.Cells(1, 2)
           .AxisTitle.Characters.Text = Trim(Mid(ExcelVariable(1), 1, InStr(ExcelVariable(1), " Run ")))
        End With
        With .Axes(xlValue, xlPrimary)
            .HasTitle = False
        End With
        .HasLegend = True
        With .Legend
            .Position = xlBottom
            .Border.Weight = xlHairline
            .Border.LineStyle = xlNone
        End With
        .Axes(xlValue).MajorGridlines.Delete
        With .PlotArea.Border
            .Weight = xlThin
            .LineStyle = xlNone
        End With
    End With

Exit_CreateChart:
      Set prmChart = Nothing
      Exit Sub
Err_Handler:
      MsgBox Err.Description & " in ExpData_vs_Simulated_Chart_Create.", vbCritical
      Resume Exit_CreateChart
End Sub

Public Sub CalculateStatistic()
Dim i As Integer
Dim j As Integer
Dim OUTLable() As String
Dim ExpLable() As String
Dim myrs As Recordset
Dim myrsStat As Recordset
    ShowStatistic = 1
    On Error Resume Next
    dbXbuild.Execute ("Drop Table [Statistic_Data]")
    dbXbuild.Execute ("Drop Table [Statistic_Calculated]")
    Err.Clear
    On Error GoTo Error1
    If NumberMatchingPairs > 0 Then
        DoEvents
    Else
        EnoughData = False
        ShowStatistic = 0
        cmdStatistic.Enabled = False
        Exit Sub
    End If
    'Oxana start
    If FileType <> "SUM" Then
        If ShowRealDates = 1 Then
            Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Order By RealDate")
        ElseIf ShowRealDates = 0 Or ShowRealDates = 2 Then
            Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Order By Date")
        End If
    Else
        Set myrs = dbXbuild.OpenRecordset("Graph_Data")
    End If
    
    With myrs
        If Not (.EOF And .BOF) Then
            DoEvents
        Else
            EnoughData = False
            myrs.Close
            Exit Sub
        End If
    End With
    dbXbuild.Execute ("Create Table [Statistic_Data] ([TheOrder] integer)")
    ReDim OUTLable(1 To NumberMatchingPairs)
    ReDim ExpLable(1 To NumberMatchingPairs)
        For i = 1 To NumberMatchingPairs
            OUTLable(i) = myrs.Fields(ColorOUTMatch(i) + 2).Name
            ExpLable(i) = myrs.Fields(ColorExpMatch(i) + 2).Name
        Next i
    For i = 1 To NumberMatchingPairs
        dbXbuild.Execute ("Alter Table [Statistic_Data] Add Column [" & _
            OUTLable(i) & "] Single")
    Next i
    For i = 1 To NumberMatchingPairs
        dbXbuild.Execute ("Alter Table [Statistic_Data] Add Column [" & _
            ExpLable(i) & "] Single")
    Next i
    For i = 1 To NumberMatchingPairs
        dbXbuild.Execute ("Alter Table [Statistic_Data] Add Column [" & _
            "Ratio" & i & "] Single")
    Next i
    
    
    Set myrsStat = dbXbuild.OpenRecordset("Statistic_Data")
    
    With myrsStat
        myrs.MoveFirst
        Do While Not myrs.EOF
            .AddNew
            For i = 1 To NumberMatchingPairs
                .Fields(OUTLable(i)).Value = IIf(IsNull(myrs(OUTLable(i)).Value) = True, Null, Round(myrs(OUTLable(i)).Value, 3)) '2
                .Fields(ExpLable(i)).Value = IIf(IsNull(myrs(ExpLable(i)).Value) = True, Null, Round(myrs(ExpLable(i)).Value, 3)) '2
                .Fields("Ratio" & i).Value = myRatio(myrs(OUTLable(i)).Value, myrs(ExpLable(i)).Value)
            Next i
            .Update
            myrs.MoveNext
        Loop
    End With
    myrsStat.Close
    Dim myTotalCount()
    Dim myrsTotal As Recordset
    ReDim myTotalCount(NumberMatchingPairs + 1)
    'Paul
    Dim NonRoundedAverage()
    ReDim NonRoundedAverage(NumberMatchingPairs + 1)
    For i = 1 To NumberMatchingPairs
        Set myrsTotal = dbXbuild.OpenRecordset("SELECT Count([" & ExpLable(i) & "]) AS [Count_Calculated] FROM [Statistic_Data]")
        With myrsTotal
            If Not (.EOF And .BOF) Then
                .MoveFirst
                myTotalCount(i) = .Fields("Count_Calculated").Value
            End If
            .Close
        End With
        dbXbuild.Execute ("Update [Statistic_Data] Set [" & ExpLable(i) & _
            "] = Null Where [" & ExpLable(i) & "] = 0")
        dbXbuild.Execute ("Update [Statistic_Data] Set [" & OUTLable(i) & _
            "] = Null Where [" & ExpLable(i) & "] is Null")
        dbXbuild.Execute ("Update [Statistic_Data] Set [" & ExpLable(i) & _
            "] = Null Where [" & OUTLable(i) & "] is Null")
    Next i
    myrs.Close
    'dbXbuild.Execute ("Create Table [Statistic_Calculated] ([Variable] Text, [Mean_Obs] Single, " _
        & "[Mean_Sim] Single, [Mean_Ratio] Single, [STDEV_Obs] Single, [STDEV_Sim] Single, " _
        & "[STDEV_Ratio] Single, [MEAN_Diff] Single,[MEAN_ABS_Diff] Single, [RMSE] Single, " _
        & "[D_stat] Single, [Obs_Number] Single,[Total_Obs_Number] Single)")
    dbXbuild.Execute ("Create Table [Statistic_Calculated] ([Variable] Text, [Mean_Obs] Single, " _
        & "[Mean_Sim] Single, [Mean_Ratio] Single, [STDEV_Obs] Single, [STDEV_Sim] Single, " _
        & "[R_Square] Single, [MEAN_Diff] Single,[MEAN_ABS_Diff] Single, [RMSE] Single, " _
        & "[D_stat] Single, [Obs_Number] Single,[Total_Obs_Number] Single)")

    Set myrs = dbXbuild.OpenRecordset("Statistic_Calculated")
    For i = 1 To NumberMatchingPairs
        With myrs
            .AddNew
            'Change name later
            '!Variable = SelectedVariable(OUTLable(i))
            Dim mmOUT As String
            Dim mmEXP As String
            mmOUT = OUTLable(i)
            mmEXP = ExpLable(i)
            '''NumberOfDecimals
            Dim NearestDigital As Integer
            NearestDigital = NumberOfDecimals(mmOUT)
            !Variable = StatVariable(SelectedVariable(mmOUT))
            '!Variable = myVariableName(SelectedVariable(mmOUT), SelectedVariable(mmEXP))
            '!Variable = myVariableName(SelectedVariable(OUTLable(i)), SelectedVariable(ExpLable(i)))
            NonRoundedAverage(i) = myAverage(ExpLable(i))
            !Mean_Obs = Round(NonRoundedAverage(i), NearestDigital)
            '!Mean_Obs = myAverage(ExpLable(i))
            !Mean_Sim = Round(myAverage(OUTLable(i)), NearestDigital)
            !Mean_Ratio = Round(myAverage("Ratio" & i), 3)
            !STDEV_Obs = mySTDEV(ExpLable(i))
            !STDEV_Sim = mySTDEV(OUTLable(i))
            '!STDEV_Ratio = mySTDEV("Ratio" & i)
            !R_Square = myRSquare(OUTLable(i), ExpLable(i))
            !MEAN_Diff = Round(myMeanDiff(OUTLable(i), ExpLable(i)), NearestDigital)
            !MEAN_ABS_Diff = Round(myMEAN_ABS_Diff(OUTLable(i), ExpLable(i)), NearestDigital)
            !RMSE = myRMSE(OUTLable(i), ExpLable(i))
            !D_stat = myD_stat(OUTLable(i), ExpLable(i), NonRoundedAverage(i))
            !Obs_Number = MyCount(ExpLable(i))
            !Total_Obs_Number = myTotalCount(i)
            .Update
        End With
    Next i
    
    Dim mymy
   ' mymy = myRSquare(OUTLable(1), ExpLable(1)
    Err.Clear
    On Error Resume Next
    dbXbuild.Execute ("Drop Table [Statistic_Data]")

    Exit Sub
Error1:
    MsgBox Err.Description & " in CalculateStatistic/frmGraph."
End Sub

Public Function myAverage(prmVariableName As String)
Dim myrs As Recordset
    On Error GoTo Error1
    Set myrs = dbXbuild.OpenRecordset("SELECT Avg([" & prmVariableName & "]) AS [Average_Calculated] FROM [Statistic_Data]")
    With myrs
        If Not (.EOF And .BOF) Then
           .MoveFirst
            myAverage = .Fields("Average_Calculated").Value
            
        End If
        .Close
    End With
    Exit Function
Error1: MsgBox Err.Description & " in myAverage/frmGraph."
End Function

Public Function myRatio(prmDataSim, prmDataObs)
    On Error GoTo Error1
    If IsNull(prmDataSim) = False And IsNull(prmDataObs) = False Then
        If prmDataObs = 0 Then
            myRatio = Null
        Else
            myRatio = Round(prmDataSim / prmDataObs, 3)
        End If
    Else
        myRatio = Null
    End If
    Exit Function
Error1: MsgBox Err.Description & " in myRatio/frmGraph."
End Function

Public Function mySTDEV(prmVariableName As String)
Dim myrs As Recordset
    On Error GoTo Error1
    Set myrs = dbXbuild.OpenRecordset("SELECT StDevP([" & prmVariableName & "]) AS [STDV_Calculated] FROM [Statistic_Data]")
    With myrs
        If Not (.EOF And .BOF) Then
           .MoveFirst
            mySTDEV = Round(.Fields("STDV_Calculated").Value, 3)
        End If
        .Close
    End With
    Exit Function
Error1: MsgBox Err.Description & " in mySTDEV/frmGraph."
End Function

Public Function myMeanDiff(prmDataSim, prmDataObs)
Dim myrs As Recordset
    On Error GoTo Error1
    Set myrs = dbXbuild.OpenRecordset("SELECT Avg([" & prmDataSim & "] - [" & _
       prmDataObs & "]) AS [MeanDiff_Calculated] FROM [Statistic_Data]")
    With myrs
        If Not (.EOF And .BOF) Then
           .MoveFirst
            myMeanDiff = Round(.Fields("MeanDiff_Calculated").Value, 3)
        End If
        .Close
    End With
    Exit Function
Error1: MsgBox Err.Description & " in myMeanDiff/frmGraph."
End Function

Public Function myMEAN_ABS_Diff(prmDataSim, prmDataObs)
Dim myrs As Recordset
    On Error GoTo Error1
    Set myrs = dbXbuild.OpenRecordset("SELECT Avg(ABS(([" & prmDataSim & "] - [" & _
       prmDataObs & "]))) AS [MEAN_ABS_Diff_Calculated] FROM [Statistic_Data]")
    With myrs
        If Not (.EOF And .BOF) Then
           .MoveFirst
            myMEAN_ABS_Diff = Round(.Fields("MEAN_ABS_Diff_Calculated").Value, 3)
        End If
        .Close
    End With
    Exit Function
Error1: MsgBox Err.Description & " in myMEAN_ABS_Diff/frmGraph."
End Function

Public Function myRMSE(prmDataSim, prmDataObs)
Dim myrs As Recordset
    On Error GoTo Error1
    'Paul
        Set myrs = dbXbuild.OpenRecordset("SELECT Avg(([" & prmDataSim & "] - [" & _
            prmDataObs & "])^2) AS [RMSE_Calculated] FROM [Statistic_Data]")
        With myrs
            If Not (.EOF And .BOF) Then
            .MoveFirst
                If IsNull(.Fields("RMSE_Calculated").Value) = True Then
                    myRMSE = Null
                Else
                    myRMSE = Round(Sqr(.Fields("RMSE_Calculated").Value), 3)
                End If
            End If
            .Close
        End With
    Exit Function
Error1: MsgBox Err.Description & " in myRMSE/frmGraph."
End Function
'Paul
Public Function myD_stat(prmDataSim, prmDataObs, prmMeanObs)
Dim myrs As Recordset
Dim Array_prmDataSim() As Single
Dim Array_prmDataObs() As Single
Dim Array_Sim_Obs_Diff() As Single
Dim i As Integer
Dim TheCount As Integer
Dim TheSum As Single
Dim TheSum2 As Single
    On Error GoTo Error1
    Set myrs = dbXbuild.OpenRecordset("Statistic_Data")
    With myrs
        If Not (.EOF And .BOF) Then
            i = 1
            .MoveFirst
            Do While Not .EOF
                If IsNull(.Fields(prmDataSim)) = False And IsNull(.Fields(prmDataObs)) = False Then
                    ReDim Preserve Array_prmDataSim(i + 1)
                    ReDim Preserve Array_prmDataObs(i + 1)
                    ReDim Preserve Array_Sim_Obs_Diff(i + 1)
                    Array_prmDataSim(i) = .Fields(prmDataSim).Value
                    Array_prmDataObs(i) = .Fields(prmDataObs).Value
                    Array_Sim_Obs_Diff(i) = Array_prmDataSim(i) - Array_prmDataObs(i)
                    i = i + 1
                End If
                .MoveNext
            Loop
        End If
        TheCount = i - 1
    End With
    TheSum = 0
    TheSum2 = 0
    For i = 1 To TheCount
        TheSum = TheSum + (Array_Sim_Obs_Diff(i)) ^ 2
    Next i
    
    For i = 1 To TheCount
        TheSum2 = TheSum2 + (Abs(Array_prmDataSim(i) - prmMeanObs) + _
            Abs(Array_prmDataObs(i) - prmMeanObs)) ^ 2
    Next i
    If TheSum2 = 0 Then
        myD_stat = Null
    Else
        myD_stat = Round(1 - TheSum / TheSum2, 3) '4
    End If
    Exit Function
Error1: MsgBox Err.Description & " in myD_stat/frmGraph."
End Function


Public Function myRSquare(prmDataSim, prmDataObs)
Dim myrs As Recordset
Dim Array_prmDataSim() As Single
Dim Array_prmDataObs() As Single
Dim Array_Sim_Obs_Diff() As Single
Dim i As Integer
Dim TheCount As Long
Dim TheSum As Single
Dim TheSum2 As Single
Dim prmSUM_Squared As Single
Dim prmSUM_SimObs As Single
Dim mySUM_Squared_Averaged As Single
Dim mySUM_D_Sim As Single
Dim mySUM_D_Obs As Single
Dim prmSumD_X_Y As Single
Dim prmAvgSum_Obs_Sum_Sim As Single
Dim mySum_Sim As Single
Dim mySum_Obs As Single
'Dim myAvg As Single


    'mySUM_D_Sim=SUM(Y2)-((SUM(Y))2)/n
    'mySUM_D_Obs=SUM(X2)-((SUM(X))2)/n
    'prmSUM_SimObs=SUM(X*Y)
    'prmAvgSum_Obs_Sum_Sim=(SUM(X)* SUM(Y))/n
    'prmSumD_X_Y=prmSUM_SimObs-prmAvgSum_Obs_Sum_Sim
    
    On Error GoTo Error1
    Set myrs = dbXbuild.OpenRecordset("Select * From [Statistic_Data] Where [" & prmDataSim & _
        "] is Not NULL And [" & prmDataObs & "] is Not NULL")
    
    With myrs
        If Not (.EOF And .BOF) Then
            .MoveLast
            TheCount = .RecordCount
        Else
            myRSquare = Null
            Exit Function
        End If
        .Close
    End With
    
    'Simulated
    Set myrs = dbXbuild.OpenRecordset("SELECT SUM(([" & prmDataSim & "])^2) AS [SUM_Squared] FROM [Statistic_Data]" & _
        " Where [" & prmDataObs & "] Is Not NULL and [" & prmDataSim & "] Is Not NULL")
    With myrs
        If IsNull(!SUM_Squared) = True Then
            prmSUM_Squared = -99
        Else
            prmSUM_Squared = !SUM_Squared
        End If
        .Close
    End With
    
    Set myrs = dbXbuild.OpenRecordset("SELECT (SUM([" & prmDataSim & "]))^2 AS [SUM_Squared_Averaged] FROM [Statistic_Data]" & _
        " Where [" & prmDataObs & "] Is Not NULL and [" & prmDataSim & "] Is Not NULL")
    With myrs
        If IsNull(!SUM_Squared_Averaged) = True Then
            mySUM_Squared_Averaged = -99
        Else
            If TheCount <> 0 Then
                mySUM_Squared_Averaged = !SUM_Squared_Averaged / TheCount
            Else
                myRSquare = Null
                Exit Function
            End If
        End If
        .Close
    End With
    
    If prmSUM_Squared <> -99 And mySUM_Squared_Averaged <> -99 Then
        mySUM_D_Sim = prmSUM_Squared - mySUM_Squared_Averaged
        If mySUM_D_Sim = 0 Then
            myRSquare = Null
            Exit Function
        End If

    End If
    
'Observed
    Set myrs = dbXbuild.OpenRecordset("SELECT SUM(([" & prmDataObs & "])^2) AS [SUM_Squared] FROM [Statistic_Data]" & _
        " Where [" & prmDataObs & "] Is Not NULL and [" & prmDataSim & "] Is Not NULL")
    With myrs
        If IsNull(!SUM_Squared) = True Then
            prmSUM_Squared = -99
        Else
            prmSUM_Squared = !SUM_Squared
        End If
        .Close
    End With
    
    Set myrs = dbXbuild.OpenRecordset("SELECT (SUM([" & prmDataObs & "]))^2 AS [SUM_Squared_Averaged] FROM [Statistic_Data]" & _
        " Where [" & prmDataObs & "] Is Not NULL and [" & prmDataSim & "] Is Not NULL")
    With myrs
        If IsNull(!SUM_Squared_Averaged) = True Then
            mySUM_Squared_Averaged = -99
        Else
            If TheCount <> 0 Then
                mySUM_Squared_Averaged = !SUM_Squared_Averaged / TheCount
            Else
                myRSquare = Null
                Exit Function
            End If
        End If
        .Close
    End With
    
    If prmSUM_Squared <> -99 And mySUM_Squared_Averaged <> -99 Then
        mySUM_D_Obs = prmSUM_Squared - mySUM_Squared_Averaged
        If mySUM_D_Obs = 0 Then
            myRSquare = Null
            Exit Function
        End If
    End If

    'prmSUM_SimObs
    Set myrs = dbXbuild.OpenRecordset("SELECT SUM([" & prmDataObs & "] * [" & prmDataSim & "]) AS [SUM_SimObs] " & _
        "FROM [Statistic_Data] Where [" & prmDataObs & "] Is Not NULL and [" & prmDataSim & "] Is Not NULL")
    With myrs
        If IsNull(!SUM_SimObs) = True Then
            prmSUM_SimObs = -99
        Else
            prmSUM_SimObs = !SUM_SimObs
        End If
        .Close
    End With
    
    'mySUM_Obs
    Set myrs = dbXbuild.OpenRecordset("SELECT SUM([" & prmDataObs & "]) AS [mySUM_Obs] FROM [Statistic_Data]" & _
        " Where [" & prmDataObs & "] Is Not NULL and [" & prmDataSim & "] Is Not NULL")
    With myrs
        If IsNull(!mySum_Obs) = True Then
            mySum_Obs = -99
        Else
            mySum_Obs = !mySum_Obs
        End If
        .Close
    End With
    
     'mySUM_Sim
    Set myrs = dbXbuild.OpenRecordset("SELECT SUM([" & prmDataSim & "]) AS [mySUM_Sim] FROM [Statistic_Data]" & _
        " Where [" & prmDataObs & "] Is Not NULL and [" & prmDataSim & "] Is Not NULL")
    With myrs
        If IsNull(!mySum_Sim) = True Then
            mySum_Sim = -99
        Else
            mySum_Sim = !mySum_Sim
        End If
        .Close
    End With
   If mySum_Sim <> -99 And mySum_Obs <> -99 Then
        prmAvgSum_Obs_Sum_Sim = (mySum_Sim * mySum_Obs) / TheCount
   End If
   
   'prmSumD_X_Y=prmSUM_SimObs-prmAvgSum_Obs_Sum_Sim
   prmSumD_X_Y = prmSUM_SimObs - prmAvgSum_Obs_Sum_Sim
   
   myRSquare = Round((prmSumD_X_Y ^ 2) / (mySUM_D_Obs * mySUM_D_Sim), 3)
   
   
    Exit Function
Error1: MsgBox Err.Description & " in myRSquare/frmGraph."
End Function

Public Function myVariableName(prmDataSim, prmDataObs)
Dim n1 As Integer
Dim n2 As Integer
Dim n3 As Integer
    On Error GoTo Error1
    n1 = InStr(prmDataSim, "(")
    n2 = InStr(prmDataSim, ")")
    n3 = InStr(prmDataObs, ")")
    If n1 <> 0 And n2 <> 0 And n3 <> 0 Then
        myVariableName = Trim(Mid(prmDataSim, 1, n1 - 1)) & " (" & _
            Trim(Mid(prmDataSim, n2 + 1)) & "/" & Trim(Mid(prmDataObs, n3 + 1)) & ")"
    Else
        myVariableName = prmDataSim
    End If
    Exit Function
Error1: MsgBox Err.Description & " in myVariableName/frmGraph."
End Function

Public Function MyCount(prmVariableName)
Dim myrs As Recordset
    On Error GoTo Error1
    Set myrs = dbXbuild.OpenRecordset("SELECT Count([" & prmVariableName & "]) AS [Count_Calculated] FROM [Statistic_Data]")
    With myrs
        If Not (.EOF And .BOF) Then
           .MoveFirst
            MyCount = .Fields("Count_Calculated").Value
        End If
        .Close
    End With
    Exit Function
Error1: MsgBox Err.Description & " in MyCount/frmGraph."
End Function

Public Sub WriteStatistic(oWS As Excel.Worksheet)
Dim myrs As Recordset
Dim i As Integer
Dim j As Integer
    On Error GoTo Error1
    With oWS
        .Range("A2").Formula = "Variable Name"
        .Range("C1").Formula = "Mean"
        .Range("E1").Formula = "Std.Dev."
        .Range("H2").Formula = "Mean Diff."
        .Range("I2").Formula = "Mean Abs.Diff."
        .Range("J2").Formula = "RMSE"
        .Range("K2").Formula = "d-Stat."
        .Range("L2").Formula = "Used Obs."
        .Range("M2").Formula = "Total Number Obs."
        .Range("B2").Formula = "Observed"
        .Range("C2").Formula = "Simulated"
        .Range("D2").Formula = "Ratio"
        .Range("E2").Formula = "Observed"
        .Range("F2").Formula = "Simulated"
        .Range("G2").Formula = "r-Square"
        
        With .Range("A1: M2")
            .Font.Size = 10
            .Font.Bold = True
        End With
        Set myrs = dbXbuild.OpenRecordset("Statistic_Calculated")
        If Not (myrs.EOF And myrs.BOF) Then
            myrs.MoveFirst
            i = 1
            j = 1
            Do While Not myrs.EOF
                For j = 1 To myrs.Fields.Count
                    oWS.Cells(i + 2, j) = myrs.Fields(j - 1).Value
                Next j
                i = i + 1
            myrs.MoveNext
            Loop
        End If
        With .Range(oWS.Cells(1, 1), oWS.Cells(myrs.RecordCount + 2, 1))
            .Columns.AutoFit
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End With
        With .Range(oWS.Cells(1, 4), oWS.Cells(myrs.RecordCount + 2, 4))
            '.Columns.AutoFit
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End With
        With .Range(oWS.Cells(1, 6), oWS.Cells(myrs.RecordCount + 2, 6))
            '.Columns.AutoFit
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End With
        With .Range(oWS.Cells(1, 7), oWS.Cells(myrs.RecordCount + 2, 7))
            '.Columns.AutoFit
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End With
        
        With .Range(oWS.Cells(1, 8), oWS.Cells(myrs.RecordCount + 2, 8))
            '.Columns.AutoFit
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End With
        With .Range(oWS.Cells(1, 9), oWS.Cells(myrs.RecordCount + 2, 9))
            '.Columns.AutoFit
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End With
        With .Range(oWS.Cells(1, 10), oWS.Cells(myrs.RecordCount + 2, 10))
            '.Columns.AutoFit
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End With
        With .Range(oWS.Cells(1, 11), oWS.Cells(myrs.RecordCount + 2, 11))
           ' .Columns.AutoFit
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End With
        With .Range(oWS.Cells(1, 12), oWS.Cells(myrs.RecordCount + 2, 12))
            '.Columns.AutoFit
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End With
        With .Range(oWS.Cells(1, 13), oWS.Cells(myrs.RecordCount + 2, 13))
            '.Columns.AutoFit
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End With
        'Range("A1:L2").Select
        With .Range("A1:M2")
            '.Columns.AutoFit
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End With
        With .Range(oWS.Cells(1, 1), oWS.Cells(myrs.RecordCount + 2, 13))
            '.Columns.AutoFit
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End With
    End With
    Exit Sub
Error1:
MsgBox Err.Description & " in WriteStatistic/frmGraph. "
End Sub


Public Function MyLastDate(myName As String)
Dim TheRun As Integer
Dim TheFile As String
Dim myrs As Recordset
    On Error Resume Next
    TheFile = Trim(Mid(myName, InStr(myName, "(") + 1, (InStr(myName, ")") - InStr(myName, "(")) - 1)) & "_OUT"
    TheRun = Val(Mid(myName, InStr(myName, "Run") + 3))
    
    Set myrs = dbXbuild.OpenRecordset("Select * From [" & TheFile & _
        "] Where [RunNumber] = " & TheRun)
    With myrs
        If Not (.EOF And .BOF) Then
            .MoveLast
            If ShowRealDates = 1 Then
                MyLastDate = Format(TheDate(!Date), "mmm d yyyy")
            ElseIf CDAYexists = True And ShowRealDates = 0 Then
                MyLastDate = !CDAY
            ElseIf DAPexists = True And ShowRealDates = 0 Then
                MyLastDate = !DAP
            ElseIf DAPexists = False And CDAYexists = False And ShowRealDates = 0 Then
                MyLastDate = !DAS
            ElseIf ShowRealDates = 2 Then
                MyLastDate = Mid(!Date, 6)
            End If
        End If
        .Close
    End With
End Function

Public Function RunName(prmNewName As String) As String
Dim myRun As Integer
Dim myFile As String
Dim RunPlace As Integer
Dim VarName As String
Dim prmMultiplier As String
Dim myrs As Recordset
    On Error GoTo Error1
        If prmNewName = "Day" Or prmNewName = "" Or prmNewName = "Date" Then
            RunName = prmNewName
            Exit Function
        End If
        
        If InStr(prmNewName, " x ") <> 0 Then
           prmMultiplier = Mid(prmNewName, InStr(prmNewName, " x "))
        Else
            prmMultiplier = ""
        End If
    If InStr(prmNewName, "TRT ") = 0 Then
        RunPlace = InStr(1, prmNewName, "Run ")
        myRun = Val(Mid(prmNewName, RunPlace + 3))
        Dim sss1 As String
        Dim sss2 As String
        Dim sss3 As String
        Dim nnn1 As Integer
        Dim nnn2 As Integer
        nnn1 = InStr(prmNewName, "Run ")
        nnn2 = InStrRev(prmNewName, "(") + 1
        sss1 = Trim(Mid(prmNewName, nnn2, nnn1 - nnn2 - 2))
        'myFile = Trim(Mid(prmNewName, InStr(prmNewName, "(") + 1, RunPlace - 3 - InStr(prmNewName, "(")))
        myFile = sss1
        VarName = Trim(Mid(prmNewName, 1, InStr(2, prmNewName, "(" & myFile) - 1))
       ' VarName = Trim(Mid(prmNewName, 1, InStr(prmNewName, "(") - 1))
        Set myrs = dbXbuild.OpenRecordset("Select * From " & myFile & "_File_Info Where RunNumber = " & _
            myRun)
        If Not (myrs.BOF And myrs.EOF) Then
            myrs.MoveFirst
            RunName = VarName & " (" & myrs!RunDescription & ")" & prmMultiplier
            myrs.Close
        End If
    Else
        If InStr(16, prmNewName, "/") <> 0 Then
            RunName = Mid(prmNewName, 1, InStr(16, prmNewName, "/") - 1) & prmMultiplier
        Else
            RunName = TreatmentDescription(prmNewName)
        End If
    End If
    Exit Function
Error1: MsgBox Err.Description & " in RunName/frmGraph."
End Function


Public Function TreatmentDescription(prmName As String)
Dim myrs As Recordset
Dim t_FileName  As String
Dim t_Treatment As Integer
Dim myVarName As String
    On Error GoTo Error1
    t_FileName = Mid(prmName, InStr(prmName, "(") + 1)
    t_FileName = Mid(t_FileName, 1, InStr(t_FileName, ")") - 1)
    t_Treatment = Val(Mid(prmName, InStr(prmName, "TRT") + 4))
    myVarName = Mid(prmName, 1, InStr(prmName, "(") - 1)
    Set myrs = dbXbuild.OpenRecordset("Select * From [" & t_FileName & _
        "_File_Info]" & " Where TRNO = " & t_Treatment)
    With myrs
        .MoveFirst
        If Trim(!Description) <> "" Then
            TreatmentDescription = myVarName & " " & "(" & Trim(!Description) & ")"
        Else
            TreatmentDescription = myVarName & " " & "( TRT " & !TRNO & " )"
        End If
        .Close
    End With
    Exit Function
Error1: MsgBox Err.Description & " in TreatmentDescription/frmGraph."
End Function

Public Function NumberOfDecimals(prmNameOfVariable As String)
Dim myrs As Recordset
Dim NameOfTable As String
Dim NameOfVariable As String
Dim PlaceOfParansis1 As Integer
Dim PlaceOfParansis2 As Integer
Dim LengthOfTableName As Integer
Dim i As Integer
Dim TryNext As Boolean

    On Error GoTo Error1
    If FileType <> "SUM" Then
        PlaceOfParansis1 = InStr(1, prmNameOfVariable, "(")
        PlaceOfParansis2 = InStr(1, prmNameOfVariable, ")")
        LengthOfTableName = PlaceOfParansis2 - PlaceOfParansis1
        NameOfVariable = Mid$(prmNameOfVariable, 1, PlaceOfParansis1 - 1)
        NameOfTable = Mid$(prmNameOfVariable, PlaceOfParansis1 + 1, LengthOfTableName - 1)
        Set myrs = dbXbuild.OpenRecordset(NameOfTable & "_OUT")
    Else
        NameOfVariable = prmNameOfVariable
        NameOfTable = "Graph_Data"
        Set myrs = dbXbuild.OpenRecordset(NameOfTable)
    End If
    
    TryNext = True
    For i = 0 To 5
        If TryNext = True Then
            With myrs
                If Not (.EOF And .BOF) Then
                    .MoveFirst
                    Do While Not .EOF
                        If IsNull(.Fields(NameOfVariable)) = False Then
                            If Round(.Fields(NameOfVariable), i) - .Fields(NameOfVariable) <> 0 Then
                                TryNext = True
                                Exit Do
                            Else
                                TryNext = False
                                NumberOfDecimals = i
                            End If
                        End If
                        .MoveNext
                    Loop
                Else
                    NumberOfDecimals = 0
                    GoTo ExitFunction
                End If
            End With
        End If
    Next i

ExitFunction:
    myrs.Close
    Exit Function
Error1:
MsgBox Err.Description & " in NumberOfDecimals/frmGraph"
End Function

Public Function StatVariable(prmVarName As String)
Dim Number_OfRun As String
Dim NameOfVariable As String
Dim PlaceOfParansis1 As Integer
Dim PlaceOfParansis2 As Integer
Dim LengthOfTableName As Integer
Dim Cut_NameOfVariable As String
Dim Length_Cut_NameOfVariable As Integer
Dim NumberSpacesTO_add As Integer
Dim i As Integer
Dim New_NameOfVariable_With_Spaces As String

    On Error GoTo Error1
    If FileType <> "SUM" Then
        PlaceOfParansis1 = InStr(1, prmVarName, "(")
        PlaceOfParansis2 = InStr(1, prmVarName, ")")
        LengthOfTableName = PlaceOfParansis2 - PlaceOfParansis1
        NameOfVariable = Mid$(prmVarName, 1, PlaceOfParansis1 - 1)
        Number_OfRun = Mid$(prmVarName, PlaceOfParansis2 + 1)
    
        Cut_NameOfVariable = Mid$(NameOfVariable, 1, 20)
        Length_Cut_NameOfVariable = Len(Cut_NameOfVariable)
        NumberSpacesTO_add = 20 - Length_Cut_NameOfVariable
        New_NameOfVariable_With_Spaces = Cut_NameOfVariable
        If NumberSpacesTO_add > 0 Then
            For i = 0 To NumberSpacesTO_add
                New_NameOfVariable_With_Spaces = New_NameOfVariable_With_Spaces & " "
            Next i
        End If
        StatVariable = New_NameOfVariable_With_Spaces & "(" & Number_OfRun & ")"
    Else
        StatVariable = prmVarName
    End If
    Exit Function
Error1:
MsgBox Err.Description & " in StatVariable/frmGraph"

End Function

'Public Sub MakePlotData1()
'Dim myrs As Recordset
'Dim TheLastDay As Integer
'Dim TheLastDate As Date
'Dim TheFirstDay As Integer
'Dim TheFirstDate As Date
'Dim myrsSelect As Recordset
'Dim i As Integer
'Dim DifferenceInLastDayAndFirstDay As Integer
'Dim NextDay As Date
'    On Error GoTo Error1
'        If ShowRealDates = 1 Then
'            Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Order By RealDate")
'            With myrs
 '               If Not (.EOF And .BOF) Then
 '                   .MoveFirst
  '                  TheFirstDate = .Fields("RealDate").Value
  '                  .MoveLast
  '                  TheLastDate = .Fields("RealDate").Value
  '              End If
  '          End With
        
    '        DifferenceInLastDayAndFirstDay = DateDiff("d", Format(TheFirstDate, "mm/dd/yyyy"), Format(TheLastDate, "mm/dd/yyyy"))
   '         For i = 1 To DifferenceInLastDayAndFirstDay
   '             NextDay = DateAdd("d", i, Format(TheFirstDate, "mm/dd/yyyy"))
   '             Set myrsSelect = dbXbuild.OpenRecordset("Select * From Graph_Data where RealDate = #" _
   '                 & NextDay & "#")
  '              If myrsSelect.EOF And myrsSelect.BOF Then
  '                  With myrs
   '                     .AddNew
   '                     .Fields("RealDate") = NextDay
  '                      .Update
  '                  End With
  '              End If
  '          Next i
        
    '    ElseIf ShowRealDates = 0 Or ShowRealDates = 2 Then
   '         Set myrs = dbXbuild.OpenRecordset("Select * From Graph_Data Order By Date")
   '         With myrs
   '             If Not (.EOF And .BOF) Then
  '                  .MoveFirst
  '                  Do While Not .EOF
  '                      If .Fields("Date").Value = 0 Then
  '                          .MoveNext
  '                      Else
  '                          Exit Do
  '                      End If
  '                  Loop
  '                  TheFirstDay = .Fields("Date").Value
  '                  .MoveLast
  '                  TheLastDay = .Fields("Date").Value
  '              End If
  '          End With
  '          For i = TheFirstDay To TheLastDay
  '              Set myrsSelect = dbXbuild.OpenRecordset("Select * From Graph_Data where Date = " & i)
  '              If myrsSelect.EOF And myrsSelect.BOF Then
  '                  With myrs
  '                      .AddNew
   '                     .Fields("Date") = i
    '                    .Update
    '                End With
     '           End If
       '     Next i
     '   End If
    '   'With myrs
  '  Exit Sub
'Error1:     MsgBox Err.Description & " in MakePlotData1."
'End Sub


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
                    'If IsNull(a(i + cc)) = False And a(i + cc) <> Empty Then
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
           'Dim NumberOFtail As Integer
            '''''''''''''''
            With myrs
          '      .MoveFirst
             '   'MyDAyDay
             '   .MoveLast
             '   NumberOFtail = 0
             '   Do While Not IsNull(.Fields(j).Value) = True
             '       If .Fields(1).Value = FirstPlantDay(j - 2) Then Exit Do
             '       NumberOFtail = NumberOFtail + 1
              '      .MovePrevious
             '   Loop
             '   NumberOfMyrs_Records = NumberOfMyrs_Records - NumberOFtail
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


Public Sub FillNullDataInOut_expVSsim()
Dim myrs As Recordset
Dim NumberOfMyrs_Records As Integer
Dim myNumberOfVariables
Dim j As Integer
Dim a()
Dim b()
Dim i As Integer

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
        myNumberOfVariables = (myrs.Fields.Count - 3) / 2
        
        For j = 1 To NumberOUTVariables
            ReDim a(NumberOfMyrs_Records + 10)
            ReDim d(NumberOfMyrs_Records + 10)
            With myrs
                .MoveFirst
                i = 1
                Do While Not (.EOF)
                    a(i) = .Fields(j + 2).Value
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
                    If IsNull(.Fields(j + 2).Value) = True Then
                        .Fields(j + 2).Value = a(i)
                    End If
                    .Update
                    i = i + 1
                    .MoveNext
                Loop
            End With

         Next j
           

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

Public Sub Time_Chart_Create(prmChart As Chart, prmoWS As Excel.Worksheet)
Dim i As Integer
Dim bb As Integer
Dim nn As Integer
Dim MyColorIndex As Integer
    On Error GoTo Err_Handler
    With prmChart.PlotArea.Interior
        .ColorIndex = 2
        .PatternColorIndex = 1
        .Pattern = xlSolid
    End With
    
    With prmChart
        .ChartType = xlXYScatter
        ''''''''''''''''
        .SetSourceData Source:=prmoWS.Range("A2:B" & ExcelRecNumber + 1), PlotBy _
            :=xlColumns
        
        If FileType <> "T-file" Then
            If ShowRealDates = 0 Then
                If NumberOfVariables > 1 Then
                    For i = 1 To NumberOfVariables - 1
                        .SeriesCollection.NewSeries
                    Next i
                End If
            Else
                If ShowRealDates = 1 Then
                    If NumberOfVariables > 2 Then
                        For i = 1 To NumberOfVariables - 2
                            .SeriesCollection.NewSeries
                        Next i
                    Else
                        For i = 1 To NumberOfVariables - 1
                            .SeriesCollection.NewSeries
                        Next i
                    End If
                Else
                    If NumberOfVariables > 2 Then
                        For i = 1 To NumberOfVariables - 2
                            .SeriesCollection.NewSeries
                        Next i
                    End If
                End If
            End If
        Else
                If NumberOfVariables > 1 Then
                    For i = 1 To NumberOfVariables - 1
                        .SeriesCollection.NewSeries
                    Next i
                End If
        End If
        If FileType <> "T-file" Then
            Err.Clear
            On Error Resume Next
            For i = 1 To NumberOfVariables - NumberExpVariables_Y
                
                .SeriesCollection(i).XValues = "=Data!R2C1" & _
                    ":R" & ExcelRecNumber + 1 & "C1"
                .SeriesCollection(i).Values = "=Data!R2C" & (i + 1) & _
                    ":R" & ExcelRecNumber + 1 & "C" & (i + 1)
                Dim bvbv As String
                bvbv = ExcelVariable(i + 1)
                .SeriesCollection(i).Name = RunName(bvbv)
                
                
                With .SeriesCollection(i).Border
                    .Weight = xlMedium
                    .LineStyle = xlAutomatic
                    
                End With
                '.SeriesCollection(i).MarkerStyle = xlNone
            Next i
            For i = NumberOfVariables - NumberExpVariables_Y + 1 To NumberOfVariables
                .SeriesCollection(i).XValues = "=Data!R2C" & (NumberOfVariables - NumberExpVariables_Y + 2) & _
                    ":R" & ExcelRecNumber + 1 & "C" & (NumberOfVariables - NumberExpVariables_Y + 2)
                .SeriesCollection(i).Values = "=Data!R2C" & (i + 2) & _
                    ":R" & ExcelRecNumber + 1 & "C" & (i + 2)
                Dim nnn As String
                nnn = ExcelVariable(i + 2)
                .SeriesCollection(i).Name = RunName(nnn)
            Next i
        Else
            For i = 1 To NumberOfVariables
                .SeriesCollection(i).XValues = "=Data!R2C1" & _
                    ":R" & ExcelRecNumber + 1 & "C1"
                .SeriesCollection(i).Values = "=Data!R2C" & (i + 1) & _
                    ":R" & ExcelRecNumber + 1 & "C" & (i + 1)
                .SeriesCollection(i).Name = ExcelVariable(i)
            Next i
        End If
        
        If ShowRealDates = 0 Then
            With .Axes(xlCategory).TickLabels
                .NumberFormat = "###"
            End With
            With .Axes(xlCategory, xlPrimary)
                .MinimumScale = prmoWS.Cells(2, 1)
                .HasTitle = True
                'vvv
                If CDAYexists = False And DAPexists = False Then
                    .AxisTitle.Characters.Text = "Days after Start of Simulation"
                Else
                    .AxisTitle.Characters.Text = "Days after Planting"
                End If

                '.AxisTitle.Characters.Text = "Days after Planting"
            End With
        ElseIf ShowRealDates = 1 Then
            With .Axes(xlCategory).TickLabels
                .Orientation = 45
                .NumberFormat = "d-mmm"
            End With
            With .Axes(xlCategory, xlPrimary)
                .MinimumScale = prmoWS.Cells(2, 1)
                .HasTitle = True
                .AxisTitle.Characters.Text = "Date"
            End With
        ElseIf ShowRealDates = 2 Then
            With .Axes(xlCategory).TickLabels
                .NumberFormat = "###"
            End With
            With .Axes(xlCategory, xlPrimary)
                .MinimumScale = prmoWS.Cells(2, 1)
                .HasTitle = True
                .AxisTitle.Characters.Text = "Day of year"
            End With
        End If
        .HasLegend = True
        With .Legend
            .Position = xlBottom
            .Border.Weight = xlMedium
            .Border.LineStyle = xlNone
        End With
        ''Grid
        .Axes(xlValue).MajorGridlines.Delete
        With .PlotArea.Border
            .Weight = xlThin
            .LineStyle = xlNone
        End With
        If FileType <> "OUT" Then
            For i = 1 To NumberOfVariables
                With .SeriesCollection(i)
                    .Smooth = False
                    .Shadow = False
                End With
            Next i
        Else
            For i = 1 To NumberOfVariables - Date_NumberExpVariables
                With .SeriesCollection(i).Border
                    .Weight = xlMedium
                    .LineStyle = xlAutomatic
                End With
                With .SeriesCollection(i)
                    MyColorIndex = Find_MyColorIndex(i)
                    .Border.ColorIndex = MyColorIndex
                    .MarkerBackgroundColorIndex = MyColorIndex
                    .MarkerForegroundColorIndex = MyColorIndex
                    .MarkerSize = 4
                    .MarkerStyle = xlCircle
                    '.MarkerStyle = xlNone
                End With
            Next i
        End If
        If NumberMatchingPairs > 0 Then
            For i = 1 To NumberMatchingPairs
                MyColorIndex = Find_MyColorIndex(i)
                'MyColorIndex = 46
                .SeriesCollection(ColorOUTMatch(i)).Border.ColorIndex = MyColorIndex
                .SeriesCollection(ColorOUTMatch(i)).MarkerBackgroundColorIndex = MyColorIndex
                .SeriesCollection(ColorOUTMatch(i)).MarkerForegroundColorIndex = MyColorIndex
                .SeriesCollection(ColorOUTMatch(i)).MarkerSize = 4
                .SeriesCollection(ColorOUTMatch(i)).MarkerStyle = xlCircle
               ' Err.Clear
                'On Error Resume Next
                With .SeriesCollection(ColorExpMatch(i))
                    .MarkerBackgroundColorIndex = MyColorIndex
                    .MarkerForegroundColorIndex = MyColorIndex
                    '.MarkerStyle = xlAutomatic
                    .MarkerSize = 4
                   'If .MarkerStyle = xlCircle Then
                        .MarkerStyle = xlTriangle
                   'End If
                End With
            Next i
        End If
    End With
Exit_CreateChart:
      Set prmChart = Nothing
      Exit Sub
Err_Handler:
      
      MsgBox Err.Description & " in Time_Chart_Create.", vbCritical
      Resume Exit_CreateChart

End Sub




Public Function Find_MyColorIndex(Index As Integer)
Dim Color_Number As Integer
    If Index > 15 Then
         Color_Number = (15 Mod Index) + Index
    Else
        Color_Number = Index
    End If
    
    Select Case Color_Number
    Case 1
        Find_MyColorIndex = 46
    Case 2
        
        Find_MyColorIndex = 55
    Case 3
        Find_MyColorIndex = 10
    Case 4
        Find_MyColorIndex = 7
    Case 5
        Find_MyColorIndex = 6
    Case 6
        Find_MyColorIndex = 45
    Case 7
        Find_MyColorIndex = 41
    Case 8
        Find_MyColorIndex = 43
    Case 9
        Find_MyColorIndex = 39
    Case 10
        Find_MyColorIndex = 12
    Case 11
        Find_MyColorIndex = 4
    Case 12
        Find_MyColorIndex = 40
    Case 13
        Find_MyColorIndex = 33
    Case 14
        Find_MyColorIndex = 56
    Case 15
        Find_MyColorIndex = 16
    
    End Select
    

End Function

Public Sub FindMaxValueScale()
    Dim myrsMydata As Recordset
    Dim myFieldName() As String
    Dim i As Integer
    Dim NumberOfMyFields As Integer
    Dim MyString As String
    Dim MaxMyField()
    Dim MinMyField()
    Dim MaxValue
    Dim MinValue
    On Error GoTo Error1
        If FileType <> "SUM" Then
            If ShowRealDates = 1 Then
                Set myrsMydata = dbXbuild.OpenRecordset("Select * From Adjusted_Graph_Data Order by RealDate")
            End If
        Else
            Set myrsMydata = dbXbuild.OpenRecordset("Adjusted_Graph_Data")
        End If
        With myrsMydata
            NumberOfMyFields = .Fields.Count - 3
            ReDim myFieldName(NumberOfMyFields)
            ReDim MaxMyField(NumberOfMyFields)
            ReDim MinMyField(NumberOfMyFields)

            For i = 0 To NumberOfMyFields - 1
                myFieldName(i) = .Fields(i + 3).Name
            Next i
            .Close
        End With
        MyString = ""
        For i = 0 To NumberOfMyFields - 1
            MyString = MyString & ", " & "Max([" & myFieldName(i) & "]) as [Max_" & myFieldName(i) & _
                "], Min([" & myFieldName(i) & "]) as [Min_" & myFieldName(i) & "]"
                
        Next i
        MyString = Trim(Mid(MyString, 2))
        'Debug.Print MyString
        Set myrsMydata = dbXbuild.OpenRecordset("Select " & MyString & _
            " From Adjusted_Graph_Data")
        'Max, Min
        
        With myrsMydata
            .MoveFirst
            For i = 0 To NumberOfMyFields - 1
                MaxMyField(i) = .Fields("Max_" & myFieldName(i)).Value
                MinMyField(i) = .Fields("Min_" & myFieldName(i)).Value
            Next i
        End With
        MaxValue = -999
        MinValue = 9999
        For i = 0 To NumberOfMyFields - 1
            If MaxMyField(i) > MaxValue Then
               MaxValue = MaxMyField(i)
            End If
            If MinMyField(i) < MinValue Then
               MinValue = MinMyField(i)
            End If
        Next i
        
        MaxAxisValue = MaxValue
        MinAxisValue = MinValue
    
    Exit Sub
Error1:
    MsgBox Err.Description & " in FindMaxValueScale/frmSelection."
End Sub

Public Sub MakeDataAdjusted_Remove_Multiplier()
Dim myrs As Recordset
Dim AdjustedFieldName() As String
Dim i As Integer
Dim NumberAdjustedFields As Integer
Dim Number_X_s As Integer
Dim AdjustedFraction() As Integer
Dim AdjustedMinimum As Double
Dim myrs_Temp As Recordset


    On Error GoTo Error1
    Set myrs = dbXbuild.OpenRecordset("Adjusted_Graph_Data")
    NumberAdjustedFields = myrs.Fields.Count - 3
    Number_X_s = 0
    
    ReDim AdjustedFieldName(NumberAdjustedFields)
    For i = 0 To NumberAdjustedFields - 1
        AdjustedFieldName(i) = myrs.Fields(i + 3).Name
    Next i
    For i = 0 To NumberAdjustedFields - 1
        If InStr(AdjustedFieldName(i), " x ") <> 0 Then
            Number_X_s = Number_X_s + 1
        End If
    Next i
    AdjustedMinimum = 999999
    If Number_X_s = NumberAdjustedFields Then
        ReDim AdjustedFraction(NumberAdjustedFields)
        
        For i = 0 To NumberAdjustedFields - 1
            AdjustedFraction(i) = Val(Mid$(AdjustedFieldName(i), InStr(AdjustedFieldName(i), " x ") + 2))
        Next i
        
        For i = 0 To NumberAdjustedFields - 1
            If AdjustedFraction(i) < AdjustedMinimum Then
                AdjustedMinimum = AdjustedFraction(i)
            End If
        Next i
        
        'myrs.Close
        dbXbuild.Execute "Create Table [Adjusted_Graph_Data_Temp] ([TheOrder] Text, " & _
            "[Date] Single, [RealDate] Date)"
        For i = 0 To NumberAdjustedFields - 1
            If AdjustedFraction(i) / AdjustedMinimum > 1 Then
                dbXbuild.Execute "Alter Table [Adjusted_Graph_Data_Temp] Add Column [" & _
                    Trim(Mid$(AdjustedFieldName(i), 1, InStr(AdjustedFieldName(i), " x "))) & " x " & _
                    AdjustedFraction(i) / AdjustedMinimum & "] " & "Single"
                NewVariableName(i + 1) = Trim(Mid$(AdjustedFieldName(i), 1, InStr(AdjustedFieldName(i), " x "))) & " x " & _
                    AdjustedFraction(i) / AdjustedMinimum
            Else
                dbXbuild.Execute "Alter Table [Adjusted_Graph_Data_Temp] Add Column [" & _
                    Trim(Mid$(AdjustedFieldName(i), 1, InStr(AdjustedFieldName(i), " x "))) & "] " & "Single"
                NewVariableName(i + 1) = Trim(Mid$(AdjustedFieldName(i), 1, InStr(AdjustedFieldName(i), " x ")))
            End If
        Next i
    Else
        myrs.Close
        Exit Sub
    End If
    Set myrs_Temp = dbXbuild.OpenRecordset("Adjusted_Graph_Data_Temp")
    
    With myrs
        .MoveFirst
        Do While Not .EOF
            myrs_Temp.AddNew
            For i = 0 To 2
               myrs_Temp.Fields(i).Value = IIf(IsNull(.Fields(i).Value) = False, .Fields(i).Value, Null)
            Next i
            For i = 0 To NumberAdjustedFields - 1
                myrs_Temp.Fields(i + 3).Value = IIf(IsNull(.Fields(i + 3).Value) = False, .Fields(i + 3).Value / AdjustedMinimum, Null)
            Next i
            myrs_Temp.Update
            .MoveNext
        Loop
    End With
    myrs_Temp.Close
    myrs.Close
    dbXbuild.Execute ("Drop Table [Adjusted_Graph_Data]")
    dbXbuild.Execute "Create Table [Adjusted_Graph_Data] ([TheOrder] Text, " & _
            "[Date] Single, [RealDate] Date)"
    
    For i = 0 To NumberAdjustedFields - 1
        If AdjustedFraction(i) / AdjustedMinimum > 1 Then
            dbXbuild.Execute "Alter Table [Adjusted_Graph_Data] Add Column [" & _
                Trim(Mid$(AdjustedFieldName(i), 1, InStr(AdjustedFieldName(i), " x "))) & " x " & _
                AdjustedFraction(i) / AdjustedMinimum & "] " & "Single"
        Else
            dbXbuild.Execute "Alter Table [Adjusted_Graph_Data] Add Column [" & _
                Trim(Mid$(AdjustedFieldName(i), 1, InStr(AdjustedFieldName(i), " x "))) & "] " & "Single"
        End If
    Next i
    dbXbuild.Execute "Insert Into [Adjusted_Graph_Data] " & _
                "Select * From [Adjusted_Graph_Data_Temp]"
      
    dbXbuild.Execute ("Drop Table [Adjusted_Graph_Data_Temp]")
    
    Exit Sub
Error1:
    MsgBox Err.Description & " in MakeDataAdjusted_Remove_Multiplier/frmGraph.)"
End Sub

Private Sub Label1_Click()

End Sub

Public Function FindExcelInstance() As Boolean
    On Error GoTo Error1
    FindExcelInstance = True
    If Dir("C:\Program Files\Microsoft Office\Office\Excel8.exe") <> "" Then
        FindExcelInstance = False
    End If
    Exit Function
Error1:
    FindExcelInstance = False
End Function

Private Sub MSChart1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    If Dir(App.path & "\Manual_GBuild.chm") <> "" Then
        App.HelpFile = App.path & "\Manual_GBuild.chm"
        Call HtmlHelp(Me.hWnd, App.HelpFile, HH_DISPLAY_TOC, 0)
    Else
        MsgBox "Cannot find" & " " & App.path & "\Manual_GBuild.chm"
    End If
 End If

End Sub

