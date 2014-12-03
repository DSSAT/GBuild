VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "GBuild"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9480
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   6480
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   53
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11060
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "07/10/2014"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "11:44 AM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1080
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "1000"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "1002"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "1012"
      End
   End
   Begin VB.Menu mnuSelection 
      Caption         =   "1065"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuOption 
      Caption         =   "1023"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "1025"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "1026"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "1027"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTechSupp 
         Caption         =   "1148"
      End
      Begin VB.Menu mnubar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "1028"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ExtentionList As String
Dim OutFileList As String



Private Sub MDIForm_Click()
    Unload frmSplash
End Sub



Private Sub MDIForm_Load()
Dim Readline As String
Dim mnOpen As Menu
Dim DriveString() As String
Dim ApplicationDrive As String
    
    LoadResStrings Me
    
    ' VSH
    Me.Caption = Me.Caption & " " & "v" & " " & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.Top = 0
    Me.Left = 0
    OpenTheFirstTime = True
    KeepSim_vs_Obs = False
    DisableExcelButton = False
    FileType = ""
    
    OneTimeDAP_Exists = 0
    CloseISselected = True
    
    If Screen.Width = 9600 Then
        gsglXFactor = 1
    Else
        gsglXFactor = Screen.Width / 9600
    End If
    
    If Screen.Height = 7200 Then
        gsglYFactor = 1
    Else
        gsglYFactor = Screen.Height / 7200
    End If
    
   ' If gsglXFactor <> 1 Or gsglYFactor <> 1 Then
        Call SetDeviceIndependentWindow(Me, 0)
   ' End If
    
    On Error GoTo ErrorMain:
    
1:    If Dir(ApplicationPathOption & "\Option.txt") <> "" Then
        Open ApplicationPathOption & "\Option.txt" For Input As #18
    Else
        MsgBox "Cannot find file" & " " & ApplicationPathOption & "\Option.txt"
        End
    End If
    Do While Not EOF(18)
        Line Input #18, Readline
        If Mid(Readline, 1, 1) <> "@" And Trim(Mid(Readline, 1, 3)) <> "" Then
            ShowExperimentalData = Val(Mid(Readline, 1, 13))
            ShowStatistic = 1
            ShowRealDates = Val(Mid(Readline, 29, 13))
            ShowX_Axis = Val(Mid(Readline, 43, 13))
            prmShowLine = Val(Mid(Readline, 57, 13))
            ExpData_vs_Simulated = Val(Mid(Readline, 71, 13))
            Show_Grid = Val(Mid(Readline, 85, 13))
            Marker_Small = Val(Mid(Readline, 99, 13))
           ' Show_Sim
            file_ShowExperimentalData = ShowExperimentalData
            file_ShowStatistic = 1
            file_ShowRealDates = ShowRealDates
            file_ShowX_Axis = ShowX_Axis
            file_prmShowLine = prmShowLine
            file_ExpData_vs_Simulated = ExpData_vs_Simulated
            file_Marker_Small = Marker_Small
            file_Show_Grid = Show_Grid
        End If
        Line Input #18, Readline
    If PaulsVersion = False Then

        If Mid(Readline, 1, 1) = "@" Then
            If Mid(Readline, 2, 1) = "1" Then
                ApplicationDrive = Mid(App.path, 1, InStr(App.path, ":") - 1)
                file_ApplicationDrive = "1"
                file_DataBasePath = ""
                If New_DataBasePath = "" Then
                    If Dir(ApplicationDrive & ":" & "\" & "DSSAT46" & "\DATA.CDE") <> "" Then
                        New_DataBasePath = ApplicationDrive & ":" & "\" & "DSSAT46"
                    ElseIf Dir(ApplicationDrive & ":" & "\" & "DSSAT5" & "\DATA.CDE") <> "" Then
                        New_DataBasePath = ApplicationDrive & ":" & "\" & "DSSAT5"
                    Else
                        MsgBox "Cannot find file" & " " & ApplicationDrive & ":" & "\" & "DSSAT46" & _
                            "\DATA.CDE" & Chr(10) & Chr(13) & "Please, select 'Option' menu button and the files in 'Data Files'."
                    End If
                End If
            Else
                If InStr(Readline, ",") <> 0 Then
                    If New_DataBasePath = "" Then
                        New_DataBasePath = Trim(Mid(Readline, 2, InStr(Readline, ",") - 2))
                        New_FilePath = Trim(Mid(Readline, InStr(Readline, ",") + 1))
                    End If
                file_DataBasePath = Trim(Mid(Readline, 2, InStr(Readline, ",") - 2))
                file_FilePath = Trim(Mid(Readline, InStr(Readline, ",") + 1))
                
                Else
                    If New_DataBasePath = "" Then
                        New_DataBasePath = Trim(Mid(Readline, 2))
                        New_FilePath = ""
                    End If
                    file_DataBasePath = Trim(Mid(Readline, 2))
                    file_FilePath = ""
                End If
            
            End If
        End If
    End If 'Paulsversion
    
    Loop
    Close 18
    
    
    CreateDataBase
    If New_DataBasePath = "" Then
       ' If Dir(Mid(App.path, 1, 1) & ":\DSSAT35\DATA.CDE") <> "" Then
       '     New_DataBasePath = Mid(App.path, 1, 1) & ":\DSSAT35"
       ' End If
        If Dir(Mid(App.path, 1, 1) & ":\DSSAT46\DATA.CDE") <> "" Then
            New_DataBasePath = Mid(App.path, 1, 1) & ":\DSSAT46"
        End If
    End If
    If Dir(New_DataBasePath & "\DATA.CDE") <> "" Then
        myReadFile (New_DataBasePath & "\DATA.CDE")
    Else
        Screen.MousePointer = vbDefault
        Dim myRespond1
        
        myRespond1 = MsgBox("Cannot find file" & " " & New_DataBasePath & "\DATA.CDE" & _
            Chr(10) & Chr(13) & "Do you want to find the file?", vbYesNo + vbQuestion)
        If myRespond1 = vbYes Then
            ''''''''''''''''
        Screen.MousePointer = vbHourglass
    
        With frmCommonDialogBox.CommonDialog1
            .CancelError = True
            On Error GoTo ProcExit
            .Filter = "All Files (*.*)|*.*"
            .FileName = ""
            .Flags = cdlOFNHideReadOnly
            .ShowOpen
            Screen.MousePointer = vbDefault
            If Len(.FileName) <> 0 Then
            If InStrRev(.FileName, "\") > 1 Then
                New_DataBasePath = Mid(.FileName, 1, InStrRev(.FileName, "\") - 1)
            End If
                MsgBox "If you want to keep" & " " & New_DataBasePath & _
                " " & "as database directory, go to Options and save the database name."
            myReadFile (New_DataBasePath & "\DATA.CDE")
        End If
    End With
        Unload frmCommonDialogBox
            Unload frmWait2
            Unload frmWait3
        Else
            Unload frmWait2
            Unload frmWait3
            End
        End If
        
    End If
    FillData 'to add codes
    ErrorDetail = False
    Screen.MousePointer = vbHourglass
    Dim SearchPath As String, FindStr As String
    Dim NumFiles As Integer
    Dim NumDirs As Integer
    Dim FileSize As Long
    Dim myrsOUT As Recordset
    Dim NameOf_OUTFILE As String
    On Error Resume Next
    dbXbuild.Execute "Drop Table OUTPUTFILES"
    Err.Clear
    On Error GoTo ErrorMain
    dbXbuild.Execute "Create Table OUTPUTFILES ([Out_files] Text)"
    
    SearchPath = New_DataBasePath & "\"
    FindStr = "OUTPUT.LST"
    FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
    If NumFiles = 0 Then
        ExtentionList = "*.*T"
        OutFileList = "*.OUT"
    Else
        ExtentionList = "*.*T"
        Set myrsOUT = dbXbuild.OpenRecordset("Select Distinct [Out_Files] From OUTPUTFILES")
        With myrsOUT
            If Not (.EOF And .BOF) Then
                OutFileList = ""
                .MoveFirst
                Do While Not .EOF
                    If IsNull(.Fields("Out_Files").Value) = False Then
                        NameOf_OUTFILE = UCase(Trim(.Fields("Out_Files").Value))
                        If UCase(Right(NameOf_OUTFILE, 3)) = "OUT" Then
                            If NameOf_OUTFILE <> "OVERVIEW.OUT" And NameOf_OUTFILE <> "SUMMARY.OUT" _
                                And NameOf_OUTFILE <> "EVALUATE.OUT" And _
                                UCase(Right(NameOf_OUTFILE, 3)) <> "OEV" Then
                            
                                OutFileList = OutFileList & "; " & NameOf_OUTFILE
                            End If
                        End If
                    End If
                    .MoveNext
                Loop
                OutFileList = Mid$(OutFileList, 3)
            Else
                ExtentionList = "*.*T"
                OutFileList = "*.OUT"
            End If
            .Close
        End With
    End If
    
    
    Screen.MousePointer = vbDefault
    If PaulsVersion = True Then
        CropName = Trim(CommandString(1)) & FileStringCommand
        
        mnuFileOpen_Click
    
    End If
    
    frmSelectionIsUnloaded = True
    
    ''''
    MainWidth = Me.Width
    MainHeight = Me.Height
    If PaulsVersion = False Then
        frmDocument.Show
        'LoadNewDoc
    End If
    ''''
    
 '   mnuOption.Enabled = False
    Exit Sub
ErrorMain:
    Unload frmWait2
    Unload frmWait3
    Close
    
    MsgBox Err.Description & " in MDIForm_Load"
ProcExit:
    End
End Sub

Private Sub LoadNewDoc()
    Static lDocumentCount As Long
    Dim frmD As frmDocument
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.Caption = "Document " & lDocumentCount
    frmD.Show
End Sub



Private Sub MDIForm_Terminate()
Dim mm
mm = 1
'*****************8
Restore_Locale
'*******************
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    

    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If

End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
       ' Case "New"
         '   LoadNewDoc
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
           ' mnuFileSave_Click
        Case "Print"
          '  mnuFilePrint_Click
        Case "Cut"
            'ToDo: Add 'Cut' button code.
            MsgBox "Add 'Cut' button code."
        Case "Copy"
            'ToDo: Add 'Copy' button code.
            MsgBox "Add 'Copy' button code."
        Case "Paste"
            'ToDo: Add 'Paste' button code.
            MsgBox "Add 'Paste' button code."
    End Select
End Sub


Private Sub mnuFile_Click()
    If FileIsClosed = True Then
        Me.mnuFileOpen.Enabled = True
    End If
End Sub

Private Sub mnuFileExit_Click()
Dim i As Integer
   On Error Resume Next
   '***************
   Dim mm
mm = 1
 
   '**************8
    dbXbuild.Close
    For i = Forms.Count - 1 To 0 Step -1
        Unload Forms(i)
    Next i
    
End Sub

Private Sub mnuFileOpen_Click()
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
Dim NumberOfSections As Integer
Dim i As Integer
Dim NumberNotOpenedFiles As Integer
Dim myrsTreat As Recordset

    Screen.MousePointer = vbHourglass
    LongFile = False
    ''''''''''''''''''''''''
Dim Readline As String
1:       If Dir(ApplicationPathOption & "\Option.txt") <> "" Then
        Open ApplicationPathOption & "\Option.txt" For Input As #18
    Else
        MsgBox "Cannot find file" & " " & ApplicationPathOption & "\Option.txt"
        End
    End If
    Do While Not EOF(18)
        Line Input #18, Readline
        If Mid(Readline, 1, 1) <> "@" And Trim(Mid(Readline, 1, 3)) <> "" Then
            ShowExperimentalData = Val(Mid(Readline, 1, 13))
            ShowStatistic = 1
            ShowRealDates = Val(Mid(Readline, 29, 13))
            ShowX_Axis = Val(Mid(Readline, 43, 13))
            prmShowLine = Val(Mid(Readline, 57, 13))
            ExpData_vs_Simulated = Val(Mid(Readline, 71, 13))
            Show_Grid = Val(Mid(Readline, 85, 13))
            Marker_Small = Val(Mid(Readline, 99, 13))
           ' Show_Sim
            file_ShowExperimentalData = ShowExperimentalData
            file_ShowStatistic = 1
            file_ShowRealDates = ShowRealDates
            file_ShowX_Axis = ShowX_Axis
            file_prmShowLine = prmShowLine
            file_ExpData_vs_Simulated = ExpData_vs_Simulated
            file_Marker_Small = Marker_Small
            file_Show_Grid = Show_Grid
        End If
        Line Input #18, Readline
    Loop
    Close 18
    
    '''''''''''''''''''''
    
    ReDim Preserve ArrayT_Tables(1)
    
    If DirectoryToPreview <> "" Then
        PaulsVersion = False
         New_FilePath = DirectoryToPreview
    End If
    FileType = ""
    CurrentFileSelected = ""
    On Error Resume Next
    NumberNotOpenedFiles = 0
    For i = 1 To NumberOfOUTfiles
        prmTheOutFileName = Mid(ArrOutFile(i), 1, Len(ArrOutFile(i)) - 4)
        dbXbuild.Execute "Drop Table [" & prmTheOutFileName & "_List] "
        dbXbuild.Execute "Drop Table [" & prmTheOutFileName & "_File_Info]"
        dbXbuild.Execute "Drop Table [" & prmTheOutFileName & "_OUT]"
    Next i
    For i = 1 To NumberOfExpFiles
        prmMyExperiment = Replace(Mid(ArrayExpFiles(i), InStr(1, _
            ArrayExpFiles(i), "\") + 1), ".", " ")
            dbXbuild.Execute "Drop Table [" & prmMyExperiment & "]"
            dbXbuild.Execute "Drop Table [" & prmMyExperiment & "_File_Info]"
            dbXbuild.Execute "Drop Table [" & prmMyExperiment & "_List]"
    Next i
    NumberOfOUTfiles = 0
    NumberOfExpFiles = 0
    Err.Clear
    
    OpenTheFirstTime = False
   '!!!!!!!
    If PaulsVersion = True Then
        FrmSelectionName = CropName
            frmWait2.Show
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
        
        Screen.MousePointer = vbHourglass
    Else
        oTHERourlIST = "*.*T"
        With dlgCommonDialog
            .CancelError = True
            On Error GoTo ProcExit
            .Filter = "Output Files|" & OutFileList & _
            "|Alt.Outputs |" & oTHERourlIST & "|T-Files |" & _
            ExtentionList & "|" _
                & "Evaluation|Evaluate.out; *.OEV|All Files |*.*"
            If New_FilePath <> "" Then
                .InitDir = New_FilePath
            Else
                .InitDir = New_DataBasePath
            End If
            If FilesName_Open = "" Then
                .FileName = ""
            Else
                .FileName = FilesName_Open
            End If
            .Flags = cdlOFNAllowMultiselect + cdlOFNHideReadOnly
            .ShowOpen
            Screen.MousePointer = vbHourglass
            If Len(.FileName) = 0 Then
                Screen.MousePointer = vbDefault
                If PaulsVersion = True Then
                     '   dbXbuild.Close
                      '  For i = Forms.Count - 1 To 0 Step -1
                       '     Unload Forms(i)
                      '  Next i
                    End
                End If
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
                Unload frmWait2
                Unload frmWait3

            CropName = .FileName
            FrmSelectionName = .FileName
            Dim myExt As String
            myExt = UCase(Mid(CropName, InStrRev(CropName, ".") + 1))
            
        End With
    End If
    
    If UCase(Mid(CropName, InStrRev(CropName, "\") + 1)) = "EVALUATE.OUT" Or _
        Right(UCase(Mid(CropName, InStrRev(CropName, "\") + 1)), 3) = "OEV" Then
        FileType = "SUM"
        frmWait2.Show
    ElseIf UCase(Mid(CropName, InStrRev(CropName, ".") + 1)) = "OUT" Then
        FileType = "OUT"
        frmWait2.Show
    ElseIf UCase(Mid(CropName, InStrRev(CropName, ".") + 3)) = "T" Then
        FileType = "T-file"
        
    Else
        FileType = "OUT"
        frmWait2.Show
    End If
    
    If FileType = "T-file" Then
        Dim ExtendedName As String
        ExtendedName = CropName
        CropName = Trim(Mid(CropName, 1, InStr(CropName, ".") + 4))
        CropName = Replace(CropName, "\ ", "\")
        If CropName <> Trim(ExtendedName) Then
            MsgBox "You may select only one T-file." & " " & _
            "File" & " " & CropName & " " & "has been selected."
        End If
    End If
    
    
    If InStr(CropName, ";") <> 0 Then
       Unload frmWait2
       Unload frmWait3

        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Err.Clear
    On Error Resume Next
    For ii = 1 To NumberOfExpFiles
        prmMyExperiment1 = Replace(Mid(ArrayExpFiles(ii), InStr(1, _
            ArrayExpFiles(ii), "\") + 1), ".", " ")
        dbXbuild.Execute "Drop Table [" & prmMyExperiment1 & "]"
    Next ii
    Err.Clear
    On Error GoTo Error1
    m = 0
    
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
        myReadFile CropName
        If EvaluateFileRead = False Then
            Unload frmWait2
            Unload frmWait3
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
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
        'UCase (Mid(CropName, InStrRev(CropName, "\") + 1))
        OUTTableNames(1) = UCase(Mid(CropName, InStrRev(CropName, "\") + 1))
       ' OUTTableNames(1) = "EVALUATE.OUT"
        'OpenFileDir = Replace(Trim(UCase(CropName)), "\EVALUATE.OUT", "")
        OpenFileDir = Replace(Trim(UCase(CropName)), "\" & OUTTableNames(1), "")
        
        If Right$(OpenFileDir, 1) = "\" Then
            OpenFileDir = OpenFileDir
        Else
            OpenFileDir = Trim(OpenFileDir) & "\"
        End If
        DirectoryToPreview = OpenFileDir
        mnuSelection.Enabled = True
        mnuSelection_Click
        prmShowLine = 0
         frmOpenFileShown = True
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    NumberOfOUTfiles = m
    If NumberOfOUTfiles = 0 Then
        MsgBox "No files selected."
        Unload frmWait2
        Unload frmWait3
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    SpacePlace(NumberOfOUTfiles) = Len(FilesString) + 1
    ReDim ArrOutFileTemp(1 To NumberOfOUTfiles + 1)
    If NumberOfOUTfiles = 1 Then
        ArrOutFileTemp(1) = FilesString
    Else
        
        For j = 0 To NumberOfOUTfiles - 1
            ArrOutFileTemp(j + 1) = Trim(Mid(FilesString, SpacePlace(j), SpacePlace(j + 1) - SpacePlace(j)))
        Next j
    End If
    ReDim ArrOutFile(NumberOfOUTfiles + 1)
    ReDim ArrOutFile(NumberOfOUTfiles + 1)
    For j = 1 To NumberOfOUTfiles
         If FileType = "OUT" Then
                ArrOutFile(j) = ArrOutFileTemp(j)
        If ObservedWasOpened = True Then
            prmShowLine = 1
            ObservedWasOpened = False
        End If
    End If
        If FileType = "T-file" Then
            
            ArrOutFile(j) = ArrOutFileTemp(j)
            prmShowLine = 0
            If ShowRealDates = 0 Then
                ShowRealDates = 1
            End If
        End If
        ObservedWasOpened = True
    Next j
    
    Err.Clear
    CannotOpenFileName = ""
    OpenFileDir = Trim(Replace(CropName, FilesString, ""))
    
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
        dbXbuild.Execute "Create Table [NamesTemp] ([Experiment] Text (25))"
        Set myrsTempNames = dbXbuild.OpenRecordset("NamesTemp")
        h = 1
        If NumberOfOUTfiles > 1 Then
            frmWait2.prgLoad.Visible = True
        Else
            frmWait2.prgLoad.Visible = False
        End If
      '  OneTimeDAP_Exists = 1

        For j = 1 To NumberOfOUTfiles
            FileToOpen = OpenFileDir & ArrOutFile(j)
            TheOutFileName = Mid(ArrOutFile(j), 1, Len(ArrOutFile(j)) - 4)
            TableName1 = TheOutFileName & "_File_Info"
            TableName2 = TheOutFileName & "_OUT"
            If NumberOfOUTfiles = 1 Then
                FileToOpen = CropName
                 DirectoryToPreview = Mid(CropName, 1, Len(CropName) - Len(ArrOutFile(1)))
            End If
            myReadFile FileToOpen
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
            
            Check_on_Years (TheOutFileName)
            
            'Call FillCreateOUTTable(TableName1, TheOutFileName, (100 / NumberOfOUTfiles) * (j - 1) + 1)
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
                            
      ' If OneTimeDAP_Exists = 0 Then
       '     DAPexists = False
       ' End If
        
        frmWait2.prgLoad.Value = 100
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
            If UCase(Trim(CannotOpenFileName)) = "EVALUATE" Then
                MsgBox "Data from files" & ": " & CannotOpenFileName & " " & "cannot be plotted together with other Output files."

            
            Else
                MsgBox "Data from files" & ": " & CannotOpenFileName & " " & "cannot be plotted with this program."
            End If
            If NumberOutTables = 0 And NumberOfExpFiles = 0 Then
                Screen.MousePointer = vbDefault
                Unload frmWait2
                Unload frmWait3
                GoTo ProcExit
            End If
        End If
        s = 0
        For j = 1 To NumberOfExpFiles
            If Dir(DirectoryToPreview & ArrayExpFiles(j)) <> "" Then
                myReadFile DirectoryToPreview & ArrayExpFiles(j)
                prmMyExperiment = ArrayExpFiles(j)
                ErrorWithExp = False
                CreateTables prmMyExperiment
                CreateExpTable prmMyExperiment
                If ErrorWithExp = False Then
                    s = s + 1
                    ReDim Preserve ArrayT_Tables(s)
                    ArrayT_Tables(s) = prmMyExperiment

                End If
                
            End If
        Next j
    End If
    Err.Clear
    On Error Resume Next
    dbXbuild.Execute "Drop Table [NamesTemp] "
    '''!!!
    NumberOfExpFiles = s
    NumberT_Tables = s
    Err.Clear
    
    If FileType = "T-file" Then
        'Selected are T-files. We are using arrays to store t-files info:
        h = 1
        NumberOfExpFiles = NumberOfOUTfiles
        ReDim ArrayExpFiles(NumberOfExpFiles + 1)
        For j = 1 To NumberOfOUTfiles
            
            ArrayExpFiles(j) = ArrOutFile(j)
            myReadFile DirectoryToPreview & ArrOutFile(j)
            ErrorExpFile = False
            CreateTables ArrOutFile(j)
            CreateExpTable ArrOutFile(j)
            If ErrorExpFile = True Then
                MsgBox "Cannot plot data from file" & " " & ArrOutFile(j)
                If PaulsVersion = True Then
                    End
                    Exit Sub
                End If
                Exit Sub
            End If
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
            dbXbuild.Execute "Delete * From [" & _
                Replace(ArrOutFile(j), ".", " ") & "_File_Info] "
                Set myrsSection = dbXbuild.OpenRecordset("Select Distinct TRNO From [" & _
                    Replace(ArrOutFile(j), ".", " ") & "]")
                FileType = "X"
                If Dir(Mid(Replace(ArrOutFile(j), " ", "."), 1, Len(ArrOutFile(j)) - 1) & "x") <> "" Then
                    myReadFile DirectoryToPreview & (Mid(Replace(ArrOutFile(j), " ", "."), 1, Len(ArrOutFile(j)) - 1) & "x")
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
    CloseISselected = False
    frmWait2.Show
    frmOpenFileShown = True
    mnuOption.Enabled = True
    mnuFileOpen.Enabled = False
    FileIsClosed = False
    Screen.MousePointer = vbDefault
    mnuSelection.Enabled = True
    
    If FileType = "T-file" Then
        Set myrsFile = dbXbuild.OpenRecordset(Replace(ArrOutFile(1), ".", " ") & "_File_Info")
        If myrsFile.EOF And myrsFile.BOF Then
            Unload frmWait3
            Unload frmWait2

            MsgBox "There is no data in the file" & " " & ArrOutFile(1) & "."
            
            GoTo Error1
        End If
        myrsFile.Close
    End If
    mnuSelection_Click

Exit Sub
Error1:
    Me.mnuFileOpen.Enabled = True
    Screen.MousePointer = vbDefault
    If Err.Number = 3078 Then GoTo 3
    Unload frmWait3
    Unload frmWait2
ProcExit:
    'MsgBox Err.Description & " in mnuFileOpen_Click."
    If PaulsVersion = True Then
        End
    End If
    Screen.MousePointer = vbDefault
End Sub




Private Sub mnuHelpAbout_Click()
    frmSplashShow = True
    frmSplash.Show
   
   ' MsgBox "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer

    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation
    Else
        On Error Resume Next
        'nRet = OSWinHelp(Me.hWnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
  
If Dir(App.path & "\Manual_GBuild.chm") <> "" Then
   App.HelpFile = App.path & "\Manual_GBuild.chm"
   Call HtmlHelp(Me.hWnd, App.HelpFile, HH_DISPLAY_TOC, 0)
 Else
    MsgBox "Cannot find" & " " & App.path & "\Manual_GBuild.chm"
 End If


End Sub



Public Sub CreateDataBase()
Dim tdfNew As TableDef
Dim wsXbuild As Workspace
Dim FieldLoop As Field
Dim DbName As String

    ' Get default Workspace.
    Set wsXbuild = DBEngine.Workspaces(0)
    On Error GoTo ErrorDbName
   'Temporary
    DbName = ApplicationPathOption + "\GraphCreateTEMP.mdb"
    
    ' Make sure there isn't already a file called SoilCreateTEMP.mdb.
     'If the file has been opened by another application then an error message is issued
     If Dir(DbName) <> "" Then
       On Error Resume Next
        Close
        Kill DbName
    End If
    ' Create a new database with the specified collating order.
    Set dbXbuild = wsXbuild.CreateDataBase(DbName, _
        dbLangGeneral)
        
    'Open DataBase and add Tables
    Set dbXbuild = wsXbuild.OpenDatabase(DbName)
    
 '   'Create LookUp tables
    
    Set tdfNew = dbXbuild.CreateTableDef("Data")
             
    With tdfNew
        ' Create fields and append them to the new TableDef
        ' object. This must be done before appending the
        ' TableDef object to the TableDefs collection of the
        ' XBuild database.

        .Fields.Append .CreateField("Type", dbText, 18)
        .Fields.Append .CreateField("Code", dbText, 10)
        .Fields.Append .CreateField("Label", dbText, 15)
        .Fields.Append .CreateField("Description", dbText)
       
        For Each FieldLoop In tdfNew.Fields
            FieldLoop.Required = False
        Next FieldLoop
        
        For Each FieldLoop In tdfNew.Fields
            FieldLoop.AllowZeroLength = True
        Next FieldLoop

        dbXbuild.TableDefs.Append tdfNew
    
   End With
   

    Exit Sub
ErrorDbName:
    MsgBox Err.Description & " in CreateSoilDataBase"
    Unload frmWait3
    Unload frmWait2
End Sub


Public Sub FillData()
Dim myrs As Recordset
Dim i As Integer
Dim Header As String
Dim prmSubstituteSymbols As String
    
    On Error GoTo Error1
    Set myrs = dbXbuild.OpenRecordset("Data")
    For i = 1 To NumberReadlines
        With myrs
            .AddNew
            If FileLoc(i) <> "" Then
                If Mid(FileLoc(i), 1, 1) <> "!" Then
                    If UCase(Mid(FileLoc(i), 1, 4)) <> "*EXP" Then
                        If Mid(FileLoc(i), 1, 1) = "*" Then
                            Header = Mid(FileLoc(i), 2, 18)
                        ElseIf Mid(FileLoc(i), 1, 1) <> "@" Then
                            .Fields("Type").Value = Header
                            prmSubstituteSymbols = RemoveSymbols_Back(Mid(FileLoc(i), 1, 7))
                            .Fields("Code").Value = prmSubstituteSymbols
                            .Fields("Label") = Replace(Mid(FileLoc(i), 8, 15), "'", "")
                            .Fields("Description").Value = Trim(Replace(Mid(FileLoc(i), 24, 55), "'", "")) & _
                                "(" & Trim(Mid(FileLoc(i), 1, 7)) & ")"
                            .Update
                        End If
                    End If
                End If
            End If
        End With
    Next i
    Err.Clear
    On Error Resume Next
    myrs.Close
    Exit Sub
Error1:
Unload frmWait2
Unload frmWait3
'MsgBox Err.Description & "in FillSata."
End
End Sub



Private Sub mnuOption_Click()
  ' If CloseISselected = False Then
        frmOptions1.Show
  ' Else
   '     MsgBox "Please open file."
   'End If
End Sub


Private Sub mnuTechSupp_Click()
   frmTecnSupp.Show
End Sub


Private Sub mnuSelection_Click()
    'If frmSelectionIsUnloaded = True Then
        frmSelection.Show
        CloseISselected = False
    'End If
End Sub


Public Sub MakingDefaultsForOpenFile()
Dim i As Integer
Dim myrs As Recordset
    On Error GoTo Error1
    Set myrs = dbXbuild.OpenRecordset("OUTPUTFILES")

    For i = 1 To NumberReadlines
        If FileLoc(i) <> "" Then
            If Mid(FileLoc(i), 1, 1) <> "!" Then
                If UCase(Mid(FileLoc(i), 1, 9)) = "@FILENAME" Then
                    i = i + 1
                    ReDim OutFileArray(2)
                    Do Until i > NumberReadlines
                        If Mid(FileLoc(i), 1, 1) = "*" Then
                            Exit Do
                        End If
                            With myrs
                                .AddNew
                                .Fields("Out_Files").Value = Trim(UCase(Mid(FileLoc(i), 1, 14)))
                                .Update
                            End With
                        i = i + 1
                    Loop
                End If
            End If
        End If
    Next i
    i = 1
    myrs.Close
    
    Exit Sub
Error1:
Unload frmWait2
Unload frmWait3
MsgBox Err.Description & "in MakingDefaultsForOpenFile."

End Sub


Function FindFilesAPI(path As String, SearchStr As String, _
        FileCount As Integer, DirCount As Integer)

Dim FileName As String ' Walking filename variable...
Dim DirName As String ' SubDirectory Name
Dim dirNames() As String ' Buffer for directory name entries
Dim nDir As Integer ' Number of directories in this path
Dim i As Integer ' For-loop counter...
Dim hSearch As Long ' Search Handle
Dim WFD As WIN32_FIND_DATA
Dim Cont As Integer

     ' Search for subdirectories.
     nDir = 0
     ReDim dirNames(nDir)
     Cont = True
     hSearch = FindFirstFile(path & "*", WFD)
     If hSearch <> INVALID_HANDLE_VALUE Then
        Do While Cont
            DirName = StripNulls(WFD.cFileName)
            ' Ignore the current and encompassing directories.
            If (DirName <> ".") And (DirName <> "..") Then
                ' Check for directory with bitwise comparison.
                If GetFileAttributes(path & DirName) And _
                    FILE_ATTRIBUTE_DIRECTORY Then
                    dirNames(nDir) = DirName
                    DirCount = DirCount + 1
                    nDir = nDir + 1
                    ReDim Preserve dirNames(nDir)
                End If
            End If
            Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
        Loop
        Cont = FindClose(hSearch)
    End If
        ' Walk through this directory and sum file sizes.
        hSearch = FindFirstFile(path & SearchStr, WFD)
        Cont = True
        If hSearch <> INVALID_HANDLE_VALUE Then
            While Cont
                FileName = StripNulls(WFD.cFileName)
                If (FileName <> ".") And (FileName <> "..") Then
                    FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * _
                    MAXDWORD) + WFD.nFileSizeLow
                    FileCount = FileCount + 1
                    myReadFile (path & FileName)
                    MakingDefaultsForOpenFile
                End If
                Cont = FindNextFile(hSearch, WFD) ' Get next file
            Wend
            Cont = FindClose(hSearch)
        End If

        ' If there are sub-directories...
        If nDir > 0 Then
            For i = 0 To nDir - 1
                FindFilesAPI = FindFilesAPI + FindFilesAPI(path & _
                dirNames(i) & "\", SearchStr, FileCount, DirCount)
            Next i
        End If
   Exit Function
Error1:
    MsgBox Err.Description & " in FindFilesAPI/frmMain."
    End Function



Public Sub Check_on_Years(prmTheOutFileName As String)
Dim myrs As Recordset
Dim myrs2 As Recordset
Dim NumberOfYears As Integer
Dim FirstYear As Single
Dim LastYear As Single

    On Error Resume Next
    dbXbuild.Execute "Drop Table FileYears"
    dbXbuild.Execute "Create Table FileYears ([Years] single)"
    Set myrs = dbXbuild.OpenRecordset(prmTheOutFileName & "_OUT")
    Set myrs2 = dbXbuild.OpenRecordset("FileYears")
    With myrs
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do While Not .EOF
                myrs2.AddNew
                myrs2.Fields("Years").Value = Val(Mid(!Date, 1, 4))
                myrs2.Update
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set myrs2 = dbXbuild.OpenRecordset("Select * From FileYears Order by Years")
    myrs2.MoveFirst
    FirstYear = myrs2!years
    myrs2.MoveLast
    LastYear = myrs2!years
    NumberOfYears = Abs(LastYear - FirstYear)
    If NumberOfYears > 3 Then
        LongFile = True
    End If
End Sub

