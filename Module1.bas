Attribute VB_Name = "Module1"
Option Explicit

Public fMainForm As frmMain
Public ApplicationPathOption As String
Public ApplicationPath As String
Public NumberReadlines As Integer
Public dbXbuild As Database
Public FileLoc() As String
Public ExcelVariable()
Public textVariables()
Public ExcelRecNumber
Public ExcelDataPlot()
Public OUTTableNames() As String
Public NumberOutTables As Integer
Public FileType As String
Public DataBasePath As String
Public file_ApplicationDrive As String
Public ShowExperimentalData As Integer
Public ShowStatistic As Integer
Public ShowRealDates As Integer
Public ShowX_Axis As Integer
Public Marker_Small As Single
Public Show_Grid As Integer
Public Show_Sim As Integer
Public ExpData_vs_Simulated As Integer
Public prmShowLine As Integer
Public file_ShowExperimentalData As Integer
Public file_ShowStatistic As Integer
Public file_ShowRealDates As Integer
Public file_ShowX_Axis As Integer
Public file_ExpData_vs_Simulated As Integer
Public file_prmShowLine As Integer
Public file_FilePath As String
Public file_DataBasePath As String
Public file_Marker_Small As Single
Public file_Show_Grid  As Integer
Public frmSelectionIsUnloaded As Boolean
Public NumberOUTVariables As Integer
Public NumberExpVariables As Integer
Public NumberExpVariables_X As Integer
Public NumberExpVariables_Y As Integer
Public Date_NumberExpVariables As Integer
Public ArrOutFile() As String
Public NumberOfOUTfiles As Integer
Public ArrayExpFiles() As String
Public NumberOfExpFiles As Integer
Public SaveConfig As Boolean
Public SaveExpShow As Boolean
Public XYVariables() As String
Public XYRuns() As String
Public XYTitle As String
Public PreviewFile As String
Public DirectoryToPreview As String
Public gsglXFactor As Single
Public gsglYFactor As Single
Public FullPath As Boolean
Public CDAYexists As Boolean
Public DAPexists As Boolean
Public ArrFile() As String
Public T_File_Y_NumberOutTables As Integer
Public T_File_X_NumberOutTables As Integer
Public Y_ArrTableNames() As String
Public X_ArrTableNames() As String
Public Excel_Number_Y_Exp_Variables As Integer
Public Excel_Number_Y_OUT_Variables As Integer
Public Excel_Number_X_Exp_Variables As Integer
Public Excel_Number_X_OUT_Variables As Integer
Public ExcelDataExist As Boolean
Public FrmSelectionName As String
Public PaulsVersion As Boolean
Public CropName As String
Public CommandString() As String
Public FileStringCommand As String
Public frmOpenFileShown As Boolean
Public New_DataBasePath As String
Public New_FilePath As String
Public CurrentFileSelected As String
Public ObservedWasOpened As Boolean
Public FileIsClosed As Boolean
Public ExperimentalDates() As String
Public Marker_Size As Single
Public FilesName_Open As String
'Main
Public TrueVariableLineSpaces() As Integer
Public LineSpaces() As Integer
Public VariablesOUTName() As String
Public NumberOfRecordsInOUTTable As Integer
Public OpenTheFirstTime As Boolean
Public StartVar As Integer
Public ErrorWithExp As Boolean
Public ErrorDetail As Boolean
Public ErrorExpFile As Boolean
Public KeepSim_vs_Obs As Boolean
Public MainHeight As Single
Public MainWidth As Single

Public OneTimeDAP_Exists As Integer

Public oTHERourlIST As String
Public InStatistic As Boolean

Public FirstPlantDay() As Integer
Public LastPlantDay() As Integer

Public FirstPlantDate() As Date
Public LastPlantDate() As Date

Public CloseISselected As Boolean

Public DisableExcelButton As Boolean

Public frmSplashShow As Boolean

Public WebPage As String

Public MaxAxisValue
Public MinAxisValue

Public EvaluateFileRead As Boolean

Public ExpExists As Boolean

Public NameSUM_File As String

Public CreateEvaluateListTable_DontDoThis As Boolean

Public NumberTempTables As Integer

Public NumberT_Tables As Integer
Public ArrayT_Tables() As String
Public LongFile As Boolean

Public gAllSelectedTr As Boolean
Public gAllSelectedRuns As Boolean

Declare Function FindFirstFile Lib "kernel32" Alias _
    "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData _
    As WIN32_FIND_DATA) As Long

Declare Function FindNextFile Lib "kernel32" Alias _
    "FindNextFileA" (ByVal hFindFile As Long, _
    lpFindFileData As WIN32_FIND_DATA) As Long

Declare Function GetFileAttributes Lib "kernel32" Alias _
    "GetFileAttributesA" (ByVal lpFileName As String) As Long

Declare Function FindClose Lib "kernel32" (ByVal hFindFile _
    As Long) As Long

Public Const MAX_PATH = 260
Public Const MAXDWORD = &HFFFF
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

Public Const LOCALE_SSHORTDATE = &H1F
Public Const LOCALE_SDECIMAL = &HE
'***************
Public Const SDecCurr = "."
Public Const SMieCurr = ","
Public Const DateFormat = "MM/dd/yyyy"
Public Const SeparatorData = "/"
Public Const Sdecimal = "."
Public Const SMie = ","

'*******************8


Public Declare Function GetSystemDefaultLCID _
        Lib "kernel32" () As Long

Public Declare Function SetLocaleInfo Lib _
     "kernel32" Alias "SetLocaleInfoA" ( _
     ByVal Locale As Long, _
     ByVal LCType As Long, _
     ByVal lpLCData As String) As Boolean

'''''''''''''''
Public AlCol As Boolean

Public DateFormat_Vechi
Public SDecimal_Vechi

'Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal _
Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal _
cchData As Long) As Long

Public Declare Function GetUserDefaultLCID% Lib "kernel32" ()


Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


Public Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
   (ByVal hwndCaller As Long, ByVal lpHelpFile As String, ByVal wCommand _
   As Long, ByVal dwData As Long) As Long
Public Const HH_DISPLAY_TOC = &H1
Public Const HH_DISPLAY_INDEX = &H2
Public Const HH_DISPLAY_SEARCH = &H3


'*****************************8
Public Function Get_locale()  ' Retrieve the regional setting

      Dim Symbol As String
      Dim iRet1 As Long
      Dim iRet2 As Long
      Dim lpLCDataVar As String
      Dim Pos As Integer
      
      Dim Locale

      Locale = GetUserDefaultLCID()
     
'*******************************************************************
      iRet1 = GetLocaleInfo(Locale, LOCALE_SSHORTDATE, lpLCDataVar, 0)
      Symbol = String$(iRet1, 0)
      iRet2 = GetLocaleInfo(Locale, LOCALE_SSHORTDATE, Symbol, iRet1)
      Pos = InStr(Symbol, Chr$(0))
      If Pos > 0 Then
           Symbol = Left$(Symbol, Pos - 1)
           'MsgBox "Regional Setting = " + Symbol
      End If
       DateFormat_Vechi = Symbol
      
'********************************************************************
    
    iRet1 = GetLocaleInfo(Locale, LOCALE_SDECIMAL, lpLCDataVar, 0)
    Symbol = String$(iRet1, 0)
    iRet2 = GetLocaleInfo(Locale, LOCALE_SDECIMAL, Symbol, iRet1)
    Pos = InStr(Symbol, Chr$(0))
      If Pos > 0 Then
           Symbol = Left$(Symbol, Pos - 1)
           'MsgBox "Regional Setting = " + Symbol
      End If
    SDecimal_Vechi = Symbol


End Function


Public Function Restore_Locale()
Dim iRet As Long
     
Dim Locale
Locale = GetUserDefaultLCID()
iRet = SetLocaleInfo(Locale, LOCALE_SSHORTDATE, DateFormat_Vechi)
iRet = SetLocaleInfo(Locale, LOCALE_SDECIMAL, SDecimal_Vechi)

End Function
'*******************************8



Public Function StripNulls(OriginalStr As String) As String

    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, _
        InStr(OriginalStr, Chr(0)) - 1)
    End If

    StripNulls = OriginalStr

End Function

Sub Main()
Dim PauseTime, Start
Dim counter As Long
Dim mm
Dim MyString As String
Dim lngLocale As Long
Dim AppCap As String
    
    AppCap = App.Title & " " & "v" & " " & App.Major & "." & App.Minor & "." & App.Revision
    
    ApplicationPathOption = IIf(Right$(App.path, 1) = "\", Mid$(App.path, 1, Len(App.path) - 1), App.path)
    frmOpenFileShown = False
    MyString = Command
    InStatistic = False
    CommandString = Split(Command, ",")
    ErrorExpFile = False
    frmSplashShow = False
    
    If App.PrevInstance = True Then
        MsgBox "Another instance of GBuild is already running!", vbCritical
        AppActivate AppCap
        End
    End If

    '*****************
    Get_locale
    '********************
    
    

    lngLocale = GetUserDefaultLCID()
    
    
 
    If SetLocaleInfo(lngLocale, _
        LOCALE_SSHORTDATE, _
        "MM/dd/yyyy") Then
    End If
   If SetLocaleInfo(lngLocale, _
        LOCALE_SDECIMAL, _
        ".") Then
    End If
    
    
    Open ApplicationPathOption & "\CommandLine.txt" For Output As #22
        Print #22, MyString
    Close 22
    FileStringCommand = ""
    NumberOfOUTfiles = UBound(CommandString) - 1
    For counter = 2 To UBound(CommandString)
        FileStringCommand = FileStringCommand & " " & Trim(CommandString(counter))
    Next
    If Trim(Command) <> "" And UBound(CommandString) < 2 Then
        MsgBox "Not enough parameters."
        Exit Sub
    End If
    If UBound(CommandString) = 2 Then
        FileStringCommand = "\" & Trim(FileStringCommand)
    ElseIf UBound(CommandString) > 2 Then
        FileStringCommand = "\ " & Trim(FileStringCommand)
    End If
    If Command = "" Then
        frmSplash.Show
        PauseTime = 1  ' Set duration.
        Start = Timer   ' Set start time.
        Do While Timer < Start + PauseTime
            DoEvents   ' Yield to other processes.
        Loop
        PaulsVersion = False
    ElseIf UBound(CommandString) = 1 Then
        PaulsVersion = False
        New_DataBasePath = Trim(CommandString(0))
        New_FilePath = Trim(CommandString(1))
    ElseIf UBound(CommandString) > 1 Then
        PaulsVersion = True
        New_DataBasePath = Trim(CommandString(0))
        New_FilePath = Trim(CommandString(1))
    ElseIf UBound(CommandString) = 0 Then
        PaulsVersion = False
        New_DataBasePath = Trim(CommandString(0))
        New_FilePath = ""
    Else
        PaulsVersion = False
        New_DataBasePath = ""
        New_FilePath = ""
    End If
    ObservedWasOpened = False
    
    Set fMainForm = New frmMain
    ExcelDataExist = False
    
    Load fMainForm
    Unload frmSplash

    fMainForm.Show
End Sub



Sub LoadResStrings(frm As Form)
    On Error Resume Next
    Dim ctl As Control
    Dim obj As Object
    Dim fnt As Object
    Dim sCtlType As String
    Dim nVal As Integer
    frm.Caption = LoadResString(CInt(frm.Tag))
    fnt.Name = LoadResString(20)
    'set the controls' captions using the caption
    'property for menu items and the Tag property
    'for all other controls
    For Each ctl In frm.Controls
        Set ctl.Font = fnt
        sCtlType = TypeName(ctl)
        If sCtlType = "Label" Then
            ctl.Caption = LoadResString(CInt(ctl.Tag))
        ElseIf sCtlType = "Menu" Then
            ctl.Caption = LoadResString(CInt(ctl.Caption))
        ElseIf sCtlType = "TabStrip" Then
            For Each obj In ctl.Tabs
                obj.Caption = LoadResString(CInt(obj.Tag))
                obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "Toolbar" Then
            For Each obj In ctl.Buttons
                obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "ListView" Then
            For Each obj In ctl.ColumnHeaders
                obj.Text = LoadResString(CInt(obj.Tag))
            Next
        Else
            nVal = 0
            nVal = Val(ctl.Tag)
            If nVal > 0 Then ctl.Caption = LoadResString(nVal)
            nVal = 0
            nVal = Val(ctl.ToolTipText)
            If nVal > 0 Then ctl.ToolTipText = LoadResString(nVal)
        End If
    Next


End Sub


Public Function TheDate(StringDate As String) As String
    Dim YearLastTwoNumbers As String
    Dim prmYear As String
    Dim prmDays As String
    'Paul
    If IsNumeric(Mid$(StringDate, 5)) = False Then
        Exit Function
    End If
Dim mm
On Error GoTo Error1
mm = Len(StringDate)

If Len(Trim(StringDate)) = 5 Then
    StringDate = "0" & Trim(StringDate)
End If

If Len(Trim(StringDate)) = 5 Then

    ' VSH
    'If Val(Mid$(Trim(StringDate), 1, Len(Trim(StringDate)) - 3)) > 15 Then
    If Val(Mid$(Trim(StringDate), 1, Len(Trim(StringDate)) - 3)) > 20 Then
        TheDate = Format(DateAdd("d", Right$(Trim(StringDate), 3) - 1, _
            "1/1/" & "19" & Mid$(Trim(StringDate), 1, 2)), _
            "mm/dd/yyyy")

    Else
        TheDate = Format(DateAdd("d", Right$(Trim(StringDate), 3) - 1, _
            "1/1/" & "20" & Right$(YearLastTwoNumbers, 2)), _
            "mm/dd/yyyy")
    End If
Else
    If InStr(1, StringDate, " ") <> 0 Then
        prmYear = Mid$(StringDate, 1, InStr(1, StringDate, " "))
        prmDays = Mid$(StringDate, InStr(1, StringDate, " ")) - 1
         
         TheDate = Format(DateAdd("d", prmDays, _
        "1/1/" & prmYear), "mm/dd/yyyy")

    End If
End If

    Exit Function
Error1:
MsgBox Err.Description & " in TheDate"
End Function


Sub SetDeviceIndependentWindow(TheForm As Form, prmType As Integer)
    'Adjust forms to screen
    Dim intFor As Integer
    On Error Resume Next
    If gsglXFactor = 1 And gsglYFactor = 1 Then Exit Sub
    
    If prmType = 0 Then
        If TheForm.WindowState <> 2 Then
            TheForm.Move TheForm.Left * gsglXFactor, _
            TheForm.Top * gsglYFactor, TheForm.Width * gsglXFactor, _
            TheForm.Height * gsglYFactor
        End If
    Else
        If TheForm.WindowState <> 2 Then
            TheForm.Height = TheForm.Height * gsglYFactor
            TheForm.Width = TheForm.Width * gsglXFactor
        End If
    End If
    
    For intFor = 0 To TheForm.Controls.Count - 1
        If TypeOf TheForm.Controls(intFor) Is Menu Then
        ElseIf TypeOf TheForm.Controls(intFor) Is CommonDialog Then
        ElseIf TypeOf TheForm.Controls(intFor) Is Image Then
        ElseIf TypeOf TheForm.Controls(intFor) Is PictureBox Then
        ElseIf TypeOf TheForm.Controls(intFor) Is Data Then
        ElseIf TypeOf TheForm.Controls(intFor) Is Chart Then
        ElseIf TypeOf TheForm.Controls(intFor) Is WebBrowser Then
            TheForm.Controls(intFor).Width = TheForm.Controls(intFor).Width * gsglXFactor
            TheForm.Controls(intFor).Height = TheForm.Controls(intFor).Height * gsglYFactor

        ElseIf TypeOf TheForm.Controls(intFor) Is ComboBox Then
            If TheForm.Controls(intFor).Font.Size * gsglXFactor > 11 Then
                TheForm.Controls(intFor).Font.Size = 11
            Else
                TheForm.Controls(intFor).Font.Size = _
                    TheForm.Controls(intFor).Font.Size * gsglXFactor
            End If
            TheForm.Controls(intFor).Weight = TheForm.Controls(intFor).Weight * gsglXFactor
        Else
            TheForm.Controls(intFor).Move TheForm.Controls(intFor).Left * _
                gsglXFactor, TheForm.Controls(intFor).Top * gsglYFactor, _
                TheForm.Controls(intFor).Width * gsglXFactor
        End If
            
        If TypeOf TheForm.Controls(intFor) Is TextBox Then
            If TheForm.Controls(intFor).FontSize * gsglXFactor > 11 Then
                TheForm.Controls(intFor).FontSize = 11
            Else
                TheForm.Controls(intFor).FontSize = _
                    TheForm.Controls(intFor).FontSize * gsglXFactor
            End If
            TheForm.Controls(intFor).Height = TheForm.Controls(intFor).Height * gsglYFactor
            TheForm.Controls(intFor).Weight = TheForm.Controls(intFor).Weight * gsglXFactor
        
        ElseIf TypeOf TheForm.Controls(intFor) Is Label Then
            If TheForm.Controls(intFor).Name = "lblTimeSeries" Or _
                TheForm.Controls(intFor).Name = "lblXYPlotting" Then
                If TheForm.Controls(intFor).FontSize * gsglXFactor > 14 Then
                    TheForm.Controls(intFor).FontSize = 14
                Else
                    TheForm.Controls(intFor).FontSize = _
                        TheForm.Controls(intFor).FontSize * gsglXFactor
                End If
            Else
                If TheForm.Controls(intFor).FontSize * gsglXFactor > 11 Then
                    TheForm.Controls(intFor).FontSize = 11
                Else
                    TheForm.Controls(intFor).FontSize = _
                        TheForm.Controls(intFor).FontSize * gsglXFactor
                End If
            End If
            TheForm.Controls(intFor).Height = TheForm.Controls(intFor).Height * gsglYFactor
            TheForm.Controls(intFor).Weight = TheForm.Controls(intFor).Weight * gsglXFactor
        ElseIf TypeOf TheForm.Controls(intFor) Is CommandButton Then
            If TheForm.Controls(intFor).FontSize * gsglXFactor > 11 Then
                TheForm.Controls(intFor).FontSize = 11
            Else
                TheForm.Controls(intFor).FontSize = _
                    TheForm.Controls(intFor).FontSize * gsglXFactor
            End If
            TheForm.Controls(intFor).Height = _
                TheForm.Controls(intFor).Height * gsglYFactor
            TheForm.Controls(intFor).Weight = _
                TheForm.Controls(intFor).Weight * gsglXFactor
        ElseIf TypeOf TheForm.Controls(intFor) Is OptionButton Then
            If TheForm.Controls(intFor).Font.Size * gsglXFactor > 11 Then
                TheForm.Controls(intFor).Font.Size = 11
            Else
                TheForm.Controls(intFor).Font.Size = _
                    TheForm.Controls(intFor).Font.Size * gsglXFactor
            End If
        ElseIf TypeOf TheForm.Controls(intFor) Is ListView Then
           TheForm.Controls(intFor).Font.Bold = False
            If TheForm.Controls(intFor).Font.Size * gsglXFactor > 11 Then
                TheForm.Controls(intFor).Font.Size = 11
            Else
                TheForm.Controls(intFor).Font.Size = _
                    TheForm.Controls(intFor).Font.Size * gsglXFactor
            End If
            TheForm.Controls(intFor).Height = TheForm.Controls(intFor).Height * gsglYFactor
        ElseIf TypeOf TheForm.Controls(intFor) Is Frame Then
            If TheForm.Controls(intFor).FontSize * gsglXFactor > 11 Then
                TheForm.Controls(intFor).FontSize = 11
            Else
                TheForm.Controls(intFor).FontSize = _
                    TheForm.Controls(intFor).FontSize * gsglXFactor
            End If
            TheForm.Controls(intFor).Height = _
                TheForm.Controls(intFor).Height * gsglYFactor
        ElseIf TypeOf TheForm.Controls(intFor) Is CheckBox Then
            If TheForm.Controls(intFor).FontSize * gsglXFactor > 11 Then
                TheForm.Controls(intFor).FontSize = 11
            Else
                TheForm.Controls(intFor).FontSize = _
                    TheForm.Controls(intFor).FontSize * gsglXFactor
            End If
        ElseIf TypeOf TheForm.Controls(intFor) Is DBGrid Then
            TheForm.Controls(intFor).Height = TheForm.Controls(intFor).Height * gsglYFactor
            TheForm.Controls(intFor).Weight = TheForm.Controls(intFor).Weight * gsglXFactor
            TheForm.Controls(intFor).HeadFont.Size = _
                TheForm.Controls(intFor).HeadFont.Size * gsglXFactor
            TheForm.Controls(intFor).Font.Size = _
                TheForm.Controls(intFor).Font.Size * gsglXFactor
        End If
        
    Next intFor



For intFor = 0 To TheForm.Controls.Count - 1
        
            'TheForm.Controls(intFor).Font = "Courier New"
            'TheForm.Controls(intFor).Font.Size = 12
       ' If TheForm.Controls(intFor).Font.Size > 7 Then
          '  TheForm.Controls(intFor).Font = "MS Sans Serif"
       ' Else
            TheForm.Controls(intFor).Font = "Microsoft Sans Serif"
        'End If
     Next intFor
'ElseIf TypeOf TheForm.Controls(intFor) Is ComboBox Then
    For intFor = 0 To TheForm.Controls.Count - 1
        If TypeOf TheForm.Controls(intFor) Is ListBox Then
           ' TheForm.Controls(intFor).Font = "Courier New"
            TheForm.Controls(intFor).Font.Size = 11
        
        End If
        If TypeOf TheForm.Controls(intFor) Is ComboBox Then
           ' TheForm.Controls(intFor).Font = "Courier New"
            TheForm.Controls(intFor).Font.Size = 11
        
        End If
        If TypeOf TheForm.Controls(intFor) Is RichTextBox Then
            TheForm.Controls(intFor).Font = "Courier New"
            TheForm.Controls(intFor).Font.Size = 11
        
        End If
        
        If TypeOf TheForm.Controls(intFor) Is TextBox Then
            TheForm.Controls(intFor).Font = "Courier New"
            TheForm.Controls(intFor).Font.Size = 11
        
        End If
        
         If TypeOf TheForm.Controls(intFor) Is DBGrid Then
            TheForm.Controls(intFor).Font = "Courier New"
           TheForm.Controls(intFor).HeadFont = "Microsoft Sans Serif"
            TheForm.Controls(intFor).Font.Size = 10
            TheForm.Controls(intFor).HeadFont.Size = 10
        End If
        'ListView
         If TypeOf TheForm.Controls(intFor) Is ListView Then
           ' TheForm.Controls(intFor).Font = "Courier New"
          ' TheForm.Controls(intFor).HeadFont = "Microsoft Sans Serif"
            TheForm.Controls(intFor).Font.Size = 11
           ' TheForm.Controls(intFor).HeadFont.Size = 11
        End If
        If TypeOf TheForm.Controls(intFor) Is MSChart Then
           ' TheForm.Controls(intFor).Font = "Courier New"
           TheForm.Controls(intFor).Title.Font = "Microsoft Sans Serif"
           ' TheForm.Controls(intFor).Font.Size = 11
            TheForm.Controls(intFor).Title.Font.Size = 24
        End If
        
     Next intFor



    Exit Sub
SDIWError: MsgBox Err.Description & " in SetDeviceIndependent..."
End Sub

Public Function RemoveSymbols(MyString As String)
Dim ChangedString As String
    'Change column names to remove unwanted symbols
    On Error GoTo Error1
     ChangedString = Replace(MyString, "X_Axis_", "")
     ChangedString = Replace(ChangedString, "777", "%")
     ChangedString = Replace(ChangedString, "555", "#")
     ChangedString = Replace(ChangedString, "999", "@")
     RemoveSymbols = Replace(ChangedString, "888", "$")
    
    Exit Function
Error1:
    MsgBox Err.Description & " in RemoveSymbols."
End Function

Public Function RemoveSymbols_Back(MyString As String)
Dim ChangedString As String
    On Error GoTo Error1
     ChangedString = Replace(MyString, "%", "777")
     ChangedString = Replace(ChangedString, "#", "555")
     ChangedString = Replace(ChangedString, "@", "999")
     RemoveSymbols_Back = Replace(ChangedString, "$", "888")
    
    Exit Function
Error1:
    MsgBox Err.Description & " in RemoveSymbols."
End Function


Public Function SelectedVariable(prmNewVariable As String) As String
Dim myrsDATA As Recordset
Dim StringCut As String
Dim myMultiplier As String
On Error GoTo Error1
    If prmNewVariable <> "" Then
        DoEvents
    Else
        Exit Function
    End If
    
    'Display proper lables
    prmNewVariable = Replace(prmNewVariable, "X_Axis_", "")
    If InStr(1, prmNewVariable, " x ") <> 0 Then
        myMultiplier = Mid$(prmNewVariable, InStr(prmNewVariable, " x "))
        prmNewVariable = Trim(Mid$(prmNewVariable, 1, InStr(prmNewVariable, " x ")))
    End If
    
    If InStr(prmNewVariable, "(") <> 0 Then
        Set myrsDATA = dbXbuild.OpenRecordset("Select * From [DATA] Where [CODE] = '" & _
            Mid$(prmNewVariable, 1, InStr(prmNewVariable, "(") - 1) & "'")
    Else
        Set myrsDATA = dbXbuild.OpenRecordset("Select * From [DATA] Where [CODE] = '" & _
            prmNewVariable & "'")
    End If
    With myrsDATA
        If Not (.EOF And .BOF) Then
            If InStr(prmNewVariable, "(") <> 0 Then
                StringCut = Mid$(prmNewVariable, InStr(prmNewVariable, "("))
            Else
                StringCut = ""
            End If
            SelectedVariable = !Label & StringCut
        Else
            SelectedVariable = prmNewVariable
        End If
    SelectedVariable = SelectedVariable & myMultiplier
    End With
    Err.Clear
    myrsDATA.Close
    Exit Function
Error1:
MsgBox Err.Description & " in SelectedVariable/frmGraph."
End Function


Public Sub myReadFile(myFileName As String)
Dim Readline As String
Dim k As Integer
Dim f
    On Error GoTo myErrorFile
    ReDim FileLoc(1)
    If Dir(myFileName) <> "" Then
        Open myFileName For Input As #1 Len = 80
        k = 0
        Do While Not EOF(1)    'loop till the end of the file
            Line Input #1, Readline
            'filter out the comments and the blank lines
                If k > 0 And k Mod (UBound(FileLoc, 1) + 1) = 0 Then
                    ReDim Preserve FileLoc(UBound(FileLoc, 1) + 1100)
                End If
                If FileType = "T-file" Then
                    If Mid$(Readline, 1, 1) <> "!" And Trim(Readline) <> "" And Mid$(Trim(Readline), 1, 1) <> "*" Then
                        FileLoc(k) = Readline
                    End If
                Else
                    If Mid$(Readline, 1, 1) <> "!" And Trim(Readline) <> "" Then
                        FileLoc(k) = Readline  'puts it into memory
                    End If
                End If
                k = k + 1
        Loop
        EvaluateFileRead = True
        Close 1
        NumberReadlines = k - 1
    Else
        If FileType = "SUM" Then
            EvaluateFileRead = False
        End If
    End If
    Exit Sub
myErrorFile:
    If InStr(UCase(myFileName), "DETAIL.CDE") <> 0 Then
        ErrorDetail = False
    Else
        MsgBox Err.Description & myFileName & " in myReadFile"
        If FileType = "SUM" Then
            EvaluateFileRead = False
        End If
        Unload frmWait2
        Unload frmWait3
        Screen.MousePointer = vbDefault
    End If
End Sub


Public Sub FillCreateOUTTable(OutputName1 As String, myOutput As String, NumberWait)
Dim i As Integer
Dim j As Integer
Dim MyExperimentID As String
Dim myTreatmentNumber As Integer
Dim myRunNumber As Integer
Dim myrsOUT As Recordset
Dim myrs1 As Recordset
    
    On Error GoTo Error1
    Set myrs1 = dbXbuild.OpenRecordset(OutputName1)
    frmWait2.prgLoad.Visible = True
    For i = 1 To NumberReadlines
        With myrs1
            
            frmWait2.prgLoad.Value = ((100 / NumberWait) / NumberReadlines) * (i - 1)
            If Trim(UCase(Mid$(FileLoc(i), 1, 4))) = "*RUN" Then
                .AddNew
                myRunNumber = Val(Mid$(FileLoc(i), 5, 5))
                !RunNumber = myRunNumber
                !RunDescription = Trim(Mid$(FileLoc(i), 19, 25))
                i = i + 1
                
            End If
            If Trim(UCase(Mid$(FileLoc(i), 1, 4))) = "MOD" Then
                !CultivarID = Trim(Mid$(FileLoc(i), 19, 8))
                
                i = i + 1
            End If
            If Trim(UCase(Mid$(FileLoc(i), 1, 4))) = "EXP" Then
                MyExperimentID = Trim(Mid$(FileLoc(i), 19, 8))
                !ExperimentID = MyExperimentID
                !CropID = Trim(Mid$(FileLoc(i), 28, 2))
                !ExpDescription = Trim(Mid$(FileLoc(i), 34, 25))
                i = i + 1
            End If
            If Trim(UCase(Mid$(FileLoc(i), 1, 4))) = "TRE" Then
                myTreatmentNumber = Val(Mid$(FileLoc(i), 12, 5))
                !TRNO = myTreatmentNumber
                !TreatmentDescription = Trim(Mid$(FileLoc(i), 19, 25))
                i = i + 1
                .Update
                
            End If
        
        End With
    Next i
    myrs1.Close
     
    
    Err.Clear
    On Error Resume Next
    myrsOUT.Close
    myrs1.Close
    Exit Sub
Error1:     MsgBox Err.Description & Err.Number & " in FillCreateOUTTable"
   ' GoTo 1
    Unload frmWait2
   Unload frmWait3
End Sub


Public Sub CreateOUTListTable(prmOutMyFile As String)
Dim VariableID() As String
Dim VariableDescription() As String
Dim i As Integer
Dim NumberOfVariables As Integer
Dim TableOUTname As String
Dim myrsTemp As Recordset
Dim myrsGrowth As Recordset
Dim myrsDATA As Recordset
Dim m As Integer
Dim k As Integer

    On Error Resume Next
    dbXbuild.Execute "Drop Table [OUTList]"
    TableOUTname = Mid$(prmOutMyFile, 1, InStr(prmOutMyFile, ".") - 1) & "_OUT"
    dbXbuild.Execute "Drop Table [" & Mid$(prmOutMyFile, 1, InStr(prmOutMyFile, ".") - 1) & "_List]"
    
    Err.Clear
    On Error GoTo Error1
    dbXbuild.Execute "Create Table [" & Mid$(prmOutMyFile, 1, InStr(prmOutMyFile, ".") - 1) & "_List" & "] " _
    & "([VariableID] Text (8), [VariableDescription] Text (250), [Selected] Text (10))"
    
    Set myrsGrowth = dbXbuild.OpenRecordset(TableOUTname)
    With myrsGrowth
        For i = 1 To .Fields.Count - 2 - StartVar
            ReDim Preserve VariableID(i)
            VariableID(i) = .Fields(i + 1 + StartVar).Name
        Next i
        NumberOfVariables = .Fields.Count - 2 - StartVar
    End With
    
    myrsGrowth.Close
    For i = 1 To NumberOfVariables
        Set myrsDATA = dbXbuild.OpenRecordset("Select * from [Data] Where [Code] = '" & VariableID(i) & "'")
        ReDim Preserve VariableDescription(i)
        Dim X As Integer
        If Not (myrsDATA.EOF And myrsDATA.BOF) Then
            k = k + 1
            VariableDescription(i) = myrsDATA.Fields("Description").Value
        Else
            VariableDescription(i) = RemoveSymbols(VariableID(i))
        End If
      
    Next i
    
    Set myrsTemp = dbXbuild.OpenRecordset(Mid$(prmOutMyFile, 1, InStr(prmOutMyFile, ".") - 1) & "_List")
    
    For i = 1 To NumberOfVariables
        If VariableID(i) <> "DAP" Then
            With myrsTemp
                .AddNew
            
                !VariableID = VariableID(i)
            
                !VariableDescription = VariableDescription(i)
                .Update
            End With
        End If
    Next i
    Err.Clear
    On Error Resume Next
    myrsTemp.Close
    myrsGrowth.Close
    myrsDATA.Close
Exit Sub
Error1: MsgBox Err.Description & " in CreatempTable."
    Unload frmWait2
    Unload frmWait3
    Screen.MousePointer = vbDefault
End Sub


Public Sub CreateExpTable(myExper As String)
Dim i As Integer
Dim LineLengh As Integer
Dim j As Integer
Dim k As Integer
Dim m As Integer
Dim VariableLineSpaces() As Integer
Dim MaxNumber As Integer
Dim VariablesName(1 To 9, 1 To 50)
Dim TableVariable()
Dim NumberOfTable As Integer
Dim MaxNumberOfTable As Integer
Dim NumberOfVariableInTable() As Integer
Dim NumberOfRecordsInTable() As Integer
Dim TheSigment
Dim prmMax As Integer
Dim MyDate() As String
Dim myTRNO() As Integer
Dim NumberMyDate As Integer
Dim NumberMyTRNO As Integer
Dim prmSubstitution As String
Dim n As Integer
Dim Adjusted As Integer
Dim MyVariable()
Dim NumberMyVariable As Integer
Dim myrs2 As Recordset
Dim myrsSelect As Recordset
Dim myrs As Recordset
Dim myrsFile As Recordset
Dim myrsDATE As Recordset
Dim myrsTRNO As Recordset
Dim mmBookmark2
Dim RecordFound As Boolean
Dim myString1 As String
Dim ErrorVariable As String

    VariablesName(1, 2) = "TRNO"
    VariablesName(1, 3) = "DATE"
If NumberReadlines = 0 Then Exit Sub
    ReDim TableVariable(1 To 9, 1 To 50, 1 To NumberReadlines)
    frmWait2.prgLoad.Value = 0
    For i = 1 To NumberReadlines
        If Trim(UCase(Mid$(FileLoc(i), 1, 4))) = "@TRN" Then
            ReDim LineSpaces(2)
            LineSpaces(1) = InStr(1, FileLoc(i), " ")
            NumberOfTable = NumberOfTable + 1
            LineLengh = Len(Trim(FileLoc(i)))
            For j = 1 To LineLengh
                ReDim Preserve LineSpaces(j + 1)
                LineSpaces(j + 1) = InStr(LineSpaces(j) + 1, Trim(FileLoc(i)), " ")
                If LineSpaces(j + 1) = 0 Then Exit For
            Next j
            MaxNumber = j + 1
            LineSpaces(MaxNumber) = Len(FileLoc(i)) + 1
            j = 0
            LineSpaces(0) = 1
            k = 1
            For j = 1 To MaxNumber
                    TheSigment = Trim(Mid$(FileLoc(i), LineSpaces(j - 1), LineSpaces(j) - LineSpaces(j - 1)))
                    ReDim Preserve VariableLineSpaces(0 To j + 1)
                    If TheSigment <> "" Then
                        'VariableLineSpaces(k) = LineSpaces(j - 1)
                        VariablesName(NumberOfTable, k) = TheSigment
                        k = k + 1
                  End If
            Next j
            
            prmMax = k
            
            Dim myValue As Long
            myValue = 14
            VariableLineSpaces(1) = 1
            VariableLineSpaces(2) = 7
            VariableLineSpaces(3) = 14
            For k = 4 To prmMax - 1
                VariableLineSpaces(k) = myValue + 6
                myValue = VariableLineSpaces(k)
            Next k
            
            
            
            VariableLineSpaces(prmMax) = LineSpaces(MaxNumber) + 1
            MaxNumber = k
            ReDim Preserve TrueVariableLineSpaces(k + 1)
            k = 1
            For j = 2 To MaxNumber
                    TrueVariableLineSpaces(j - 1) = VariableLineSpaces(j) - VariableLineSpaces(j - 1)
                TrueVariableLineSpaces(j - 1) = 6
            Next j
            
            For k = 3 To MaxNumber
                prmSubstitution = Replace(VariablesName(NumberOfTable, k), "#", "555")
                prmSubstitution = Replace(prmSubstitution, "%", "777")
                prmSubstitution = Replace(prmSubstitution, "@", "999")
                prmSubstitution = Replace(prmSubstitution, "$", "888")
                VariablesName(NumberOfTable, k) = prmSubstitution
                Err.Clear
                On Error GoTo Error1
                If VariablesName(NumberOfTable, k) <> "" Then
                    ErrorVariable = VariablesName(NumberOfTable, k)
                    dbXbuild.Execute "Alter Table [" & Replace(myExper, ".", " ") & "_Start] Add Column " & VariablesName(NumberOfTable, k) & " Single "
                    dbXbuild.Execute "Alter Table [" & Replace(myExper, ".", " ") & "] Add Column " & VariablesName(NumberOfTable, k) & " Single "
                End If
77:            Next k
            n = 1
            For k = i + 1 To NumberReadlines
                
                If Trim(UCase(Mid$(FileLoc(k), 1, 6))) = "" Then
                    Do While Trim(UCase(Mid$(FileLoc(k), 1, 6))) = "" Or Trim((Mid$(FileLoc(k), 1, 1))) = "*"
                        k = k + 1
                        If k >= NumberReadlines Then Exit For
                    Loop
                    If Trim(UCase(Mid$(FileLoc(k), 1, 4))) = "@TRN" Then
                        i = k - 1
                        Exit For
                    End If
                Else
                    If Trim(UCase(Mid$(FileLoc(k), 1, 4))) = "@TRN" Then
                        i = k - 1
                        Exit For
                    End If
                End If
                TableVariable(NumberOfTable, 1, n) = myExper
                TableVariable(NumberOfTable, 2, n) = Mid$(FileLoc(k), 1, TrueVariableLineSpaces(1))
                If Trim(TableVariable(NumberOfTable, 2, n)) = "" Then
                    Adjusted = 1
                    TableVariable(NumberOfTable, 2, n) = Mid$(FileLoc(k), 1, TrueVariableLineSpaces(1) + 1)
                Else
                    Adjusted = 0
                End If
                Dim bbb As String
                bbb = Trim(Mid$(FileLoc(k), VariableLineSpaces(2) - 1 + Adjusted, TrueVariableLineSpaces(2)))
                Dim NewDateData As String
                If InStr(bbb, " ") <> 0 Then
                    If Len(Trim(Mid$(FileLoc(k), VariableLineSpaces(2) + Adjusted, TrueVariableLineSpaces(2)))) = 3 Then
                        NewDateData = "0" & Trim(Mid$(FileLoc(k), VariableLineSpaces(2) + Adjusted, TrueVariableLineSpaces(2)))
                    Else
                        NewDateData = Trim(Mid$(FileLoc(k), VariableLineSpaces(2) + Adjusted, TrueVariableLineSpaces(2)))
                    End If
                    TableVariable(NumberOfTable, 3, n) = NewDateData
                Else
                    If Len(Trim(Mid$(FileLoc(k), VariableLineSpaces(2) - 1 + Adjusted, TrueVariableLineSpaces(2)))) = 3 Then
                        NewDateData = "0" & Trim(Mid$(FileLoc(k), VariableLineSpaces(2) - 1 + Adjusted, TrueVariableLineSpaces(2)))
                    Else
                        NewDateData = Trim(Mid$(FileLoc(k), VariableLineSpaces(2) - 1 + Adjusted, TrueVariableLineSpaces(2)))
                    End If
                    TableVariable(NumberOfTable, 3, n) = NewDateData
                End If
                
                For m = 3 To MaxNumber - 1
                    TableVariable(NumberOfTable, m + 1, n) = IIf(Len(Trim(Mid$(FileLoc(k), _
                        VariableLineSpaces(m) - 1 + Adjusted, TrueVariableLineSpaces(m)))) = 0, _
                        "-99", Mid$(FileLoc(k), VariableLineSpaces(m) - 1 + Adjusted, TrueVariableLineSpaces(m)))
                Next m
                    
                    '****Number fields in a section
                ReDim Preserve NumberOfVariableInTable(NumberOfTable)
                NumberOfVariableInTable(NumberOfTable) = m
                i = k
                n = n + 1
                Adjusted = 0
            Next k
            
        End If
    
    ReDim Preserve NumberOfRecordsInTable(NumberOfTable + 1)
    NumberOfRecordsInTable(NumberOfTable) = n - 2
    Next i
    frmWait2.Show
    DoEvents
    frmWait2.prgLoad.Value = 10
    
    '****Number of different sections
    MaxNumberOfTable = NumberOfTable
    Err.Clear
    On Error GoTo Error2
    
    Set myrs = dbXbuild.OpenRecordset(Replace(myExper, ".", " ") & "_Start")
    With myrs
        For i = 1 To MaxNumberOfTable
            frmWait2.prgLoad.Visible = True
           
            For m = 1 To NumberOfRecordsInTable(i) + 1
                frmWait2.prgLoad.Value = (100 / MaxNumberOfTable) * (i - 1) + ((100 / MaxNumberOfTable) / NumberOfRecordsInTable(i)) * (m - 1)
                .AddNew
                .Fields("TRNO").Value = Val(TableVariable(i, 2, m))
                If Len(Trim((TableVariable(i, 3, m)))) = 4 Then
                    TableVariable(i, 3, m) = "0" & Trim(TableVariable(i, 3, m))
                End If
                If Len(Trim((TableVariable(i, 3, m)))) = 5 Then
                ' VSH
                    'If Val(Mid$(Trim((TableVariable(i, 3, m))), 1, 2)) > 15 Then
                    If Val(Mid$(Trim((TableVariable(i, 3, m))), 1, 2)) > 20 Then
                        TableVariable(i, 3, m) = "19" & _
                            Mid$(Trim((TableVariable(i, 3, m))), 1, 2) & _
                            " " & Mid$(Trim((TableVariable(i, 3, m))), 3)
                    Else
                        TableVariable(i, 3, m) = "20" & _
                            Mid$(Trim((TableVariable(i, 3, m))), 1, 2) & _
                            " " & Mid$(Trim((TableVariable(i, 3, m))), 3, 3)
                    End If
                ElseIf Len(Trim((TableVariable(i, 3, m)))) < 5 Then
                    TableVariable(i, 3, m) = Null
                End If
                .Fields("Section").Value = i
                .Fields("TheOrder") = TableVariable(i, 3, m)
                .Fields("Date") = Val(Mid$(TableVariable(i, 3, m), 6))
                For k = 3 To NumberOfVariableInTable(i) - 1
                    If IsNumeric(TableVariable(i, k + 1, m)) = False Then
                        TableVariable(i, k + 1, m) = -99
                    End If
                    .Fields(VariablesName(i, k)).Value = IIf(Val(TableVariable(i, k + 1, m)) = -99, Null, Val(TableVariable(i, k + 1, m)))
                Next k
                .Update
            Next m
        Next i
        frmWait2.prgLoad.Value = 100
        ReDim MyVariable(1 To .Fields.Count - 4)
        NumberMyVariable = .Fields.Count - 4
        For i = 1 To NumberMyVariable
            MyVariable(i) = .Fields(i + 3).Name
        Next i
        .Close
    End With
    
    myString1 = ""
    For i = 1 To NumberMyVariable
        myString1 = myString1 & ", " & " [" & MyVariable(i) & "]"
    Next i
    dbXbuild.Execute "Insert Into [" & Replace(myExper, ".", " ") & "] " & _
                "Select * From [" & Replace(myExper, ".", " ") & "_Start] Where [Section] = 1 Order By [TRNO], [Date]" '& myString1 & ";"
    Set myrsFile = dbXbuild.OpenRecordset("Select * From [" & Replace(myExper, ".", " ") & "] ")
    With myrsFile
        If Not (.EOF And .BOF) Then
            .MoveFirst
            For i = 2 To MaxNumberOfTable
                frmWait2.prgLoad.Visible = True
                'frmWait2.prgLoad.Value = 100 / MaxNumberOfTable * (i - 1)
                frmWait3.prgLoad.Visible = True
               ' frmWait3.prgLoad.Value = 100 / MaxNumberOfTable * (i - 1)
                Set myrsDATE = dbXbuild.OpenRecordset("Select * From [" & Replace(myExper, ".", " ") & "_Start] Where " _
                & "[Section] = " & i)
                Do While Not myrsDATE.EOF
                    frmWait2.prgLoad.Value = (100 / MaxNumberOfTable) * (i - 1) + ((100 / MaxNumberOfTable) * (myrsDATE.PercentPosition / 100))
                    frmWait3.prgLoad.Value = (100 / MaxNumberOfTable) * (i - 1) + ((100 / MaxNumberOfTable) * (myrsDATE.PercentPosition / 100))

                    .FindFirst ("TheOrder = '" & myrsDATE.Fields("TheOrder").Value & "'")
                    If .NoMatch Then
                        Err.Clear
                        On Error GoTo Error3
                        .AddNew
                        .Fields("TRNO").Value = myrsDATE.Fields("TRNO").Value
                        .Fields("Date").Value = Val(Mid$(myrsDATE.Fields("TheOrder").Value, 6))
                        
                        .Fields("TheOrder") = myrsDATE.Fields("TheOrder").Value
                        For k = 1 To NumberMyVariable
                            .Fields(MyVariable(k)).Value = myrsDATE.Fields(MyVariable(k)).Value
                        Next k
                        .Update
                    Else
                        .MoveFirst
                            Err.Clear
                            On Error GoTo Error4
                        RecordFound = False
                        For k = 1 To NumberMyVariable
                            .MoveFirst
                            Do While Not .EOF
                                If .Fields("TheOrder").Value = myrsDATE.Fields("TheOrder").Value And _
                                    .Fields("TRNO").Value = myrsDATE.Fields("TRNO").Value Then
                                    RecordFound = True
                                    mmBookmark2 = .Bookmark
                                    Exit Do
                                End If
                                .MoveNext
                            Loop
                            If IsNull(myrsDATE.Fields(MyVariable(k)).Value) = False Then
                                Err.Clear
                                On Error GoTo Error5
                                
                                If RecordFound = True Then
                                    .Bookmark = mmBookmark2
                                    If IsNull(.Fields(MyVariable(k)).Value) = True Then
                                        .Edit
                                        .Fields(MyVariable(k)).Value = myrsDATE.Fields(MyVariable(k)).Value
                                        .Update
                                    Else
                                        .AddNew
                                        .Fields("Date").Value = Val(Mid$(myrsDATE.Fields("TheOrder").Value, 6))
                                        .Fields("TheOrder") = myrsDATE.Fields("TheOrder").Value
                                        .Fields("TRNO").Value = myrsDATE.Fields("TRNO").Value
                                        .Fields(MyVariable(k)).Value = myrsDATE.Fields(MyVariable(k)).Value
                                        .Update
                                    End If
                                Else
                                    .AddNew
                                    .Fields("Date").Value = Val(Mid$(myrsDATE.Fields("TheOrder").Value, 6))
                                    .Fields("TheOrder") = myrsDATE.Fields("TheOrder").Value
                                    .Fields("TRNO").Value = myrsDATE.Fields("TRNO").Value
                                    .Fields(MyVariable(k)).Value = myrsDATE.Fields(MyVariable(k)).Value
                                    .Update
                                End If
                            End If
                        Next k
                    End If
                    myrsDATE.MoveNext
                Loop
            Next i
            frmWait2.prgLoad.Visible = False
            frmWait3.prgLoad.Value = 100
        End If
    End With
    Err.Clear
    On Error Resume Next
    myrs2.Close
    myrsSelect.Close
    myrs.Close
    myrsFile.Close
    myrsDATE.Close
    myrsTRNO.Close
    
    Err.Clear
    On Error GoTo Error2
    dbXbuild.Execute "Drop TABLE [" & Replace(myExper, ".", " ") & "_Start]"
    dbXbuild.Execute "Alter Table [" & Replace(myExper, ".", " ") & "] Drop Column [Section]"
    
    Dim Field_Remove() As String
    Dim MyNumberOfRecords As Integer
    Dim kk As Integer
    
    Set myrs = dbXbuild.OpenRecordset(Replace(myExper, ".", " "))
    If Not (myrs.EOF And myrs.BOF) Then
        myrs.MoveLast
        MyNumberOfRecords = myrs.RecordCount
    End If
    kk = 1
    For i = 1 To NumberMyVariable
        Set myrs = dbXbuild.OpenRecordset("Select * From [" & Replace(myExper, ".", " ") & _
        "] Where [" & MyVariable(i) & "] Is Null")
        If Not (myrs.EOF And myrs.BOF) Then
            myrs.MoveLast
            If MyNumberOfRecords = myrs.RecordCount Then
                ReDim Preserve Field_Remove(kk)
                Field_Remove(kk) = MyVariable(i)
                kk = kk + 1
            End If
        End If
    Next i

    myrs.Close
    If kk > 1 Then
        For i = 1 To kk - 1
            dbXbuild.Execute "Alter Table [" & Replace(myExper, ".", " ") & "] Drop Column [" _
            & Field_Remove(i) & "]"
        Next i
    End If
    Err.Clear
    On Error Resume Next
    Err.Clear
    On Error GoTo Error2
   
    Exit Sub
Error1:
    Unload frmWait2
  '  MsgBox Err.Description
    Unload frmWait3
    'MsgBox Err.Description
    Screen.MousePointer = vbDefault
    ErrorExpFile = True
    If Err.Number = 3380 Then
        MsgBox "Error in file" & " " & myExper & ". " & "Variable" & " " & ErrorVariable & " " & "is repeated." & _
        Chr(10) & Chr(13) & " " & "The program will read only the data from the first column.", vbCritical
        ErrorExpFile = False
        GoTo 77
    Else
    'MsgBox Err.Description
    End If
    Exit Sub
Error2:
    MsgBox Err.Description
        'MsgBox "Error in experimental file:" & " " & Err.Description & " " & myExper
       ' MsgBox Err.Description & " 2 in CreateExpTable."
    Exit Sub
Error3:
    
    Unload frmWait2
    Unload frmWait3
    Screen.MousePointer = vbDefault
    ErrorWithExp = True
    Exit Sub
Error4:
    'MsgBox "Error in experimental file" & " " & myExper
   ' MsgBox Err.Description
    ErrorWithExp = True
    Unload frmWait2
    Unload frmWait3
    Screen.MousePointer = vbDefault
    Exit Sub
    

Error5:
    'MsgBox "Error in experimental file" & " " & myExper
    Unload frmWait2
    Unload frmWait3
    Screen.MousePointer = vbDefault
    ErrorWithExp = True
End Sub



Public Sub CreateExpList(prmExpFile As String)
Dim VariableDescription() As String
Dim i As Integer
Dim NumberOfVariables As Integer
Dim TableExp As String
Dim myrsTemp As Recordset
Dim myrsExp As Recordset
Dim myrsDATA As Recordset
Dim m As Integer
Dim VariableExpID() As String
Dim k As Integer
Dim X As Integer

    prmExpFile = Replace(prmExpFile, ".", " ")
    On Error Resume Next
    dbXbuild.Execute "Drop Table [" & prmExpFile & "_List] "
    
    Err.Clear
    On Error GoTo Error1
    dbXbuild.Execute "Create Table [" & prmExpFile & "_List] ([VariableID] Text (8), [VariableDescription] Text (250), [Selected] Text (10))"
    Set myrsExp = dbXbuild.OpenRecordset(prmExpFile)
    With myrsExp
        For i = 4 To .Fields.Count
            ReDim Preserve VariableExpID(i)
            VariableExpID(i - 3) = .Fields(i - 1).Name
        Next i
        NumberOfVariables = .Fields.Count - 2
        .Close
    End With
    
    For i = 1 To NumberOfVariables - 1
        Set myrsDATA = dbXbuild.OpenRecordset("Select * From [Data] Where  " & _
        "[Code] = '" & VariableExpID(i) & "'")
        ReDim Preserve VariableDescription(i)
        
        If Not (myrsDATA.EOF And myrsDATA.BOF) Then
            k = k + 1
            VariableDescription(i) = myrsDATA.Fields("Description").Value
        Else
            'Remove
            VariableDescription(i) = RemoveSymbols(VariableExpID(i))
        End If
      '  If Len(VariableDescription(i)) > 50 Then VariableDescription(i) = Left$(VariableDescription(i), 50)
    Next i
    Set myrsTemp = dbXbuild.OpenRecordset(prmExpFile & "_List")
    For i = 1 To NumberOfVariables - 1
        With myrsTemp
            .AddNew
            !VariableID = VariableExpID(i)
            !VariableDescription = VariableDescription(i)
            .Update
        End With
    Next i
    Err.Clear
    On Error Resume Next
    myrsTemp.Close
    myrsExp.Close
    myrsDATA.Close
    Exit Sub
Error1: MsgBox Err.Description & " in CreateExpList."
    Unload frmWait2
    Unload frmWait3
    Screen.MousePointer = vbDefault
End Sub


Public Sub CreateTempTreatmentTable()
Dim myrs As Recordset
Dim k As Integer
Dim i As Integer
   
    On Error GoTo Error1
     dbXbuild.Execute "Create Table [" & _
           "Treatment_Names] ([TRNO] Integer, [Description] Text)"
    Set myrs = dbXbuild.OpenRecordset("Treatment_Names")
    With myrs
        For k = 1 To NumberReadlines
            If UCase(Mid$(FileLoc(k), 1, 11)) = "*TREATMENTS" Then
                For i = k + 2 To NumberReadlines
                    If Mid$(FileLoc(i), 1, 1) = "*" Or Trim(Mid$(FileLoc(i), 1, 10)) = "" Then
                        k = NumberReadlines
                        Exit For
                    End If
                    .AddNew
                    !TRNO = Val(Mid$(FileLoc(i), 1, 3))
                    !Description = Mid$(FileLoc(i), 10, 25)
                    .Update
                Next i
            End If
        Next k
        .Close
    End With
    Exit Sub

Error1:     MsgBox Err.Description & "in " & "CreateTempTreatmentTable"
End Sub


Public Sub CreateOutDataTable(MyOUT As String)
Dim LineLengh As Integer
Dim j As Integer
Dim k As Integer
Dim m As Integer
Dim VariableLineSpaces() As Integer
Dim TheSigment As String
Dim i As Integer
Dim prmSubstituteSymbols As String
Dim kk As Integer
Dim mm As Integer
Dim TheSecondKK As Integer
Dim ff
Dim iii As Double
Dim vvv As Integer
Dim NumberOut As Integer
'Dim prmExiFor As Boolean
'****
Dim MyfirstLineStart As Integer
Dim bb As Integer
NumberOut = 1

For iii = 1 To NumberReadlines
    On Error Resume Next
    dbXbuild.Execute "Drop Table [" & MyOUT & "_OUT" & NumberOut & "]"
    Err.Clear
    On Error GoTo Error1
    dbXbuild.Execute "Create Table [" & MyOUT & "_OUT" & NumberOut & "] ([RunNumber] Integer, [DATE] Text (10))"
    ReDim VariablesOUTName(5)
    
     For i = iii To NumberReadlines
    
        If Trim(UCase((Replace(Mid$(FileLoc(i), 1, 6), " ", "")))) = "@DATE" Then
           For bb = 1 To i
                If Trim(Mid$(FileLoc(i - bb), 1, 4)) = "*RUN" Then
                    MyfirstLineStart = i - bb
                    Exit For
                End If
            Next bb
         
            
            LineLengh = Len(Trim(FileLoc(i)))
            j = 1
            For kk = 6 To LineLengh
                If InStr(Mid$(FileLoc(i), kk, 1), " ") <> 0 Then
                    DoEvents
                Else
                    ReDim LineSpaces(1)
                    LineSpaces(1) = kk - 1
                    Exit For
                End If
            Next kk
            j = 2
            TheSecondKK = kk
                
            For kk = TheSecondKK To LineLengh
                If InStr(Mid$(FileLoc(i), kk, 1), " ") = 0 Then
                    DoEvents
                Else
                    ReDim Preserve LineSpaces(j + 1)
                    LineSpaces(j) = kk
                    mm = kk + 1
                    For mm = kk To LineLengh
                        If InStr(Mid$(FileLoc(i), mm, 1), " ") <> 0 Then
                            DoEvents
                        Else
                            Exit For
                        End If
                    Next mm
                    kk = mm
                    j = j + 1
                End If
            Next kk
            LineSpaces(j) = LineLengh + 1
            NumberOUTVariables = j
            j = 1
            
            ReDim Preserve VariablesOUTName(NumberOUTVariables + 1)
            
            
            
            If UCase(Trim(Mid$(FileLoc(i), LineSpaces(1), LineSpaces(2) - LineSpaces(1)))) = "CDAY" Then
                CDAYexists = True
                dbXbuild.Execute "Alter Table [" & MyOUT & "_OUT" & NumberOut & "] Add Column [CDAY] Integer"
                StartVar = 1
            Else
                CDAYexists = False
                If UCase(Trim(Mid$(FileLoc(i), LineSpaces(2), LineSpaces(3) - LineSpaces(2)))) = "DAP" Then
                    dbXbuild.Execute "Alter Table [" & MyOUT & "_OUT" & NumberOut & "]  Add Column [DAS] Integer"
                    dbXbuild.Execute "Alter Table [" & MyOUT & "_OUT" & NumberOut & "]  Add Column [DAP] Integer"
                    DAPexists = True
                    OneTimeDAP_Exists = 1
                    StartVar = 2
                ElseIf UCase(Trim(Mid$(FileLoc(i), LineSpaces(1), LineSpaces(2) - LineSpaces(1)))) = "DAS" Then
                    dbXbuild.Execute "Alter Table [" & MyOUT & "_OUT" & NumberOut & "]  Add Column [DAS] Integer"
                        DAPexists = False
                    StartVar = 1
                End If
            End If
            
            For j = 1 To NumberOUTVariables - StartVar - 1
               VariablesOUTName(j) = Trim(Mid$(FileLoc(i), LineSpaces(j + StartVar), _
                LineSpaces(j + StartVar + 1) - LineSpaces(j + StartVar)))
            Next j
            NumberOUTVariables = NumberOUTVariables - StartVar - 1
            For j = 1 To NumberOUTVariables
                prmSubstituteSymbols = RemoveSymbols_Back(VariablesOUTName(j))
                VariablesOUTName(j) = prmSubstituteSymbols
            Next j
            Err.Clear
            For k = 1 To NumberOUTVariables
                dbXbuild.Execute "Alter Table [" & MyOUT & "_OUT" & NumberOut & "]  Add Column [" & VariablesOUTName(k) & "] Single "
            Next k
            Exit For 'Only one time reading variables
        End If
        
        
        
        If Trim(UCase((Replace(Mid$(FileLoc(i), 1, 6), " ", "")))) = "@YEAR" Then
            For bb = 1 To i
                If Trim(Mid$(FileLoc(i - bb), 1, 4)) = "*RUN" Then
                    MyfirstLineStart = i - bb
                    Exit For
                End If
            Next bb
            
            LineLengh = Len(Trim(FileLoc(i)))
            j = 1
            For kk = 7 To LineLengh
                If InStr(Mid$(FileLoc(i), kk, 1), " ") <> 0 Then
                    DoEvents
                Else
                    ReDim LineSpaces(1)
                    LineSpaces(0) = kk - 1
                    Exit For
                End If
            Next kk
            j = 1
            TheSecondKK = kk
                
            For kk = TheSecondKK To LineLengh
                If InStr(Mid$(FileLoc(i), kk, 1), " ") = 0 Then
                    DoEvents
                Else
                    ReDim Preserve LineSpaces(j + 1)
                    LineSpaces(j) = kk
                    mm = kk + 1
                    For mm = kk To LineLengh
                        If InStr(Mid$(FileLoc(i), mm, 1), " ") <> 0 Then
                            DoEvents
                        Else
                            Exit For
                        End If
                    Next mm
                    kk = mm
                    j = j + 1
                End If
            Next kk
            LineSpaces(j) = LineLengh + 1
            NumberOUTVariables = j
            j = 1
            ReDim Preserve VariablesOUTName(NumberOUTVariables + 1)
            ff = UCase(Trim(Mid$(FileLoc(i), LineSpaces(1), LineSpaces(2) - LineSpaces(1))))
            If UCase(Trim(Mid$(FileLoc(i), LineSpaces(1), LineSpaces(2) - LineSpaces(1)))) = "CDAY" Then
                CDAYexists = True
                dbXbuild.Execute "Alter Table [" & MyOUT & "_OUT" & NumberOut & "]  Add Column [CDAY] Integer"
                StartVar = 1
            Else
                CDAYexists = False
                If UCase(Trim(Mid$(FileLoc(i), LineSpaces(2), LineSpaces(3) - LineSpaces(2)))) = "DAP" Then
                    dbXbuild.Execute "Alter Table [" & MyOUT & "_OUT" & NumberOut & "]  Add Column [DAS] Integer"
                    dbXbuild.Execute "Alter Table [" & MyOUT & "_OUT" & NumberOut & "]  Add Column [DAP] Integer"
                    DAPexists = True
                    OneTimeDAP_Exists = 1
                    StartVar = 2
                ElseIf UCase(Trim(Mid$(FileLoc(i), LineSpaces(1), LineSpaces(2) - LineSpaces(1)))) = "DAS" Then
                    dbXbuild.Execute "Alter Table [" & MyOUT & "_OUT" & NumberOut & "]  Add Column [DAS] Integer"
                    Err.Clear
                    On Error Resume Next
                    dbXbuild.Execute "Alter Table [" & MyOUT & "_OUT" & NumberOut & "]  Add Column [DAP] Integer"
                    Err.Clear
                    On Error GoTo Error1
                    DAPexists = False
                    StartVar = 1
                
                End If
            End If
            ''''''
            For j = 1 To NumberOUTVariables - StartVar - 1
                VariablesOUTName(j) = Trim(Mid$(FileLoc(i), LineSpaces(j + StartVar), _
                LineSpaces(j + StartVar + 1) - LineSpaces(j + StartVar)))
            Next j
            
            NumberOUTVariables = NumberOUTVariables - StartVar - 1
                 
            For j = 1 To NumberOUTVariables
                prmSubstituteSymbols = RemoveSymbols_Back(VariablesOUTName(j))
                VariablesOUTName(j) = prmSubstituteSymbols
            Next j
                
            For k = 1 To NumberOUTVariables
                If VariablesOUTName(k) <> "" Then
                    dbXbuild.Execute "Alter Table [" & MyOUT & "_OUT" & NumberOut & "]  Add Column [" & VariablesOUTName(k) & "] Single "
                End If
            Next k
            Exit For 'Only one time reading variables
        End If
    
    Next i
    Dim TheLastLine As Boolean
    TheLastLine = False
    For vvv = 1 To NumberReadlines - i
        If Mid$(FileLoc(i + vvv), 1, 2) = "*R" Then
            TheLastLine = True
            Exit For
        Else
            DoEvents
        End If
    Next vvv
    i = i + vvv
    If TheLastLine = True Then
       ' Call FillGROWTH(MyOUT & "_OUT" & NumberOut, MyfirstLineStart, i - 1)
        Call FillGROWTH(MyOUT & "_OUT" & NumberOut, MyfirstLineStart, i)
    Else
        Call FillGROWTH(MyOUT & "_OUT" & NumberOut, MyfirstLineStart, i + 1)
    End If
   ' Call FillGROWTH(MyOUT & "_OUT" & NumberOut, MyfirstLineStart, i + 1)

'****
    iii = i
    
    NumberOut = NumberOut + 1
    
    
    frmWait2.prgLoad.Value = IIf(Round(iii * 100 / NumberReadlines) > 100, 100, Round(iii * 100 / NumberReadlines))

Next iii
    NumberTempTables = NumberOut - 1
    Exit Sub
Error1: MsgBox Err.Description & iii & " in CreateOUTDataTable."
    Unload frmWait2
    Unload frmWait3
    Screen.MousePointer = vbDefault
End Sub


Public Function RemoveSpaces(MyString As String) As String
Dim myTrim As String
    On Error GoTo Error1
    myTrim = Replace(Trim(MyString), " ", "")
    myTrim = Replace(Trim(myTrim), "\", "")
    RemoveSpaces = myTrim
    Exit Function
Error1:     MsgBox Err.Description
End Function


Public Sub CreateEvaluateTable()
Dim LineLengh As Integer
Dim j As Integer
Dim k As Integer
Dim i As Integer
Dim prmSubstituteSymbols As String
Dim kk As Integer
Dim mm As Integer
Dim TheSecondKK As Integer
    On Error Resume Next
    dbXbuild.Execute "Drop Table [Evaluate_SUM]"
    Err.Clear
    On Error GoTo Error1
    dbXbuild.Execute "Create Table [Evaluate_SUM] ([A_remove] Integer, [B_remove] Integer )"
    
    For i = 1 To NumberReadlines
        If Mid$(Trim(FileLoc(i)), 1, 1) = "@" Then
            LineLengh = Len(Trim(FileLoc(i)))
            j = 1
            For kk = 2 To LineLengh
                If InStr(Mid$(FileLoc(i), kk, 1), " ") <> 0 Then
                    DoEvents
                Else
                    ReDim LineSpaces(1)
                    LineSpaces(1) = kk - 1
                    Exit For
                End If
            Next kk
            j = 2
            TheSecondKK = kk
            For kk = TheSecondKK To LineLengh
                If InStr(Mid$(FileLoc(i), kk, 1), " ") = 0 Then
                    DoEvents
                Else
                    ReDim Preserve LineSpaces(j + 1)
                    LineSpaces(j) = kk
                    mm = kk + 1
                    For mm = kk To LineLengh
                        If InStr(Mid$(FileLoc(i), mm, 1), " ") <> 0 Then
                            DoEvents
                        Else
                            Exit For
                        End If
                    Next mm
                    kk = mm
                    j = j + 1
                End If
            Next kk
            LineSpaces(j) = LineLengh + 1
            NumberOUTVariables = j
            j = 1
            ReDim VariablesOUTName(2)
            ReDim Preserve VariablesOUTName(NumberOUTVariables + 1)

            For j = 1 To NumberOUTVariables - 2
               VariablesOUTName(j) = Trim(Mid$(FileLoc(i), LineSpaces(j), _
                LineSpaces(j + 1) - LineSpaces(j)))
            Next j
            VariablesOUTName(NumberOUTVariables - 1) = Trim(Mid$(FileLoc(i), LineSpaces(j)))
            
            NumberOUTVariables = NumberOUTVariables - 1
            For j = 1 To NumberOUTVariables
                prmSubstituteSymbols = RemoveSymbols_Back(VariablesOUTName(j))
                VariablesOUTName(j) = prmSubstituteSymbols
            Next j
            Err.Clear
            For k = 1 To NumberOUTVariables
                dbXbuild.Execute "Alter Table [Evaluate_SUM] Add Column [" & VariablesOUTName(k) & "] Single "
            Next k
                dbXbuild.Execute "Alter Table [Evaluate_SUM] Drop Column [A_remove]"
                dbXbuild.Execute "Alter Table [Evaluate_SUM] Drop Column [B_remove]"
            i = i + 1
            Do While Trim(FileLoc(i)) = ""
                If i > NumberReadlines Then
                    Exit Do
                End If
                i = i + 1
             Loop
                FillEvaluate (i)
            Exit For 'Only one time reading variables
        End If
    Next i
    Exit Sub
Error1: MsgBox Err.Description & " in CreateEvaluateTable."
    Unload frmWait2
    Unload frmWait3
    Screen.MousePointer = vbDefault
End Sub


Public Sub CreateEvaluateListTable()
Dim VariableID() As String
Dim VariableDescription() As String
Dim i As Integer
Dim NumberOfVariables As Integer
Dim TableOUTname As String
Dim myrsTemp As Recordset
Dim myrsGrowth As Recordset
Dim myrsDATA As Recordset
Dim m As Integer
Dim k As Integer
Dim EvaluateNULL_Names() As String
Dim EvalRecordNumber As Integer
Dim myFieldName() As String
Dim Null_Variables() As String
Dim NumberNullVariables As Integer

    On Error Resume Next
    dbXbuild.Execute "Drop Table [Evaluate_List]"
    dbXbuild.Execute "Drop Table [TempEvaluateList]"
    Err.Clear
    On Error GoTo Error1
        dbXbuild.Execute "Create Table [TempEvaluateList] " _
    & "([VariableID] Text (8))"
    dbXbuild.Execute "Create Table [Evaluate_List] " _
    & "([VariableID] Text (8), [VariableDescription] Text (250), [Selected] Text (10))"
    
    ReDim Null_Variables(1)
    Set myrsGrowth = dbXbuild.OpenRecordset("Evaluate_SUM")
    With myrsGrowth
        NumberOfVariables = .Fields.Count
        ReDim myFieldName(NumberOfVariables + 1)
        If Not (.EOF And .BOF) Then
            .MoveLast
            EvalRecordNumber = .RecordCount
            For i = 0 To NumberOfVariables - 1
                myFieldName(i) = .Fields(i).Name
            Next i
        End If
        .Close
    End With
    k = 0
    For i = 0 To NumberOfVariables - 1
        Set myrsGrowth = dbXbuild.OpenRecordset("Select * From [Evaluate_SUM] Where [" & _
            myFieldName(i) & "] is null")
        With myrsGrowth
            If Not (.EOF And .BOF) Then
                .MoveLast
                If .RecordCount = EvalRecordNumber Then
                    k = k + 1
                    ReDim Preserve Null_Variables(k)
                    Null_Variables(k) = myFieldName(i)
                End If
            End If
            .Close
        End With
    Next i
    NumberNullVariables = k
    Err.Clear
    On Error Resume Next
    If NumberNullVariables > 0 Then
        For k = 1 To NumberNullVariables
            dbXbuild.Execute ("Alter Table [Evaluate_SUM] Drop Column [" & _
                Mid$(Null_Variables(k), 1, Len(Null_Variables(k)) - 1) & "O" & "]")
            dbXbuild.Execute ("Alter Table [Evaluate_SUM] Drop Column [" & _
                Mid$(Null_Variables(k), 1, Len(Null_Variables(k)) - 1) & "P" & "]")
            dbXbuild.Execute ("Alter Table [Evaluate_SUM] Drop Column [" & _
                Mid$(Null_Variables(k), 1, Len(Null_Variables(k)) - 1) & "M" & "]")
            dbXbuild.Execute ("Alter Table [Evaluate_SUM] Drop Column [" & _
                Mid$(Null_Variables(k), 1, Len(Null_Variables(k)) - 1) & "S" & "]")
        Next k
    End If
    Err.Clear
    On Error GoTo Error1
    
    Set myrsGrowth = dbXbuild.OpenRecordset("Evaluate_SUM")
    With myrsGrowth
        NumberOfVariables = .Fields.Count
        If NumberOfVariables = 0 Then
            Call MsgBox("The file does not have any data for plotting.", vbOKOnly + vbCritical)
            CreateEvaluateListTable_DontDoThis = True
            If PaulsVersion = True Then
                End
            End If
            Exit Sub
       Else
            CreateEvaluateListTable_DontDoThis = False
       End If
        Dim myFieldName2() As String
        ReDim myFieldName(NumberOfVariables + 1)
        ReDim myFieldName2(NumberOfVariables + 1)
        If Not (.EOF And .BOF) Then
            For i = 0 To NumberOfVariables - 1
                myFieldName(i) = .Fields(i).Name
            Next i
        End If
        .Close
    End With

    Set myrsTemp = dbXbuild.OpenRecordset("TempEvaluateList")
    With myrsTemp
        For i = 0 To NumberOfVariables - 1
            .AddNew
            !VariableID = Mid$(myFieldName(i), 1, Len(myFieldName(i)) - 1)
            myFieldName2(i) = Mid$(myFieldName(i), 1, Len(myFieldName(i)) - 1)
            .Update
        Next i
        .Close
    End With
    ''''''
    'myrsGrowth.Close
    Dim myNewList() As String
    Dim bb As Integer
    Dim Number_myNewList As Integer
    bb = 0
    For i = 0 To NumberOfVariables - 1
        Set myrsTemp = dbXbuild.OpenRecordset("Select * From TempEvaluateList Where VariableID = '" & _
            myFieldName2(i) & "'")
        If Not (myrsTemp.EOF And myrsTemp.BOF) Then
            myrsTemp.MoveLast
            If myrsTemp.RecordCount = 1 Then
                bb = bb + 1
                ReDim Preserve myNewList(bb)
                myNewList(bb) = myFieldName(i)
            End If
        End If
        myrsTemp.Close
    Next i
    Number_myNewList = bb
    For i = 1 To Number_myNewList
        dbXbuild.Execute ("Alter Table [Evaluate_SUM] Drop Column [" & _
            myNewList(i) & "]")
        dbXbuild.Execute ("Delete * From TempEvaluateList Where VariableID = '" & _
            Mid$(myNewList(i), 1, Len(myNewList(i)) - 1) & "'")
    Next i
        Set myrsGrowth = dbXbuild.OpenRecordset("Evaluate_SUM")
        NumberOfVariables = myrsGrowth.Fields.Count
        myrsGrowth.Close
    
    ''''''
    Set myrsTemp = dbXbuild.OpenRecordset("Select Distinct [VariableID] From [TempEvaluateList]")
    With myrsTemp
        If Not (.EOF And .BOF) Then
            .MoveLast
            NumberOfVariables = .RecordCount
            ReDim myFieldName(NumberOfVariables + 1)
            .MoveFirst
            i = 1
            Do While Not .EOF
                myFieldName(i) = !VariableID
                i = i + 1
                .MoveNext
            Loop
        End If
        .Close
    End With
    dbXbuild.Execute "Drop Table [TempEvaluateList]"
    
    For i = 1 To NumberOfVariables
        Set myrsDATA = dbXbuild.OpenRecordset("Select * from [Data] Where [Code] = '" & myFieldName(i) & "'")
        ReDim Preserve VariableDescription(i)
        Dim X As Integer
        If Not (myrsDATA.EOF And myrsDATA.BOF) Then
            k = k + 1
            VariableDescription(i) = myrsDATA.Fields("Description").Value
        Else
            VariableDescription(i) = RemoveSymbols(myFieldName(i))
        End If
      '  If Len(VariableDescription(i)) > 50 Then VariableDescription(i) = Left$(VariableDescription(i), 50)
    Next i
    
    Set myrsTemp = dbXbuild.OpenRecordset("Evaluate_List")
    With myrsTemp
        For i = 1 To NumberOfVariables
            .AddNew
            !VariableID = myFieldName(i)
            !VariableDescription = VariableDescription(i)
            .Update
        Next i
        .Close
    End With
    Err.Clear
    On Error Resume Next
    myrsTemp.Close
    myrsGrowth.Close
    myrsDATA.Close
Exit Sub
Error1: MsgBox Err.Description & " in CreateEvaluateListTable."
    Unload frmWait2
    Unload frmWait3
    CreateEvaluateListTable_DontDoThis = True
    Screen.MousePointer = vbDefault
End Sub

Public Sub FillEvaluate(prmNumberLine As Integer)
Dim myrs As Recordset
Dim k As Integer
Dim m As Integer
Dim TableVariable()
Dim NumberOfRecordsInTable As Integer
Dim i As Integer
ReDim TableVariable(NumberOUTVariables + 6, NumberReadlines)
Dim n As Integer
    On Error GoTo Error1
    n = 1
    LineSpaces(0) = 1
    For k = prmNumberLine To NumberReadlines
        If Trim(FileLoc(k)) = "" Then
            k = k + 1
            'Exit For
        ElseIf IsNumeric(Mid(Trim(FileLoc(k)), 1, 1)) = False Then
            k = k + 1
            'Exit For
        Else
            'For m = 1 To NumberOUTVariables + 2
            For m = 1 To NumberOUTVariables
                TableVariable(m, n) = IIf(Len(Trim(Mid$(FileLoc(k), _
                    LineSpaces(m), LineSpaces(m + 1) - LineSpaces(m)))) = 0, -99, _
                    Trim(Mid$(FileLoc(k), LineSpaces(m), LineSpaces(m + 1) - LineSpaces(m))))
            Next m
        n = n + 1
        End If
        
   Next k
    NumberOfRecordsInTable = n - 1
    NumberOfRecordsInOUTTable = NumberOfRecordsInTable
    
    ReDim VariablesOUTName(NumberOUTVariables + 1)
    Set myrs = dbXbuild.OpenRecordset("Evaluate_SUM")
    With myrs
        For n = 1 To NumberOfRecordsInTable
            .AddNew
             For k = 0 To NumberOUTVariables - 1
                
                .Fields(k).Value = Val(TableVariable(k + 1, n))
                VariablesOUTName(k) = .Fields(k).Name
            Next k
                .Update
            Next n
    End With
    myrs.Close
     For k = 0 To NumberOUTVariables - 1
        dbXbuild.Execute "UPDATE [Evaluate_SUM]" & _
            " SET [" & VariablesOUTName(k) & "] = Null Where " & _
            "[" & VariablesOUTName(k) & "] = -99"
        dbXbuild.Execute "UPDATE [Evaluate_SUM]" & _
            " SET [" & VariablesOUTName(k) & "] = Null Where " & _
            "[" & VariablesOUTName(k) & "] = -99.0"
        dbXbuild.Execute "UPDATE [Evaluate_SUM]" & _
            " SET [" & VariablesOUTName(k) & "] = Null Where " & _
            "[" & VariablesOUTName(k) & "] = -99.00"
        dbXbuild.Execute "UPDATE [Evaluate_SUM]" & _
            " SET [" & VariablesOUTName(k) & "] = Null Where " & _
            "[" & VariablesOUTName(k) & "] = -99.000"
        dbXbuild.Execute "UPDATE [Evaluate_SUM]" & _
            " SET [" & VariablesOUTName(k) & "] = Null Where " & _
            "[" & VariablesOUTName(k) & "] = -9"
    Next k
    Err.Clear
    On Error Resume Next
    myrs.Close

Exit Sub
Error1: MsgBox Err.Description & "in " & "FillEvaluate."
End Sub


Public Sub FillGROWTH(OutputFile1 As String, prmFirstLine As Integer, prmNumberLine As Integer)
Dim myrs As Recordset
Dim myrsSelect As Recordset
Dim k As Integer
Dim m As Integer
Dim TableVariable()
Dim NumberOfTable As Integer
Dim MaxNumberOfTable As Integer
Dim NumberOfVariableInTable As Integer
Dim NumberOfRecordsInTable As Integer
Dim i As Long
ReDim TableVariable(NumberOUTVariables + 6, NumberReadlines + 1)
Dim n As Long
Dim MyReadRun As Integer

    On Error GoTo Error1
    LineSpaces(0) = 1
    For k = prmFirstLine To prmNumberLine
        
        If Trim(UCase(Mid$(FileLoc(k), 1, 6))) = "" Then
            'k = k + 1
           ' Exit For
        ElseIf Trim(UCase(Mid$(FileLoc(k), 1, 4))) = "*RUN" Then
            MyReadRun = Val(Mid$(FileLoc(k), 6, 5))
       
        
        ElseIf Trim(UCase(Mid$(FileLoc(k), 1, 4))) = "@YEA" Then
            k = k + 1
            n = 1
            Do While k <= prmNumberLine
             If k = 3002 Then
                k = 3002
            End If
                
                For m = 1 To NumberOUTVariables + StartVar + 1
                    TableVariable(m + 1, n) = IIf(Len(Trim(Mid$(FileLoc(k), _
                        LineSpaces(m - 1), LineSpaces(m) - LineSpaces(m - 1)))) = 0, -99, _
                        Trim(Mid$(FileLoc(k), LineSpaces(m - 1), LineSpaces(m) - LineSpaces(m - 1))))
                    If m > 1 And IsNumeric(TableVariable(m + 1, n)) = False Then
                        TableVariable(m + 1, n) = -99
                    End If
                Next m
                n = n + 1
                k = k + 1
            Loop
   
   
        
        ElseIf Trim(UCase(Mid$(FileLoc(k), 1, 4))) = "@DAT" Then
            k = k + 1
            n = 1
            Do While k <= prmNumberLine
                For m = 1 To NumberOUTVariables + StartVar + 1
                    TableVariable(m + 1, n) = IIf(Len(Trim(Mid$(FileLoc(k), _
                        LineSpaces(m - 1), LineSpaces(m) - LineSpaces(m - 1)))) = 0, -99, _
                        Trim(Mid$(FileLoc(k), LineSpaces(m - 1), LineSpaces(m) - LineSpaces(m - 1))))
                    If m > 1 And IsNumeric(TableVariable(m + 1, n)) = False Then
                        TableVariable(m + 1, n) = -99
                    End If
                Next m
                n = n + 1
                k = k + 1
            Loop
        End If

   
   
   
   Next k
    NumberOfRecordsInTable = n - 3
    NumberOfRecordsInOUTTable = NumberOfRecordsInTable
    Dim DateString As String
    Set myrs = dbXbuild.OpenRecordset(OutputFile1)
    With myrs
        'was not updated "If CDAYexists = True"
        If CDAYexists = True Then
            For m = 1 To NumberOfRecordsInTable + 1
                .AddNew
                .Fields("RunNumber") = MyReadRun
               ' DateString = Trim(Mid$(FileLoc(prmNumberLine + m - 1), 1, LineSpaces(1)))
                DateString = TableVariable(2, m)
                If Len(DateString) = 4 Then
                    DateString = "0" & DateString
                End If
                If Len(DateString) = 3 Then
                    DateString = "00" & DateString
                End If
                If Len(DateString) = 2 Then
                    DateString = "000" & DateString
                End If
                If Len(DateString) = 1 Then
                    DateString = "0000" & DateString
                End If
                
                ' VSH
                'If Val(Mid$(DateString, 1, 2)) > 15 Then
                If Val(Mid$(DateString, 1, 2)) > 20 Then
                    DateString = "19" & _
                        Mid$(DateString, 1, 2) & _
                        " " & Mid$(DateString, 3)
                Else
                    DateString = "20" & _
                        Mid$(DateString, 1, 2) & _
                        " " & Mid$(DateString, 3, 3)
                End If

               .Fields("Date") = DateString
                .Fields("CDAY") = TableVariable(3, m)
                For k = 1 To NumberOUTVariables
                    If VariablesOUTName(k) <> "" Then
                        .Fields(VariablesOUTName(k)).Value = Val(TableVariable(k + 3, m))
                    End If
                Next k
                .Update
            Next m
        Else
            For m = 1 To NumberOfRecordsInTable
                .AddNew
                .Fields("RunNumber") = MyReadRun
                DateString = TableVariable(2, m)
                If Len(DateString) = 4 Then
                    DateString = "0" & DateString
                End If
                If Len(DateString) = 5 Then
                ' VSH
                    'If Val(Mid(DateString, 1, 2)) > 15 Then
                    If Val(Mid(DateString, 1, 2)) > 20 Then
                        DateString = 19 & _
                            Mid(DateString, 1, 2) & _
                            " " & Mid(DateString, 3)
                    Else
                        DateString = 20 & _
                            Mid(DateString, 1, 2) & _
                            " " & Mid(DateString, 3, 3)
                    End If
                ElseIf Len(DateString) = 10 Then
                    DateString = Replace(DateString, "  ", " ")
                ElseIf Len(DateString) < 5 Then
                    DateString = ""
                End If
                
                .Fields("Date") = DateString
                .Fields("DAS") = TableVariable(3, m)
                
                If DAPexists = True Then
                    .Fields("DAP") = TableVariable(4, m)
                End If
                For k = 1 To NumberOUTVariables '- StartVar - 1
                    If VariablesOUTName(k) <> "" Then
                    .Fields(VariablesOUTName(k)).Value = Val(TableVariable(k + StartVar + 2, m))
                    End If
                Next k
                .Update
            Next m
        End If
    End With
    myrs.Close
    For k = 1 To NumberOUTVariables
        dbXbuild.Execute "UPDATE [" & OutputFile1 & "]" & _
            " SET [" & VariablesOUTName(k) & "] = Null Where " & _
            "[" & VariablesOUTName(k) & "] = -99"
        dbXbuild.Execute "UPDATE [" & OutputFile1 & "]" & _
            " SET [" & VariablesOUTName(k) & "] = Null Where " & _
            "[" & VariablesOUTName(k) & "] = -99.0"
        dbXbuild.Execute "UPDATE [" & OutputFile1 & "]" & _
            " SET [" & VariablesOUTName(k) & "] = Null Where " & _
            "[" & VariablesOUTName(k) & "] = -99.00"
        dbXbuild.Execute "UPDATE [" & OutputFile1 & "]" & _
            " SET [" & VariablesOUTName(k) & "] = Null Where " & _
            "[" & VariablesOUTName(k) & "] = -99.000"
        dbXbuild.Execute "UPDATE [" & OutputFile1 & "]" & _
            " SET [" & VariablesOUTName(k) & "] = Null Where " & _
            "[" & VariablesOUTName(k) & "] = -9"
    Next k
           
            dbXbuild.Execute "DELETE * FROM [" & OutputFile1 & "]" & _
            " Where " & _
            "[DATE] = ''"

    
    If CDAYexists = True Then
                dbXbuild.Execute "DELETE * FROM [" & OutputFile1 & "]" & _
            " Where " & _
            "[CDAY] = -99"
    
    Else
                dbXbuild.Execute "DELETE * FROM [" & OutputFile1 & "]" & _
            " Where " & _
            "[DAS] = -99"
    End If
    
    Err.Clear
    On Error Resume Next
    myrs.Close
    myrsSelect.Close

Exit Sub
Error1:
'MsgBox Err.Description & " in " & "FillGROWTH" & " k=" & k & " m=" & m
End Sub





Public Sub CreateTables(myExper As String)
    On Error Resume Next
    dbXbuild.Execute "Drop Table [" & Replace(myExper, ".", " ") & "_Start]"
    dbXbuild.Execute "Drop Table [" & Replace(myExper, ".", " ") & "]"
    dbXbuild.Execute "Create Table [" & Replace(myExper, ".", " ") & "_Start] ( [Section] Integer, [TheOrder] Text (8) ,[DATE] Integer,[TRNO] Integer)"
    dbXbuild.Execute "Create Table [" & Replace(myExper, ".", " ") & "] ( [Section] Integer, [TheOrder] Text (8), [DATE] Integer, [TRNO] Integer)"

End Sub


Public Sub Create_OneOutTable(prmOutFile As String)
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim m As Integer
Dim myrs As Recordset
Dim NumberOfVariables As Integer
Dim prmVariablesNames() As String
Dim MaxNumberFieldsNames As Integer
Dim Total_MaxNumberFieldsNames As Integer

    On Error Resume Next
    dbXbuild.Execute "Drop Table [" & prmOutFile & "_OUT]"
    Err.Clear
    On Error GoTo Error1
    
        Set myrs = dbXbuild.OpenRecordset(prmOutFile & "_OUT" & 1)
        With myrs
            MaxNumberFieldsNames = .Fields.Count
            ReDim prmVariablesNames(MaxNumberFieldsNames)
            'k = 0
            For j = 0 To MaxNumberFieldsNames - 1
                prmVariablesNames(j) = .Fields(j).Name
            Next j
        End With
        Total_MaxNumberFieldsNames = j - 1
    
    Dim NameIsFound As Boolean
    If NumberTempTables > 1 Then
        For i = 2 To NumberTempTables
            Set myrs = dbXbuild.OpenRecordset(prmOutFile & "_OUT" & i)
            With myrs
               MaxNumberFieldsNames = .Fields.Count
                For j = 0 To MaxNumberFieldsNames - 1
                    NameIsFound = False
                    For m = 0 To Total_MaxNumberFieldsNames
                        If prmVariablesNames(m) = .Fields(j).Name Then
                            NameIsFound = True
                            Exit For
                        End If
                        
                    Next m
                    If NameIsFound = False Then
                        Total_MaxNumberFieldsNames = Total_MaxNumberFieldsNames + 1
                        ReDim Preserve prmVariablesNames(Total_MaxNumberFieldsNames + 1)
                        prmVariablesNames(Total_MaxNumberFieldsNames) = .Fields(j).Name
                    End If
                Next j
                .Close
            End With
        Next i
    End If
    
    dbXbuild.Execute "Create Table [" & prmOutFile & "_OUT]([RunNumber] Integer, [DATE] Text)"

    For j = 2 To Total_MaxNumberFieldsNames
             dbXbuild.Execute "Alter Table [" & prmOutFile & "_OUT] Add Column [" & _
        prmVariablesNames(j) & "] Single"

    Next j
    
    
    For i = 1 To NumberTempTables
        dbXbuild.Execute "Insert Into [" & prmOutFile & "_OUT] " & _
                "Select * From [" & prmOutFile & "_OUT" & i & "]"

    Next i
    If NumberTempTables = 1 Then
        Set myrs = dbXbuild.OpenRecordset(prmOutFile & "_OUT1")
        myrs.Close
    End If
    For i = 1 To NumberTempTables
        dbXbuild.Execute "Drop Table [" & prmOutFile & "_OUT" & i & "]"

    Next i

    
    Exit Sub
Error1:     MsgBox Err.Description & " in Create_OneOutTable/frmMain."
End Sub


