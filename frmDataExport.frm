VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmDataExport 
   Caption         =   "Form1"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10425
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save file"
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
      Left            =   7200
      TabIndex        =   3
      Tag             =   "1038"
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
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
      Left            =   6240
      TabIndex        =   2
      Tag             =   "1010"
      Top             =   6720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Default         =   -1  'True
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
      Left            =   8160
      TabIndex        =   1
      Tag             =   "1059"
      Top             =   6720
      Width           =   855
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   10610
      _Version        =   393217
      BackColor       =   16777215
      ScrollBars      =   3
      RightMargin     =   50000
      TextRTF         =   $"frmDataExport.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Orientation     =   2
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmDataExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NewFormName As String


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim i As Integer
Dim k As Integer
Dim StartData As Integer
Dim myLenghtOfLine As Integer
Dim NumberOfCuts As Integer
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
    
   ' Printer.Orientation = 2
    Printer.Print " "
    
        Err.Clear
        On Error Resume Next
        Printer.FontName = "Courier New"
        Err.Clear
        On Error GoTo Error1
        Printer.FontSize = 14
        Printer.FontBold = True
        Printer.Print " "
        Printer.Print " "
        Printer.Print "                              Graph Data"
        Printer.Print " "
        Printer.Print " "
        Printer.FontSize = 10
        Printer.FontBold = True
        
        myReadFile (App.path & "\" & "Graph_Data.txt")
        
      '  For i = 1 To NumberReadlines
      '      Printer.Print Spc(1), "Variables:"
      '      Printer.Print Spc(1), FileLoc(i)
      '      If Mid(FileLoc(i), 1, 4) = "Data" Then
       '         StartData = i + 1
      '          Exit For
     '       End If
     '   Next i
        StartData = 2
        If Len(Trim(FileLoc(StartData))) <= 84 Then
            NumberOfCuts = 1
        
        Else
            Dim mm
            
            NumberOfCuts = Round(Len(Trim(FileLoc(StartData))) / 84)
            If NumberOfCuts * 84 > Len(Trim(FileLoc(StartData))) Then
                NumberOfCuts = NumberOfCuts - 1
            End If
            NumberOfCuts = NumberOfCuts + 1
        
        End If
        For k = 0 To NumberOfCuts - 1
            For i = 1 To NumberReadlines
            Printer.Print " ", Mid(FileLoc(i), 1 + 84 * k, 84)
            Next i
            Printer.Print ""
        Next k
    Printer.EndDoc
                                          
                                          
    Exit Sub
Error1:     MsgBox Err.Description
    Exit Sub
ErrHandler:

End Sub

Private Sub cmdSave_Click()
    FileCopy App.path & "\" & "Graph_Data.txt", NewFormName
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 19
        cmdSave_Click
    Case 14
        cmdCancel_Click
    End Select

End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Width = 9450
    Me.Height = 7200

   ' If gsglXFactor <> 1 Or gsglYFactor <> 1 Then
        Call SetDeviceIndependentWindow(Me, 1)
   ' End If
    Me.Width = Screen.Width
    Me.Height = Screen.Height - 300
    FindGraph_Data_File
    Label1.FontSize = Label1.FontSize * 1.5
    'Me.Caption = App.path & "\" & "Graph_Data.txt"
    Me.Caption = NewFormName
    
    'LoadResStrings Me
    Dim LabelTop As Integer
    LabelTop = Me.Label1.Top + Label1.Height + 50 * gsglYFactor
    RichTextBox1.Top = LabelTop
    RichTextBox1.Left = 5
    RichTextBox1.Width = Me.Width - 100 * gsglXFactor
   ' RichTextBox1.Top = 120
    RichTextBox1.Height = Me.Height - _
LabelTop - 800 * gsglYFactor
    If Dir(App.path & "\" & "Graph_Data.txt") <> "" Then
        RichTextBox1.FileName = App.path & "\" & "Graph_Data.txt"
    Else
        MsgBox "No data available."
    End If
    Me.cmdSave.Top = RichTextBox1.Top + RichTextBox1.Height + _
50 * gsglYFactor
    cmdCancel.Top = cmdSave.Top
    Me.Label1.Caption = "File:" & " " & NewFormName
    
    

End Sub

Public Sub FindGraph_Data_File()
Dim fName As String
Dim Array_Data_Graph_file() As String
Dim i As Integer
Dim Number_fNames As Integer
Dim Array_StringFileIndexes() As String
Dim ArrayFileIndexes() As Integer
Dim MaxNumber As Single

On Error GoTo Error1
    fName = Dir(DirectoryToPreview & "Graph_data*.txt", vbNormal + vbDirectory)
    i = 1
    Do While fName <> ""
        If fName <> "_" And fName <> "__" Then
            ReDim Preserve Array_Data_Graph_file(i + 1)
            ReDim Preserve Array_StringFileIndexes(i + 1)
            ReDim Preserve ArrayFileIndexes(i + 1)
            Array_Data_Graph_file(i) = UCase(Replace(fName, ".txt", ""))
            Array_StringFileIndexes(i) = Replace(Array_Data_Graph_file(i), "GRAPH_DATA", "")
            ArrayFileIndexes(i) = Round(Val(Array_StringFileIndexes(i)))
            i = i + 1
        End If
        fName = Dir()
    Loop
    Number_fNames = i - 1
    MaxNumber = -99
    For i = 1 To Number_fNames
        If ArrayFileIndexes(i) > MaxNumber Then
            MaxNumber = ArrayFileIndexes(i)
        End If
    Next i
    If MaxNumber = -99 Then MaxNumber = 0
    NewFormName = DirectoryToPreview & "Graph_data" & MaxNumber + 1 & ".txt"
    
    Exit Sub
Error1: MsgBox Err.Description & " in FindGraph_Data_File/frmDataExport."
End Sub
