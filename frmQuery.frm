VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmQuery 
   Caption         =   "Ad Hoc Query"
   ClientHeight    =   5550
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8235
   Icon            =   "frmQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   5505
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   3480
         Top             =   2220
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   7050
         Top             =   1860
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Choose the access database"
         Filter          =   "Access Database(*.mdb)|*.mdb"
         InitDir         =   "c:\"
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "RUN QUERY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4080
         TabIndex        =   16
         Top             =   2610
         Width           =   3750
      End
      Begin VB.CommandButton cmdReport 
         Caption         =   "View Report"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   330
         TabIndex        =   15
         Top             =   2610
         Width           =   3750
      End
      Begin VB.TextBox txtCriteria 
         Height          =   285
         Left            =   5820
         TabIndex        =   14
         Top             =   2280
         Width           =   1065
      End
      Begin VB.ComboBox cmboOperator 
         Height          =   315
         Left            =   5820
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1920
         Width           =   825
      End
      Begin VB.TextBox txtCriteriaField 
         Height          =   285
         Left            =   3060
         TabIndex        =   12
         Top             =   1920
         Width           =   1395
      End
      Begin VB.ListBox lstCriteria 
         DragIcon        =   "frmQuery.frx":0442
         Height          =   840
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Frame Frame10 
         Height          =   1515
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   3225
         Begin VB.ListBox lstTableSelected 
            DragIcon        =   "frmQuery.frx":0884
            Height          =   1035
            Left            =   1710
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   390
            Width           =   1395
         End
         Begin VB.ListBox lstTables 
            DragIcon        =   "frmQuery.frx":0CC6
            Height          =   1035
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   390
            Width           =   1395
         End
         Begin VB.Label Label16 
            Caption         =   "Selected fields"
            Height          =   195
            Left            =   1710
            TabIndex        =   10
            Top             =   180
            Width           =   1305
         End
         Begin VB.Label Label15 
            Caption         =   "Available fields"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   210
            Width           =   1275
         End
      End
      Begin VB.Frame Frame9 
         Height          =   1515
         Left            =   3600
         TabIndex        =   1
         Top             =   120
         Width           =   4455
         Begin VB.ListBox lstFields 
            DragIcon        =   "frmQuery.frx":1108
            Height          =   1035
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   390
            Width           =   1395
         End
         Begin VB.ListBox lstSelected 
            DragIcon        =   "frmQuery.frx":154A
            Height          =   1035
            Left            =   1710
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   390
            Width           =   2655
         End
         Begin VB.Label Label8 
            Caption         =   "Available fields"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   210
            Width           =   1275
         End
         Begin VB.Label Label7 
            Caption         =   "Selected fields"
            Height          =   195
            Left            =   1710
            TabIndex        =   4
            Top             =   210
            Width           =   1300
         End
      End
      Begin MSDataGridLib.DataGrid grdQuery 
         Bindings        =   "frmQuery.frx":198C
         Height          =   1905
         Left            =   120
         TabIndex        =   17
         Top             =   3060
         Visible         =   0   'False
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   3360
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label19 
         Caption         =   "Criteria Value"
         Height          =   225
         Left            =   4500
         TabIndex        =   21
         Top             =   2310
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "Selection Criteria"
         Height          =   285
         Left            =   4500
         TabIndex        =   20
         Top             =   1980
         Width           =   1245
      End
      Begin VB.Label Label17 
         Caption         =   "Selection Criteria"
         Height          =   285
         Left            =   300
         TabIndex        =   19
         Top             =   1710
         Width           =   1245
      End
      Begin VB.Label lblDrag 
         Height          =   285
         Left            =   3060
         TabIndex        =   18
         Top             =   2400
         Width           =   1275
      End
   End
   Begin VB.Menu mnuNew 
      Caption         =   "New Query"
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strSQLQuery As String
Dim objWrd As Word.Application
Dim mboolWordRunning As Boolean ' Flag For final word unload
Private WithEvents objADODC As Adodc
Attribute objADODC.VB_VarHelpID = -1
Sub modReformatForm()
'sub to clear the appropriate controls/
lstFields.Clear
lstTables.Clear
lstTableSelected.Clear
lstSelected.Clear
strfile = ""
strSQLQuery = ""
 lstTables.Enabled = True
DoEvents
txtCriteriaField = ""
cmboOperator.Clear
grdQuery.Enabled = True
lstCriteria.Clear
lstCriteria.Enabled = True
txtCriteria = ""

On Error Resume Next
X = Controls.Remove(objADODC) 'remove control added onthe fly
grdQuery.ClearFields
On Error GoTo 0
DoEvents
End Sub


Private Sub modcreateQuery()
Dim strstrSqlString As String
Dim strTableName As String
Dim strTableQuery As String

Select Case strFieldType
    Case "DAT"
        txtCriteria = "'" & txtCriteria & "'"
    Case "STR"
        txtCriteria = "'" & txtCriteria & "'"
    Case Else
        txtCriteria = txtCriteria
End Select


strTableName = lstTableSelected.Text
strTableQuery = ""
strSQLQuery = ""

'loop through list box to get selected fields
For i = 0 To lstSelected.ListCount - 1
        strTableQuery = strTableQuery & lstSelected.List(i) & ","
Next i

If Left$(strTableQuery, 1) = "," Then strTableQuery = Right$(strTableQuery, Len(strTableQuery) - 1)
If Right$(strTableQuery, 1) = "," Then strTableQuery = Left$(strTableQuery, Len(strTableQuery) - 1)
              
strSQLQuery = "SELECT " & strTableQuery & " FROM [" & strTableName & "] WHERE [" & strTableName & "].[" & txtCriteriaField & "] " & cmboOperator.Text & " " & txtCriteria

End Sub
Sub modCreateReport()
Dim X As Integer
Dim newline As String


Screen.MousePointer = vbHourglass
'Startup Word if not started, or switch to existing one
Call modGetWord
objWrd.Application.ScreenUpdating = False
objWrd.Visible = False
objWrd.Application.WindowState = wdWindowStateMinimize
objWrd.Documents.Add
newline = Chr$(13) & Chr$(10)

objWrd.Selection.InsertAfter "Query Builder - Adhoc Report" 'set title of report
objWrd.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter 'align title in centre of page
objWrd.Selection.Font.Bold = True 'bold the text
objWrd.Selection.Font.Underline = wdUnderlineSingle 'underscore
objWrd.Selection.Font.Name = "Arial" 'set font to arial
objWrd.Selection.Font.Size = 12 'set size of font to 12
objWrd.Selection.InsertAfter newline & newline 'insert 2 new lines
objWrd.Selection.MoveDown wdLine, 2  'move down the two lines as word extends when insertion
objWrd.Selection.InsertAfter "Adhoc SQL : " & objADODC.RecordSource 'insert the SQL Text into the report
objWrd.Selection.Font.Bold = True 'bold the text
objWrd.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft 'left align the text
objWrd.Selection.Font.Size = 9 'set font size to 9
objWrd.Selection.InsertAfter newline & newline 'insert 2 new lines
objWrd.Selection.MoveDown wdLine, 2 'move down the two lines

objWrd.ActiveDocument.PageSetup.Orientation = wdOrientLandscape 'set the page orientation to landscape
objWrd.ActiveWindow.View.Type = wdPageView 'change view of report to page view
'format table to number of columns in grid
objWrd.ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=3, NumColumns:=grdQuery.Columns.Count
objWrd.Selection.SelectRow 'select the entire top row
objWrd.Selection.Font.Underline = wdUnderlineNone 'no underscore
    With objWrd.Selection.Cells
        With .Shading
            .Texture = wdTexture25Percent
            .ForegroundPatternColorIndex = wdAuto
            .BackgroundPatternColorIndex = wdWhite
        End With
    End With
objWrd.Selection.Rows.HeadingFormat = wdToggle 'set first row as header so that they show on subsequent pages of report
objWrd.Selection.MoveLeft

'loop round the grid to get the column headers to make the report headers
For i = 0 To grdQuery.Columns.Count - 1
    objWrd.Selection.InsertAfter grdQuery.Columns(i).Caption
    objWrd.Selection.MoveRight wdCell, 1
Next

objWrd.Selection.SelectRow
objWrd.Selection.Font.Underline = wdUnderlineNone
objWrd.Selection.Font.Bold = False
objWrd.Selection.MoveLeft
objADODC.Recordset.MoveFirst 'move to the first record int he ado control

'loop round the resultset of the contol and insert the values into the table
Do While Not objADODC.Recordset.EOF
        For X = 0 To objADODC.Recordset.Fields.Count - 1
        objWrd.Selection.InsertAfter "" & objADODC.Recordset.Fields(X).Value
        objWrd.Selection.Font.Bold = False
        objWrd.Selection.Font.Underline = wdUnderlineNone

            'set the alignment of the cell dependig on the type of field eg. integer etc
            If objADODC.Recordset.Fields(X).Type = adVarChar Then
                    objWrd.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            ElseIf objADODC.Recordset.Fields(X).Type = adInteger Then
                    objWrd.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
            ElseIf objADODC.Recordset.Fields(X).Type = adDate Then
                    objWrd.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            ElseIf objADODC.Recordset.Fields(X).Type = adBigInt Then
                    objWrd.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
            ElseIf objADODC.Recordset.Fields(X).Type = adNumeric Then
                    objWrd.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
            Else: objWrd.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            End If
            
        objWrd.Selection.MoveRight wdCell, 1 'move to the next cell in the table
        Next
    objWrd.Selection.SelectRow
    objWrd.Selection.InsertRows 1 'insert a new row
    objWrd.Selection.MoveLeft Unit:=wdCharacter, Count:=1

objADODC.Recordset.MoveNext 'next record int he ado control
Loop

'delete any empty rows which may be in the table
Do While objWrd.Selection.Information(wdWithInTable)
     objWrd.Selection.SelectRow
     objWrd.Selection.Rows.Delete
Loop

objWrd.Application.Activate 'activate word
objWrd.ActiveDocument.PrintPreview 'preview the report
objWrd.ActiveWindow.ActivePane.View.Zoom.Percentage = 100 'zoom the report to 100%
objWrd.Application.ScreenUpdating = True
objWrd.Application.WindowState = wdWindowStateMaximize
objWrd.Visible = True
Set objWrd = Nothing 'clear object variable
Screen.MousePointer = vbDefault

End Sub

Private Sub modGetWord()

    On Error Resume Next
    Set objWrd = GetObject(, "Word.Application")

    If Err.Number <> 0 Then
        mboolWordRunning = True
    Else
        mboolWordRunning = False
    End If

    Err.Clear
    Call modWordClass

    If mboolWordRunning = True Then
        Set objWrd = New Word.Application
    End If
End Sub

Public Sub modWordClass()

    Const WM_USER = 1024
    Dim hWnd As Long
    
    ' get handle for class opusapp ie. winword.exe
    hWnd = FindWindow("OpusApp", vbNullString)

    If hWnd = 0 Then ' 0 means Word not running.
        Exit Sub
    Else
        ' Word is running so use the SendMessage API function to enter it
        '     in the Running Object Table.
        SendMessage hWnd, WM_USER + 18, 0, 0
    End If

End Sub


Private Sub cmdReport_Click()
Call modCreateReport
End Sub

Private Sub cmdRun_Click()
Call modcreateQuery
DoEvents

Set objADODC = Controls.Add("MSAdodcLib.adodc", "adodcQuery", Frame1)
With objADODC     'format the instantiated command button object
            .Width = 7785
            .Height = 330
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & strfile
            .Left = 120
            .Caption = "Data"
            .Top = 5100
            .Visible = True
            .Enabled = True
            .RecordSource = strSQLQuery
            DoEvents
            .Refresh
End With

Set grdQuery.DataSource = objADODC
grdQuery.Refresh
grdQuery.Visible = True
grdQuery.Enabled = True
cmdReport.Enabled = True
End Sub

Private Sub Form_Load()

Call loadDAOData(frmQuery)

End Sub


Private Sub lstCriteria_DblClick()
txtCriteriaField = lstCriteria.Text
lstCriteria.Enabled = False

Dim dbsFIELDNAMES As DAO.Database
Dim tdfTest As DAO.TableDef
Dim fldloop As DAO.Field
Dim X As Integer

'sub to add a new datagrid control, add a ado data control, bind the grid
'to the data control and then format the grid
Set dbsFIELDNAMES = OpenDatabase(strfile)
Set tdfTest = dbsFIELDNAMES.TableDefs(lstTableSelected.Text)

For Each fldloop In tdfTest.Fields

    If fldloop.Name = Trim(txtCriteriaField) Then
        Select Case fldloop.Type
                Case 3
                    strFieldType = "INT"
                    cmboOperator.Clear
                    cmboOperator.AddItem " < "
                    cmboOperator.AddItem " > "
                    cmboOperator.AddItem " <> "
                    cmboOperator.AddItem " = "
                    cmboOperator.AddItem " <= "
                    cmboOperator.AddItem " >= "
                Case 4
                    strFieldType = "INT"
                    cmboOperator.Clear
                    cmboOperator.AddItem " < "
                    cmboOperator.AddItem " > "
                    cmboOperator.AddItem " <> "
                    cmboOperator.AddItem " = "
                    cmboOperator.AddItem " <= "
                    cmboOperator.AddItem " >= "

                Case 7
                    strFieldType = "FLT"
                    cmboOperator.Clear
                    cmboOperator.AddItem " < "
                    cmboOperator.AddItem " > "
                    cmboOperator.AddItem " <> "
                    cmboOperator.AddItem " = "
                    cmboOperator.AddItem " <= "
                    cmboOperator.AddItem " >= "
                Case 8
                    strFieldType = "DAT"
                    cmboOperator.Clear
                    cmboOperator.AddItem " < "
                    cmboOperator.AddItem " > "
                    cmboOperator.AddItem " <> "
                    cmboOperator.AddItem " = "
                    cmboOperator.AddItem " <= "
                    cmboOperator.AddItem " >= "
                Case 10
                    strFieldType = "STR"
                    cmboOperator.Clear
                    cmboOperator.AddItem " <> "
                    cmboOperator.AddItem " = "
                    cmboOperator.AddItem " LIKE "
                Case 12
                    strFieldType = "MEM"
                    cmboOperator.Clear
                    cmboOperator.AddItem " <> "
                    cmboOperator.AddItem " = "
                    cmboOperator.AddItem " LIKE "
        End Select
    End If
              
Next fldloop

End Sub


Private Sub lstFields_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim DY   ' Declare variable.
   DY = TextHeight("A")   ' Get height of one line.
   lblDrag.Move (Frame9.Left + lstFields.Left), lstFields.Top + Y + DY, lstFields.Width, DY
   lblDrag.Drag   ' Drag label outline.
End Sub


Private Sub lstSelected_DragDrop(Source As Control, X As Single, Y As Single)
    lstSelected.AddItem "[" & lstTableSelected.Text & "].[" & lstFields.Text & "]"
    lstFields.RemoveItem lstFields.ListIndex
End Sub


Private Sub lstTables_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim DY
DY = TextHeight("A")
lblDrag.Move lstTables.Left, lstTables.Top + Y + DY, lstTables.Width, DY
lblDrag.Drag
End Sub


Private Sub lstTableSelected_Click()
Dim dbsFIELDNAMES As DAO.Database
Dim tdfTest As DAO.TableDef
Dim fldloop As DAO.Field
Dim X As Integer

lstFields.Clear
lstCriteria.Clear

Set dbsFIELDNAMES = OpenDatabase(strfile)
Set tdfTest = dbsFIELDNAMES.TableDefs(lstTableSelected.Text)

For Each fldloop In tdfTest.Fields
    
    lstFields.AddItem fldloop.Name
    lstCriteria.AddItem fldloop.Name
    
Next fldloop
End Sub


Private Sub lstTableSelected_DragDrop(Source As Control, X As Single, Y As Single)
 lstTableSelected.AddItem lstTables.Text
 lstTables.Enabled = False
End Sub


Private Sub mnuNew_Click()
modReformatForm
Call loadDAOData(frmQuery)
End Sub


