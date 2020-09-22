Attribute VB_Name = "modReport"
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public strfile As String
Sub loadDAOData(frm As Form)

Dim dbsFIELDNAMES As DAO.Database
Dim tdfTest As DAO.TableDef
Dim fldloop As DAO.Field

startagain:
On Error Resume Next
'load the commondialog form for opening a file
frm.CommonDialog1.ShowOpen
'assign chosen file to this var
strfile = frm.CommonDialog1.FileName

DoEvents

If strfile = "" Then GoTo startagain

frm.Show
frm.lstTables.Clear
Dim X As Integer
frm.grdQuery.ClearFields
'sub to add a new datagrid control, add a ado data control, bind the grid
'to the data control and then format the grid
Set dbsFIELDNAMES = OpenDatabase(strfile)
'Set tdfTest = dbsFIELDNAMES.TableDefs![STAFF MEMBER]

For Each tdfTest In dbsFIELDNAMES.TableDefs
    
    frm.lstTables.AddItem tdfTest.Name  'assign the field name to the grid column header
    
Next tdfTest

End Sub


