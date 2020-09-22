VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Initalization Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private mvarDataBaseName As String
Private mvarConnectionString As String

Public Type DataConnection
    RS As ADODB.Recordset
    CON As ADODB.Connection
End Type


Public Property Let ConnectionString(ByVal vData As String)
    mvarConnectionString = vData
End Property

Public Property Get ConnectionString() As String
    ConnectionString = mvarConnectionString
End Property

Public Property Let DataBaseName(ByVal vData As String)
    mvarDataBaseName = vData
    mvarConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & mvarDataBaseName & ";"
End Property

Public Property Get DataBaseName() As String
    DataBaseName = mvarDataBaseName
End Property

'Opening / Closing / Misc Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function CreateConnection(sTable As String, Optional sSQL As String) As DataConnection
'Creates a new connection to the database
Dim DB As DataConnection
If mvarDataBaseName = "" Then
    MsgBox "No database specified!", vbInformation, "DataHelper"
    Exit Function
End If

    Set DB.RS = New ADODB.Recordset
    Set DB.CON = New ADODB.Connection
    SelectRecords DB, sTable, sSQL
End Function

Public Function OpenDB(sTable As String, sConnection As String, CON As ADODB.Connection) As ADODB.Recordset
'Opens a connection to a database
    Dim sConnect As String
    Dim sSQL As String
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    sSQL = "select * from " & sTable

    CON.Open sConnection
    RS.LockType = adLockBatchOptimistic
    RS.Open sSQL, CON
    Set OpenDB = RS
    'Set RS = Nothing
End Function

Public Sub CloseDB(DB As DataConnection)
'closes a connection to a database
On Error Resume Next
DB.RS.Close
DB.CON.Close
End Sub


Public Function GetRecordCount(DB As DataConnection) As Long
'returns the total number of records in the record set
Dim L As Long
If isRecordSetEmpty(DB) = True Then
    GetRecordCount = 0
    Exit Function
End If
Screen.MousePointer = 11
    Do While Not DB.RS.EOF
        L = L + 1
        DoEvents
        DB.RS.MoveNext
    Loop
DB.RS.MoveFirst
Screen.MousePointer = 0
GetRecordCount = L
End Function

Public Function isRecordSetEmpty(DB As DataConnection) As Boolean
'Returns True if recordset is empty
If DB.RS.BOF = True And DB.RS.EOF = True Then
    isRecordSetEmpty = True
Else
    isRecordSetEmpty = False
End If
End Function

Public Sub SelectRecords(DB As DataConnection, sTable As String, Optional sWhat As String)
'Selects a specific record set
CloseDB DB
Dim sSQL As String
If sWhat <> "" And sWhat <> "*" Then
    sSQL = "Select * FROM " & sTable & " " & sWhat
Else
    sSQL = "Select * from " & sTable
End If

    Set DB.RS = New ADODB.Recordset
    Set DB.CON = New ADODB.Connection
    DB.RS.LockType = adLockBatchOptimistic
    DB.RS.CursorType = adOpenDynamic
    DB.CON.Open mvarConnectionString
    DB.RS.Open sSQL, DB.CON, adOpenDynamic, adLockOptimistic
End Sub

'Table Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub RenameTable(sConnect As String, sFrom As String, sTo As String)
On Error GoTo EH
Dim DB As Connection
Dim sqlStatement As String
Set DB = New Connection
DB.Open sConnect
    sqlStatement = "SELECT " & ResolveTable(sFrom) & ".* INTO " & ResolveTable(sTo) & " FROM " & ResolveTable(sFrom)
    Call DB.Execute(sqlStatement)
    sqlStatement = "DROP TABLE " & ResolveTable(sFrom)
    Call DB.Execute(sqlStatement)
DB.Close
Exit Sub
EH:
MsgBox Err.Number & " - " & Err.Description, vbCritical, "Rename Table"
Exit Sub
End Sub

Private Function ResolveTable(inputTable As String) As String
    ResolveTable = IIf(InStr(1, inputTable, " ") <> 0 Or IsNumeric(Left(inputTable, 1)), "[" & inputTable & "]", inputTable)
End Function

Public Sub TablesToCombo(CMB As Object)
 Dim DB As Connection
 Dim RS As Recordset

Set DB = New Connection
Set RS = New Recordset

DB.Open mvarConnectionString
Set RS = DB.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
        If Not RS Is Nothing Then
            Do While Not RS.EOF
                If UCase(Left(RS!Table_name, 4)) <> "MSYS" Then
                    If UCase(Left(RS!Table_name, 11)) <> "SWITCHBOARD" Then
                        newTableName = RS!Table_name
                        If newTableName <> "" Then
                        CMB.AddItem newTableName
                        End If
                    End If
                End If
                RS.MoveNext
            Loop
        End If
RS.Close
DB.Close

End Sub

Public Sub ClearTable(sTable As String)
Dim DB As Connection
On Error GoTo EH
Set DB = New Connection
DB.Open mvarConnectionString
DB.Execute "DELETE * FROM " & sTable
DB.Close
Exit Sub
EH:
MsgBox Err.Number & " - " & Err.Description, vbCritical, "Clear Table"
Exit Sub
End Sub

'Field Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub FieldsToCombo(CMB As Object, sTable As String, Optional bNoDuplicates As Boolean)
'Loads all the fields to the listview

 Dim DB As Connection
 Dim RS As Recordset
 Dim I As Integer
 Dim H As Integer
 Dim bFound As Boolean
On Error GoTo EH
Set DB = New Connection
Set RS = New Recordset

CMB.Clear
'add the parts list
Set RS = OpenDB(sTable, mvarConnectionString, DB)
For I = 0 To RS.Fields.Count - 1
bFound = False
If bNoDuplicates = True Then
    For H = 0 To CMB.ListCount - 1
        If LCase(RS.Fields(I).Name) = LCase(CMB.List(H)) Then
            bFound = True
            Exit For
        End If
    Next H
End If
If bFound = False Then
    CMB.AddItem RS.Fields(I).Name
End If
Next I

RS.Close
DB.Close
If CMB.ListCount > 0 Then
CMB.ListIndex = 0
End If
Exit Sub
EH:
MsgBox Err.Number & " - " & Err.Description, vbCritical, "Fields To Combo"
Exit Sub
End Sub


'Datagrid Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub SortDBGrid(sFRm As Object, grdObject As Object, adoPrimaryRS As Recordset, ColIndex As Integer)
    Dim strColName As String
    Dim sTMP As String
    Dim Ipos As Long
    Static bSortAsc As Boolean
    Static strPrevCol As String
On Error GoTo EH
    strColName = grdObject.Columns(ColIndex).DataField
    If strColName = strPrevCol Then


        If bSortAsc Then
            adoPrimaryRS.Sort = strColName & " DESC"
            bSortAsc = False
        Else
            adoPrimaryRS.Sort = strColName
            bSortAsc = True
        End If
    Else
        adoPrimaryRS.Sort = strColName
        bSortAsc = True
    End If
    
Exit Sub
EH:
MsgBox Err.Description, vbCritical, "Generating Sort Order"
Exit Sub
End Sub


