VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLinkSqlite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Module Name:  Form_frmLinkSqlite
' Author:       Geoffrey Griffith
' Company:      Imagine Thought Software
' License:
'
'     Copyright © 2026 - Geoffrey L. Griffith
'
'     Permission is hereby granted, free of charge, to any person obtaining a copy
'     of this software and associated documentation files (the “Software”), to deal
'     in the Software without restriction, including without limitation the rights
'     to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'     copies of the Software, and to permit persons to whom the Software is
'     furnished to do so, subject to the following conditions:
'
'     The above copyright notice and this permission notice shall be included in
'     all copies or substantial portions of the Software.
'
'     THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'     IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'     FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'     AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'     LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'     OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'     SOFTWARE.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Option Compare Database
Option Explicit



Private Sub btnAbout_Click()
On Error GoTo HandleErrors

    DoCmd.OpenForm "frmAbout"

ExitMethod:
    Exit Sub
HandleErrors:
    MsgBox "Error: " & Err.Description, vbCritical, "Error " & Nz(Err.Number)
    Resume ExitMethod
End Sub



Private Sub btnAddTable_Click()
On Error GoTo HandleErrors

    Dim iIdx As Long
    Dim sValue As String
    
    If Me.ltsTablesInDb.ItemsSelected.Count < 1 Then
        Exit Sub
    End If
    
    iIdx = Me.ltsTablesInDb.ItemsSelected(0)
    sValue = Me.ltsTablesInDb.ItemData(iIdx)
    Me.ltsTablesInDb.RemoveItem iIdx
    
    Me.lstTablesToAdd.AddItem sValue
    
    Me.ltsTablesInDb.Selected(0) = True
    
ExitMethod:
    Exit Sub
HandleErrors:
    MsgBox "Error: " & Err.Description, vbCritical, "Error " & Nz(Err.Number)
    Resume ExitMethod
End Sub



Private Sub btnRemoveTable_Click()
On Error GoTo HandleErrors

    Dim iIdx As Long
    Dim sValue As String
    
    If Me.lstTablesToAdd.ItemsSelected.Count < 1 Then
        Exit Sub
    End If
    
    iIdx = Me.lstTablesToAdd.ItemsSelected(0)
    sValue = Me.lstTablesToAdd.ItemData(iIdx)
    Me.lstTablesToAdd.RemoveItem iIdx
    
    Me.ltsTablesInDb.AddItem sValue
    
    Me.lstTablesToAdd.Selected(0) = True
    
ExitMethod:
    Exit Sub
HandleErrors:
    MsgBox "Error: " & Err.Description, vbCritical, "Error " & Nz(Err.Number)
    Resume ExitMethod
End Sub



Private Sub btnBrowse_Click()
On Error GoTo HandleErrors

    Dim sPath As String
    Dim sTables As String
    
    sPath = BrowseToSqliteFile()
    If Len(sPath) < 1 Then
        MsgBox "No file was selected.", vbInformation, "No File Selected"
    Else
        Me.txtDbPath = sPath
    End If

ExitMethod:
    Exit Sub
HandleErrors:
    MsgBox "Error: " & Err.Description, vbCritical, "Error " & Nz(Err.Number)
    Resume ExitMethod
End Sub



Public Function BrowseToSqliteFile() As String
On Error GoTo HandleErrors

    Dim sTables As String
    Dim sFile As String
    Dim sql As String
    
    Dim oFSO As Object  ' "Scripting.FileSystemObject"
    Dim cn As Object
    Dim rs As Object
    

    With Application.FileDialog(3) ' 3 = msoFileDialogFilePicker
        .AllowMultiSelect = False
        .Title = "Select a SQLite file to link your database to:"
    
        .Filters.Clear
        .Filters.Add "SQLite Files", "*.sqlite;*.sqlite3;*.db;*.db3;*.s3db;*.sl3"
        .Filters.Add "All Files", "*.*"
        .FilterIndex = 0
        
        If .Show <> True Then
            GoTo ExitMethod
        Else
            sFile = .SelectedItems(1)
        End If
    End With
    
    ' Ensure the file exists
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    If oFSO.FileExists(sFile) Then
        BrowseToSqliteFile = sFile
    Else ' File path invalid
        BrowseToSqliteFile = ""
        GoTo ExitMethod
    End If
    
    ' Clear the list controls
    Me.ltsTablesInDb.RowSource = ""
    Me.ltsTablesInDb.Requery
    Me.lstTablesToAdd.RowSource = ""
    Me.lstTablesToAdd.Requery

    'Create an ADODB connection via the SQLite ODBC driver
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionString = "Driver={SQLite3 ODBC Driver};Database=" & sFile & ";"
    cn.Open

    ' Get the SQLite table names
    sql = "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name;"
    Set rs = cn.Execute(sql)
    
    While (Not rs.EOF) And (Not rs.BOF)
        sTables = rs.Fields(0)
        If Me.chkIgnore Then
            If (sTables <> "sqlite_sequence") Then ' Ignore the "sqlite_sequence"
                Me.ltsTablesInDb.AddItem sTables
            End If
        Else
            Me.ltsTablesInDb.AddItem sTables
        End If
        rs.MoveNext
    Wend
    
    Me.ltsTablesInDb.Selected(0) = True
    
ExitMethod:
On Error Resume Next
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
    Exit Function
HandleErrors:
    MsgBox "Error: " & Err.Description, vbCritical, "Error " & Nz(Err.Number)
    Resume ExitMethod
End Function



Private Function LinkSqliteTable( _
    DatabasePath As String, _
    TableName As String, _
    Optional LocalTableName As String = "", _
    Optional Timeout As Long = 2000 _
    ) As Boolean
On Error GoTo HandleErrors

    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim sConn As String
    Dim sPath As String

    If (Len(LocalTableName) < 1) Then
        LocalTableName = "sqlt_" & TableName
    End If
    
    Set db = CurrentDb
On Error Resume Next
    Set td = db.TableDefs(LocalTableName)
On Error GoTo HandleErrors
    If Not (td Is Nothing) Then
        MsgBox "The table name '" & LocalTableName & "' already exists.", vbInformation, "Table Already Exists"
        GoTo ExitMethod
    End If

    sConn = "ODBC;Driver=SQLite3 ODBC Driver;" & _
            "Database=" & DatabasePath & ";" & _
            "LongNames=0;" & _
            "Timeout=" & Timeout & ";"

    Set td = db.CreateTableDef(LocalTableName)
    td.Connect = sConn
    td.SourceTableName = TableName
    db.TableDefs.Append td
    db.TableDefs.Refresh
    DoEvents
    
    DoCmd.RunCommand acCmdSave
    db.TableDefs.Refresh
    DoEvents
    
    LinkSqliteTable = True
    
ExitMethod:
On Error Resume Next
    Set td = Nothing
    Set db = Nothing
    Exit Function
HandleErrors:
On Error Resume Next
    If Nz(Err.Number, 0) = 3146 Then
        MsgBox "Error: " & Err.Description & vbNewLine & vbNewLine & "Note: Ensure SQLite ODBC driver is installed.", vbCritical, "Error " & Nz(Err.Number)
        Debug.Print "Error " & Nz(Err.Number) & ":" & Err.Description
    Else
        MsgBox "Error: " & Err.Description, vbCritical, "Error " & Nz(Err.Number)
        Debug.Print "Error " & Nz(Err.Number) & ":" & Err.Description & " Note: Ensure SQLite ODBC driver is installed."
    End If
    Resume ExitMethod
End Function



Private Sub btnClose_Click()
On Error Resume Next
    DoCmd.Close
End Sub



Private Sub btnConnect_Click()
On Error GoTo HandleErrors

    Dim i As Long
    Dim iSuccess As Long
    Dim iFailed As Long
    
    If (Len(Nz(Me.txtDbPath, "")) < 1) Then
        Me.txtDbPath = BrowseToSqliteFile
        If (Len(Me.txtDbPath) < 1) Then
            MsgBox "You did not select a SQLite database file!", vbExclamation, "Database Not Selected"
            GoTo ExitMethod
        End If
    End If
    
    iSuccess = 0
    iFailed = 0
    For i = 0 To Me.lstTablesToAdd.ListCount - 1
        If LinkSqliteTable(Me.txtDbPath, Me.lstTablesToAdd.ItemData(i)) Then
            iSuccess = iSuccess + 1
        Else
            iFailed = iFailed + 1
        End If
    Next

    If iFailed < 1 Then
        MsgBox iSuccess & " SQLite tables linked successfully!", vbInformation, "Linked Tables Created"
    Else
        If iSuccess < 1 Then
            MsgBox iFailed & " SQLite tables FAILED to link!", vbCritical, "Failed To Link"
        Else
            MsgBox iSuccess & " SQLite tables linked successfully, " & iFailed & " FAILED to link!", vbExclamation, "Some Links Failed"
        End If
    End If

ExitMethod:
    Exit Sub
HandleErrors:
    MsgBox "Error: " & Err.Description, vbCritical, "Error " & Nz(Err.Number)
    Resume ExitMethod
End Sub



Private Sub FormHeader_Click()
On Error GoTo HandleErrors

    DoCmd.OpenForm "frmAbout"
    
ExitMethod:
    Exit Sub
HandleErrors:
    MsgBox "Error: " & Err.Description, vbCritical, "Error " & Nz(Err.Number)
    Resume ExitMethod
End Sub



Private Sub txtLink_Click()
On Error GoTo HandleErrors

    Dim oShell As Object
    Set oShell = CreateObject("Shell.Application")
    oShell.Open CVar("http://" & Me.txtLink)
    
ExitMethod:
    Exit Sub
HandleErrors:
    MsgBox "Error: " & Err.Description, vbCritical, "Error " & Nz(Err.Number)
    Resume ExitMethod
End Sub



Private Sub txtLinkx64Driver_Click()
On Error GoTo HandleErrors

    Dim oShell As Object
    Set oShell = CreateObject("Shell.Application")
    oShell.Open CVar("http://www.ch-werner.de/sqliteodbc/sqliteodbc_w64.exe")
    
ExitMethod:
    Exit Sub
HandleErrors:
    MsgBox "Error: " & Err.Description, vbCritical, "Error " & Nz(Err.Number)
    Resume ExitMethod
End Sub



Private Sub txtLinkx86Driver_Click()
On Error GoTo HandleErrors

    Dim oShell As Object
    Set oShell = CreateObject("Shell.Application")
    oShell.Open CVar("http://www.ch-werner.de/sqliteodbc/sqliteodbc.exe")
    
ExitMethod:
    Exit Sub
HandleErrors:
    MsgBox "Error: " & Err.Description, vbCritical, "Error " & Nz(Err.Number)
    Resume ExitMethod
End Sub

