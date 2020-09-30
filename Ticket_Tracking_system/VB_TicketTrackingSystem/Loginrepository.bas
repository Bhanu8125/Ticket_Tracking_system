Attribute VB_Name = "Loginrepository"
Dim Depts As New Collection
Public Function GetDept() As Collection
    Dim Isget As Boolean
    Isget = False
    DeptFromDatabase
    Isget = True
    Set GetDept = Depts
End Function
Private Sub DeptFromDatabase()
      On Error GoTo errhand
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    
    'Helper
    Dim sqlst As String
    Dim ConString As String
    
    sqlst = "select Distinct(Dept) from Employee"
    ConString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=db_tickettracking;Data Source=."
    'to open connection
    con.Open ConString
    'execute command
    Set cmd.ActiveConnection = con
    cmd.CommandText = sqlst
    Set rs = cmd.Execute
    Set Depts = New Collection
    While Not rs.EOF
        Depts.Add CStr(rs(0))
        rs.MoveNext
    Wend
    rs.Close
    con.Close
    Exit Sub
errhand:
    MsgBox Err.Description
    'Err.Raise 1001, , "Error in getting Data , check logo file foe more details"
End Sub
