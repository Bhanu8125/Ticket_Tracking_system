Attribute VB_Name = "CreateTicketRepository"
Dim Employees As New Collection
Public Function GetEmployees() As Collection
    Dim Isget As Boolean
    Isget = False
    EmpFromDatabase
    Isget = True
    Set GetEmployees = Employees
End Function
Private Sub EmpFromDatabase()
      On Error GoTo errHand
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    
    'Helper
    Dim sqlst As String
    Dim ConString As String
    
    sqlst = "select EmployeeName from Employee where Dept <> 'Devops';"
    ConString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=db_tickettracking;Data Source=."
    'to open connection
    con.Open ConString
    'execute command
    Set cmd.ActiveConnection = con
    cmd.CommandText = sqlst
    Set rs = cmd.Execute
    Set Employees = New Collection
    While Not rs.EOF
        Employees.Add CStr(rs(0))
        rs.MoveNext
    Wend
    rs.Close
    con.Close
    Exit Sub
errHand:
    Err.Raise 1001, , "Error in getting Data , check logo file foe more details"
End Sub

Public Function insert(ByVal Ename As String, ByVal Dated As String, ByVal Severity As String, ByVal Description As String) As Boolean
    Dim Isinsert As Boolean
    Isinsert = False
    CreateTicketinDB Ename, Dated, Severity, Description
    Isinsert = True
     insert = Isinsert
End Function
Private Sub CreateTicketinDB(ByVal Ename As String, ByVal Dated As String, ByVal Severity As String, ByVal Description As String)
    On Error GoTo errHand
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    
    'Helper
    Dim sqlst As String
    Dim ConString As String
    
    ConString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=db_tickettracking;Data Source=."
    'to open connection
    con.Open ConString
    Set cmd.ActiveConnection = con
    'execute command
         cmd.Parameters.Append cmd.CreateParameter("@Eid", adVarChar, adParamInput, 7, Ename)
         cmd.Parameters.Append cmd.CreateParameter("@date", adVarChar, adParamInput, 30, Dated)
         cmd.Parameters.Append cmd.CreateParameter("@severity", adVarChar, adParamInput, 10, Severity)
         cmd.Parameters.Append cmd.CreateParameter("@Desc", adVarChar, adParamInput, 30, Description)
        sqlst = "sp_CreateTicket"
        cmd.CommandText = sqlst
        cmd.CommandType = adCmdStoredProc
        cmd.Execute
    con.Close
    Exit Sub
errHand:
    MsgBox Err.Description
    'Err.Raise 1001, , "Error in getting Data , check logo file foe more details"
End Sub

