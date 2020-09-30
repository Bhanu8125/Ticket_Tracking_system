Attribute VB_Name = "CloseTicketRepository"
Dim Employees As New Collection
Dim TicketId As New Collection
Public Message As String
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
    
    sqlst = "select EmployeeName from Employee where Dept = 'Devops';"
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
Public Function GetTicketId() As Collection
    Dim Isget As Boolean
    Isget = False
    TicketIdFromDatabase
    Isget = True
    Set GetTicketId = TicketId
End Function
Private Sub TicketIdFromDatabase()
      On Error GoTo errHand
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    
    'Helper
    Dim sqlst As String
    Dim ConString As String
    
    sqlst = "select Ticket_Id from Ticket where status = 'open';"
    ConString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=db_tickettracking;Data Source=."
    'to open connection
    con.Open ConString
    'execute command
    Set cmd.ActiveConnection = con
    cmd.CommandText = sqlst
    Set rs = cmd.Execute
    Set TicketId = New Collection
    While Not rs.EOF
        TicketId.Add CStr(rs(0))
        rs.MoveNext
    Wend
    rs.Close
    con.Close
    Exit Sub
errHand:
    MsgBox Err.Description
    'Err.Raise 1001, , "Error in getting Data , check logo file foe more details"
End Sub
Public Function UpdateTicket(ByVal TicketId As Integer, ByVal EmployeeName As String, ByVal Resolution As String) As Boolean
     Dim Isget As Boolean
    Isget = False
    UpdateTicketInDB TicketId, EmployeeName, Resolution
    Isget = True
    UpdateTicket = Isget
End Function
Private Sub UpdateTicketInDB(ByVal TicketId As Integer, ByVal EmployeeName As String, ByVal Resolution As String)
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
         cmd.Parameters.Append cmd.CreateParameter("@TicketId", adInteger, adParamInput, , TicketId)
         cmd.Parameters.Append cmd.CreateParameter("@Employee", adVarChar, adParamInput, 30, EmployeeName)
         cmd.Parameters.Append cmd.CreateParameter("@Resolution", adVarChar, adParamInput, 10, Resolution)
         cmd.Parameters.Append cmd.CreateParameter("result", adInteger, adParamOutput)
        
        sqlst = "sp_CloseTicket"
        cmd.CommandText = sqlst
        cmd.CommandType = adCmdStoredProc
        cmd.Execute
        If cmd("result") Then
            Message = "Ticket  " & TicketId & " is Closed"
        End If
    con.Close
    Exit Sub
errHand:
    MsgBox Err.Description
    'Err.Raise 1001, , "Error in getting Data , check logo file foe more details"
End Sub
