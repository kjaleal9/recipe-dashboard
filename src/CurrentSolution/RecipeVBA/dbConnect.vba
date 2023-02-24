'////////////////////////////////////////
'
'dbConnect Class Version 1.3.0
'v6.3.00_01
'Created by: Carey Warren
'Modified by: Carey Warren
'Date: 2015-10-02

Private cn As adodb.Connection
Private rs As adodb.Recordset
Private comm As adodb.Command
Private par As adodb.Parameter
Private tempServerName As String
Private tempDatabaseName As String
Private tempUserName As String
Private tempPassword As String
Private tempDSN As String
Private db As String

Public Sub Connect()
On Error GoTo ErrHandler
    Set cn = New adodb.Connection
    cn.CursorLocation = adUseClient
    cn.CommandTimeout = 10
    cn.ConnectionTimeout = 20
    cn.Mode = adModeReadWrite
    cn.Open "Driver={SQL Server};Server=" & ServerName & ";DATABASE=" & DatabaseName & ";UID=" & Username & ";pwd=" & Password & ";"
   
Exit Sub

'Error message
ErrHandler:
    
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on dbConnect Connect", ftDiagSeverityError

End Sub

'1.3///////////////
'added DSN connection
Public Sub ConnectDSN()
On Error GoTo ErrHandler
    Set cn = New adodb.Connection
    cn.CursorLocation = adUseClient
    cn.CommandTimeout = 20
    cn.ConnectionTimeout = 10
    cn.Mode = adModeReadWrite
    cn.Open DSN
   
Exit Sub

'Error message
ErrHandler:
    
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on dbConnect Connect", ftDiagSeverityError

End Sub


Public Sub Disconnect()
On Error GoTo ErrHandler
    cn.Close
    Set cn = Nothing
Exit Sub

'Error message
ErrHandler:
    
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on dbConnect Disconnect", ftDiagSeverityError

End Sub

Private Sub Class_Terminate()
    'Destroy objects
    Set rs = Nothing
    Set comm = Nothing
    Set par = Nothing
    Set cn = Nothing
End Sub

'used to execute an insert or delete sql statement from outside the class
Public Sub execute(sqlstring)
On Error GoTo ErrHandler
    If cn.State = adStateOpen Then
        cn.execute (sqlstring)
    End If
Exit Sub
    'Error message
ErrHandler:
    
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on dbConnect Execute", ftDiagSeverityError
End Sub

'used to return a recordset outside the class
Public Function getRecords(sqlstring) As adodb.Recordset
On Error GoTo ErrHandler
    If cn.State = adStateOpen Then
        Set rs = cn.execute(sqlstring)
        Set getRecords = rs.Clone
        Set rs = Nothing
    End If
Exit Function

'Error message
ErrHandler:
    
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on dbConnect getRecords", ftDiagSeverityError
End Function

'run a command from a display
Public Function executeCommand(storedprocedure As String) As adodb.Recordset
On Error GoTo ErrHandler
    If cn.State = adStateOpen Then
        'Setup Command to SQL database
        Set comm = New adodb.Command
        Set comm.ActiveConnection = cn
        comm.CommandText = storedprocedure          'Name of Stored Procedure
        comm.CommandType = adCmdStoredProc          'Command type
        
        'Execute setup Command with Parameters
        Set executeCommand = comm.execute()           'Execute command
    End If
Exit Function

'Error message
ErrHandler:
    
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on dbConnect executeCommand", ftDiagSeverityError
End Function

'run a command from a display with 1 parameter
Public Function executeCommand1Par(storedprocedure As String, par As adodb.Parameter) As adodb.Recordset
On Error GoTo ErrHandler
    If cn.State = adStateOpen Then
        'Setup Command to SQL database
        Set comm = New adodb.Command
        Set comm.ActiveConnection = cn
        comm.CommandText = storedprocedure          'Name of Stored Procedure
        comm.CommandType = adCmdStoredProc          'Command type
        comm.parameters.Append par
        
        
        'Execute setup Command with Parameters
        Set executeCommand1Par = comm.execute()           'Execute command
    End If
Exit Function

'Error message
ErrHandler:
    
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on dbConnect executeCommand1Par", ftDiagSeverityError
End Function

'run a command from a display with 2 parameter
Public Function executeCommand2Par(storedprocedure As String, par1 As adodb.Parameter, par2 As adodb.Parameter) As adodb.Recordset
On Error GoTo ErrHandler
    If cn.State = adStateOpen Then
        'Setup Command to SQL database
        Set comm = New adodb.Command
        Set comm.ActiveConnection = cn
        comm.CommandText = storedprocedure          'Name of Stored Procedure
        comm.CommandType = adCmdStoredProc          'Command type
        comm.parameters.Append par1
        comm.parameters.Append par2
        
        'Execute setup Command with Parameters
        Set executeCommand2Par = comm.execute()           'Execute command
    End If
Exit Function

'Error message
ErrHandler:
    
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on dbConnect executeCommand2Par", ftDiagSeverityError
End Function
'run a command from a display with 3 parameter
Public Function executeCommand3Par(storedprocedure As String, par1 As adodb.Parameter, par2 As adodb.Parameter, par3 As adodb.Parameter) As adodb.Recordset
On Error GoTo ErrHandler
    If cn.State = adStateOpen Then
        'Setup Command to SQL database
        Set comm = New adodb.Command
        Set comm.ActiveConnection = cn
        comm.CommandText = storedprocedure          'Name of Stored Procedure
        comm.CommandType = adCmdStoredProc          'Command type
        comm.parameters.Append par1
        comm.parameters.Append par2
        comm.parameters.Append par3
        
        'Execute setup Command with Parameters
        Set executeCommand3Par = comm.execute()           'Execute command
    End If
Exit Function

'Error message
ErrHandler:
    
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on dbConnect executeCommand3Par", ftDiagSeverityError
End Function

'run a command from a display with 4 parameter
Public Function executeCommand4Par(storedprocedure As String, par1 As adodb.Parameter, par2 As adodb.Parameter, par3 As adodb.Parameter, par4 As adodb.Parameter) As adodb.Recordset
On Error GoTo ErrHandler
    If cn.State = adStateOpen Then
        'Setup Command to SQL database
        Set comm = New adodb.Command
        Set comm.ActiveConnection = cn
        comm.CommandText = storedprocedure          'Name of Stored Procedure
        comm.CommandType = adCmdStoredProc          'Command type
        comm.parameters.Append par1
        comm.parameters.Append par2
        comm.parameters.Append par3
        comm.parameters.Append par4
        
        'Execute setup Command with Parameters
        Set executeCommand4Par = comm.execute()           'Execute command
    End If
Exit Function

'Error message
ErrHandler:
    
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on dbConnect executeCommand5Par", ftDiagSeverityError
End Function

'run a command from a display with 5 parameter
Public Function executeCommand5Par(storedprocedure As String, par1 As adodb.Parameter, par2 As adodb.Parameter, par3 As adodb.Parameter, par4 As adodb.Parameter, par5 As adodb.Parameter) As adodb.Recordset
On Error GoTo ErrHandler
    If cn.State = adStateOpen Then
        'Setup Command to SQL database
        Set comm = New adodb.Command
        Set comm.ActiveConnection = cn
        comm.CommandText = storedprocedure          'Name of Stored Procedure
        comm.CommandType = adCmdStoredProc          'Command type
        comm.parameters.Append par1
        comm.parameters.Append par2
        comm.parameters.Append par3
        comm.parameters.Append par4
        comm.parameters.Append par5
        
        'Execute setup Command with Parameters
        Set executeCommand5Par = comm.execute()           'Execute command
    End If
Exit Function

'Error message
ErrHandler:
    
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on dbConnect executeCommand5Par", ftDiagSeverityError
End Function


'return connection
Public Function getConnection() As adodb.Connection
    'Setup Command to SQL database
    Set getConnection = cn
End Function

Public Property Get ServerName() As String
    ServerName = tempServerName
End Property

Public Property Let ServerName(ByVal Name As String)
    tempServerName = Name
End Property

Public Property Get DatabaseName() As String
    DatabaseName = tempDatabaseName
End Property

Public Property Let DatabaseName(ByVal Name As String)
    tempDatabaseName = Name
End Property

Public Property Get Username() As String
    Username = tempUserName
End Property

Public Property Let Username(ByVal Name As String)
    tempUserName = Name
End Property

Public Property Get Password() As String
    Password = tempPassword
End Property

Public Property Let Password(ByVal value As String)
    tempPassword = value
End Property
'1.3///////////////
Public Property Get DSN() As String
    DSN = tempDSN
End Property
'1.3///////////////
Public Property Let DSN(ByVal value As String)
    tempDSN = value
End Property

'run a command from a display with 10 parameter
Public Function executeCommand10Par(storedprocedure As String, par1 As adodb.Parameter, par2 As adodb.Parameter, par3 As adodb.Parameter, par4 As adodb.Parameter, par5 As adodb.Parameter, par6 As adodb.Parameter, par7 As adodb.Parameter, par8 As adodb.Parameter, par9 As adodb.Parameter, par10 As adodb.Parameter) As adodb.Recordset
On Error GoTo ErrHandler
    If cn.State = adStateOpen Then
        'Setup Command to SQL database
        Set comm = New adodb.Command
        Set comm.ActiveConnection = cn
        comm.CommandText = storedprocedure          'Name of Stored Procedure
        comm.CommandType = adCmdStoredProc          'Command type
        comm.parameters.Append par1
        comm.parameters.Append par2
        comm.parameters.Append par3
        comm.parameters.Append par4
        comm.parameters.Append par5
        comm.parameters.Append par6
        comm.parameters.Append par7
        comm.parameters.Append par8
        comm.parameters.Append par9
        comm.parameters.Append par10
        
        'Execute setup Command with Parameters
        Set executeCommand10Par = comm.execute()           'Execute command
    End If
Exit Function

'Error message
ErrHandler:
    
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on dbConnect executeCommand10Par", ftDiagSeverityError
End Function

'run a command from a display with 7 parameter
Public Function executeCommand7Par(storedprocedure As String, par1 As adodb.Parameter, par2 As adodb.Parameter, par3 As adodb.Parameter, par4 As adodb.Parameter, par5 As adodb.Parameter, par6 As adodb.Parameter, par7 As adodb.Parameter) As adodb.Recordset
On Error GoTo ErrHandler
    If cn.State = adStateOpen Then
        'Setup Command to SQL database
        Set comm = New adodb.Command
        Set comm.ActiveConnection = cn
        comm.CommandText = storedprocedure          'Name of Stored Procedure
        comm.CommandType = adCmdStoredProc          'Command type
        comm.parameters.Append par1
        comm.parameters.Append par2
        comm.parameters.Append par3
        comm.parameters.Append par4
        comm.parameters.Append par5
        comm.parameters.Append par6
        comm.parameters.Append par7
        
        'Execute setup Command with Parameters
        Set executeCommand7Par = comm.execute()           'Execute command
    End If
Exit Function

'Error message
ErrHandler:
    
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on dbConnect executeCommand7Par", ftDiagSeverityError
End Function

'run a command from a display with 11 parameter
Public Function executeCommand11Par(storedprocedure As String, par1 As adodb.Parameter, par2 As adodb.Parameter, par3 As adodb.Parameter, par4 As adodb.Parameter, par5 As adodb.Parameter, par6 As adodb.Parameter, par7 As adodb.Parameter, par8 As adodb.Parameter, par9 As adodb.Parameter, par10 As adodb.Parameter, par11 As adodb.Parameter) As adodb.Recordset
On Error GoTo ErrHandler
    If cn.State = adStateOpen Then
        'Setup Command to SQL database
        Set comm = New adodb.Command
        Set comm.ActiveConnection = cn
        comm.CommandText = storedprocedure          'Name of Stored Procedure
        comm.CommandType = adCmdStoredProc          'Command type
        comm.parameters.Append par1
        comm.parameters.Append par2
        comm.parameters.Append par3
        comm.parameters.Append par4
        comm.parameters.Append par5
        comm.parameters.Append par6
        comm.parameters.Append par7
        comm.parameters.Append par8
        comm.parameters.Append par9
        comm.parameters.Append par10
        comm.parameters.Append par11
        
        'Execute setup Command with Parameters
        Set executeCommand11Par = comm.execute()           'Execute command
    End If
Exit Function

'Error message
ErrHandler:
    
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on dbConnect executeCommand11Par", ftDiagSeverityError
End Function

