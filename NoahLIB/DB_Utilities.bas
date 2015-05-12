Attribute VB_Name = "DB_Utilities"
' version: 1.0

Option Explicit

' connMyUtil e rsUtil sono gli oggetti "di default" per comunicare col DB MySQL

Public connMyUtil As ADODB.Connection
Public rsUtil As ADODB.Recordset


Public tryReset As Integer
Public Const MAX_RESET = 2
Public Const MAX_RESET_2 = 4
'
'




' resets conn.. and writes to log
Public Sub resetConn()
    Call logToFile("reset Conn")
    On Error Resume Next
    connMyUtil.Close
    On Error GoTo 0
    Set connMyUtil = Nothing
    
    If tryReset >= MAX_RESET_2 Then
        Call MsgBox("[DB_Utilities:resetConn] Unable to reach the server, exit program.", vbCritical)
        End
    End If
    
End Sub



' check that default MySQL connection is open, otherwise opens it
Public Sub checkConn()

    If connMyUtil Is Nothing Then
        Set connMyUtil = getMySQLConnection
    
    ElseIf connMyUtil.State <> adStateOpen Then
        
        On Error Resume Next
        connMyUtil.Close
        On Error GoTo 0
        Set connMyUtil = getMySQLConnection
        
    Else
    'Else
    End If
    
    If rsUtil Is Nothing Then
        Set rsUtil = New ADODB.Recordset
    End If

End Sub



' executes a scalar SQL query and returns the output
Public Function execScalarSQL(ByVal strSQL As String)
    Dim i As Integer
    Dim errMessage As String
    Dim errMessageLog As String

    If Not (rsUtil Is Nothing) Then
        If rsUtil.State = 1 Then rsUtil.Close
    End If
    Set rsUtil = Nothing
    Call checkConn
    
    On Error GoTo errSQL
    rsUtil.Open strSQL, connMyUtil, adOpenDynamic, adLockOptimistic
    ' worked..
    tryReset = 0
    
    On Error GoTo errMove
    rsUtil.MoveFirst
    execScalarSQL = rsUtil(0)
    Set rsUtil = Nothing
    
    Exit Function
    
errMove:
        On Error GoTo 0
'        execScalarSQL = -1
        execScalarSQL = Nothing
        Set rsUtil = Nothing
        Exit Function
    
errSQL:
        On Error GoTo 0
        
        ' RETRY 5 times
        If tryReset <= MAX_RESET Then
            tryReset = tryReset + 1
            Call resetConn
            execScalarSQL = execScalarSQL(strSQL)
            Exit Function
        End If
    
        ' if 5 retries don't work, then print the errors
        For i = 0 To connMyUtil.Errors.count - 1
            errMessage = "SQL error in query:" & vbCrLf & vbCrLf & _
                         strSQL & vbCrLf & vbCrLf & _
                         "Error message: " & vbCrLf & vbCrLf & _
                        connMyUtil.Errors(i)
            errMessageLog = vbTab & "SQL error:" & vbTab & _
                         strSQL & vbTab & _
                         "Error Message: " & vbTab & _
                        connMyUtil.Errors(i)
                        
            MsgBox (errMessage)
            logToFile (errMessageLog)
        Next
    
        'reset the connection
        Set connMyUtil = Nothing

End Function


' executes a SQL query
Public Function execCommandSQL(ByVal strSQL As String, Optional ByRef error)
    Dim i As Integer
    Dim errMessage As String
    Dim errMessageLog As String

    
    If Not (rsUtil Is Nothing) Then
        If rsUtil.State = 1 Then rsUtil.Close
    End If
    Set rsUtil = Nothing
    Call checkConn
    
    On Error GoTo errSQL
    execCommandSQL = connMyUtil.Execute(strSQL)
    ' worked..
    tryReset = 0
    
    Exit Function
    
errSQL:
        On Error GoTo 0
        
        ' RETRY 5 times
        If tryReset <= MAX_RESET Then
            tryReset = tryReset + 1
            Call resetConn
            execCommandSQL = execCommandSQL(strSQL, error)
            Exit Function
        End If
        
        
        For i = 0 To connMyUtil.Errors.count - 1
            errMessage = "SQL error in query:" & vbCrLf & vbCrLf & _
                         strSQL & vbCrLf & vbCrLf & _
                         "Error message: " & vbCrLf & vbCrLf & _
                        connMyUtil.Errors(i)
            errMessageLog = vbTab & "SQL error:" & vbTab & _
                         strSQL & vbTab & _
                         "Error Message: " & vbTab & _
                        connMyUtil.Errors(i)
                        
            MsgBox (errMessage)
            logToFile (errMessageLog)
        Next
        error = True
        
        'reset the connection
        Set connMyUtil = Nothing
        

End Function


'esegue la query SQL, e restituisce un recordset

Public Sub execTableSQL(ByVal strSQL As String)
    Dim i As Integer
    Dim errMessage As String
    Dim errMessageLog As String

    
    If Not (rsUtil Is Nothing) Then
        If rsUtil.State = 1 Then rsUtil.Close
    End If
    Set rsUtil = Nothing
    Call checkConn
    
    On Error GoTo errSQL
    rsUtil.Open strSQL, connMyUtil, adOpenDynamic, adLockOptimistic
    
    On Error GoTo errGoto
    rsUtil.MoveFirst
    ' worked
    Exit Sub
    
errGoto:
    On Error GoTo 0
    Set rsUtil = Nothing
    Exit Sub
    
errSQL:
        On Error GoTo 0
        For i = 0 To connMyUtil.Errors.count - 1
            errMessage = "SQL error in query:" & vbCrLf & vbCrLf & _
                         strSQL & vbCrLf & vbCrLf & _
                         "Error message: " & vbCrLf & vbCrLf & _
                        connMyUtil.Errors(i)
            errMessageLog = vbTab & "SQL error:" & vbTab & _
                         strSQL & vbTab & _
                         "Error Message: " & vbTab & _
                        connMyUtil.Errors(i)
                        
            MsgBox (errMessage)
            logToFile (errMessageLog)
        Next
        
        'reset the connection
        Set connMyUtil = Nothing

End Sub


' DB 2015-02-10 execScalarSQL_WithConn is not used anywhere, commented -
'Public Function execScalarSQL_WithConn(ByVal strSQL As String, ByRef conn As ADODB.Connection)
'    Dim i As Integer
'    Dim errMessage As String
'    Dim errMessageLog As String
'    If Not (rsUtil Is Nothing) Then
'        If rsUtil.State = 1 Then rsUtil.Close
'    End If
'    Set rsUtil = Nothing
'    Call checkConn
'
'    On Error GoTo errSQL
'    rsUtil.Open strSQL, conn, adOpenDynamic, adLockOptimistic
    
'    On Error GoTo errMove
'    rsUtil.MoveFirst
'    execScalarSQL = rsUtil(0)
'    Set rsUtil = Nothing
'
'    Exit Function
'errMove:
'        On Error GoTo 0
'        execScalarSQL = -1
'        Set rsUtil = Nothing
'        Exit Function
'
'errSQL:
'
'        On Error GoTo 0
'        For i = 0 To connMyUtil.Errors.count - 1
'            errMessage = "SQL error in query:" & vbCrLf & vbCrLf & _
'                         strSQL & vbCrLf & vbCrLf & _
'                         "Error message: " & vbCrLf & vbCrLf & _
'                        connMyUtil.Errors(i)
'            errMessageLog = vbTab & "SQL error:" & vbTab & _
'                         strSQL & vbTab & _
'                         "Error Message: " & vbTab & _
'                        connMyUtil.Errors(i)
'
'            MsgBox (errMessage)
'            logToFile (errMessageLog) '
'        Next
'
'        'reset the connection
'        Set connMyUtil = Nothing'
'
'
'End Function


'****************************************************************************************************************
' STORES IN rsUtil
'****************************************************************************************************************
Public Sub execTableSQL_withConn(ByVal strSQL As String, ByRef conn As ADODB.Connection)
    Dim i As Integer
    Dim errMessage As String
    Dim errMessageLog As String

    
    If Not (rsUtil Is Nothing) Then
        If rsUtil.State = 1 Then rsUtil.Close
    End If
    Set rsUtil = Nothing
    Call checkConn
    
    On Error GoTo errSQL
    rsUtil.Open strSQL, conn
    On Error GoTo errGoto
    rsUtil.MoveFirst
    
    Exit Sub
    
errGoto:
    On Error GoTo 0
    Exit Sub
        
errSQL:
    On Error GoTo 0
    For i = 0 To connMyUtil.Errors.count - 1
        errMessage = "SQL error in query:" & vbCrLf & vbCrLf & _
                     strSQL & vbCrLf & vbCrLf & _
                     "Error message: " & vbCrLf & vbCrLf & _
                    connMyUtil.Errors(i)
        errMessageLog = vbTab & "SQL error:" & vbTab & _
                     strSQL & vbTab & _
                     "Error Message: " & vbTab & _
                    connMyUtil.Errors(i)
                    
        MsgBox (errMessage)
        logToFile (errMessageLog)
    Next

    
End Sub


' nota: non usa rsUtil
' non usa connMyUtil
Public Sub execTableSQL_withConn2(ByVal strSQL As String, ByRef conn As ADODB.Connection, ByRef rs As ADODB.Recordset)
    Dim i As Integer
    Dim errMessage As String
    Dim errMessageLog As String
    
    If Not (rs Is Nothing) Then
        If rs.State = 1 Then rs.Close
    End If
    Set rs = Nothing
    
    Set rs = New ADODB.Recordset
    
    On Error GoTo errSQL
    rs.Open strSQL, conn
    
    On Error GoTo errMove
    rs.MoveFirst
    
    Exit Sub
    
errMove:
    On Error GoTo 0
    Set rs = Nothing
    Exit Sub

errSQL:
    On Error GoTo 0
    For i = 0 To conn.Errors.count - 1
        errMessage = "SQL error in query:" & vbCrLf & vbCrLf & _
                     strSQL & vbCrLf & vbCrLf & _
                     "Error message: " & vbCrLf & vbCrLf & _
                    conn.Errors(i)
        errMessageLog = vbTab & "SQL error:" & vbTab & _
                     strSQL & vbTab & _
                     "Error Message: " & vbTab & _
                    conn.Errors(i)
                    
        MsgBox (errMessage)
        logToFile (errMessageLog)
    Next

    
End Sub


' nota: stores the recordset in rs
' usa la connessione connMyUtil

Public Sub execTableSQL_withRS(ByVal strSQL As String, ByRef rs As ADODB.Recordset)
    Dim i As Integer
    Dim errMessage As String
    Dim errMessageLog As String

    Call checkConn
    
    If Not (rs Is Nothing) Then
        If rs.State = 1 Then rs.Close
    End If
    Set rs = Nothing
    
    Set rs = New ADODB.Recordset
    
    On Error GoTo errSQL
    rs.Open strSQL, connMyUtil
    tryReset = 0
    
    On Error GoTo errMove
    rs.MoveFirst
    
    Exit Sub

errMove:
    ' il recordset è vuoto
    On Error GoTo 0
    Set rs = Nothing
    Exit Sub

errSQL:
    
    ' RETRY 5 times
    If tryReset <= MAX_RESET Then
        tryReset = tryReset + 1
        Call resetConn
        Call execTableSQL_withRS(strSQL, rs)
        Exit Sub
    End If


    On Error GoTo 0
    For i = 0 To connMyUtil.Errors.count - 1
        errMessage = "SQL error in query:" & vbCrLf & vbCrLf & _
                     strSQL & vbCrLf & vbCrLf & _
                     "Error message: " & vbCrLf & vbCrLf & _
                    connMyUtil.Errors(i)
        errMessageLog = vbTab & "SQL error:" & vbTab & _
                     strSQL & vbTab & _
                     "Error Message: " & vbTab & _
                    connMyUtil.Errors(i)

        MsgBox (errMessage)
        logToFile (errMessageLog)
    Next
    
    'reset the connection
    Set connMyUtil = Nothing
    
End Sub

' gets the last result row count
Public Function getRowsFoundCount() As Long
    getRowsFoundCount = DB_Utilities.execScalarSQL("select found_rows();")
End Function


' creates the default connection to the MySQL server, and opens it
Public Function getMySQLConnection()

    Dim vError As Variant
    Dim sErrors As String
        
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    cn.ConnectionString = "DRIVER={MySQL ODBC 5.2 Unicode Driver};" & _
                          "SERVER=wsa-03;DATABASE=lossdb;UID=root;PWD=pass; OPTION=3"
    
    On Error Resume Next
    cn.Open
    On Error GoTo 0
    
    If cn.State = adStateOpen Then
        'MsgBox "Connection Succeeded"
    Else
        For Each vError In cn.Errors
            sErrors = sErrors & vError.Description & vbNewLine
        Next vError
        If sErrors <> "" Then
            MsgBox sErrors, vbExclamation
            Call logToFile(sErrors)
        Else
            MsgBox "Connection Failed", vbExclamation
            Call logToFile("Connection Failed")
        End If
            
        If (InStr(sErrors, "Can't connect") > 0) Or (InStr(sErrors, "Lost connection") > 0) Then
            Call MsgBox("Unable to connect to the LOSSDB server (server is down?), exit now.", vbCritical, "Error")
            End
        End If
        If InStr(sErrors, "MySQL") > 0 Then
            Call MsgBox("Unable to connect to the LOSSDB server (wrong username?), exit now.", vbCritical, "Error")
            End
        End If
        
        If MsgBox("Please install MySQL Connector 5.2 ." & vbCrLf & "Install it now?", vbYesNo) = vbYes Then
            Call Shell("explorer.exe ""K:\_SOFTWARES\INSTALLAZIONI\MySql\Installazioni\mysql-connector-odbc-5.2.6-win32.msi""", vbNormalFocus)
            End
    '        lRet = Shell("K:\_SOFTWARES\INSTALLAZIONI\MySql\Installazioni\mysql-connector-odbc-5.2.6-win32.msi", vbNormalFocus)
    '        If lRet = 0 Then MsgBox "Could start....check File path!", vbSystemModal
        End If
        
            
End If

Set getMySQLConnection = cn

End Function



'**********************************************************************
' connects to MS SQL using SQL autentication and gives back the connection
'
Public Function getSQLServerConnectionCatrader()
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=SQLOLEDB;Data Source=wsa-12\sql12;" & _
                      "Initial Catalog=AirCT2Exp;User Id=catrader;Password=catrader;"
    conn.Open
    
    Set getSQLServerConnectionCatrader = conn
End Function

'**********************************************************************
' connects to MS SQL using Windows autentication and gives back the connection
'
Public Function getSQLServerConnectionSSPI(ByVal servername As String)
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=SQLOLEDB;Data Source=" & servername & ";" & _
                      "Integrated Security=SSPI;"
    conn.Open
    
    Set getSQLServerConnectionSSPI = conn
End Function


Public Function getSQLServerConnection(servername As String)
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=SQLOLEDB;Data Source=" & servername & ";" & _
                      "Integrated Security=SSPI;"
    conn.Open
    
    Set getSQLServerConnection = conn
End Function


' logs to MySQL DB
Public Sub LogToDb(ByRef conn As ADODB.Connection, ByVal txt As String)
    'MsgBox "CALL lossdb.insertlog ('" & esc1(txt) & "');"
    'asd = "CALL lossdb.insertlog ('" & esc1(txt) & "');"
    'Cells(1, 10) = "CALL lossdb.insertlog ('" & esc1(txt) & "');"
    conn.Execute "CALL lossdb.insertlog ('" & esc1(txt) & "');"
End Sub

' escapes ' characters
Function esc1(ByVal txt As String)
    Dim val1 As String
    
    val1 = Trim(Replace(txt, "\", "\\"))
    esc1 = Trim(Replace(val1, "'", "\'"))
End Function














