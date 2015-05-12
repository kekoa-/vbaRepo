Attribute VB_Name = "modDBInterface"
Option Explicit

Public Sub startTransaction()
    Call DB_Utilities.execCommandSQL("START TRANSACTION;")
End Sub

Public Sub commitTransaction()
    Call DB_Utilities.execCommandSQL("COMMIT;")
End Sub

Public Sub rollbackTransaction()
    Call DB_Utilities.execCommandSQL("ROLLBACK;")
End Sub

Public Function getVar(ByVal varName As String)
    getVar = DB_Utilities.execScalarSQL("select " & varName & " as varToVBA;")
End Function


Public Function getLastInsertID_num() As Long
    getLastInsertID_num = DB_Utilities.execScalarSQL("select last_insert_id();")
End Function

Public Sub insertStringKey(ByVal tableName As String, ByVal keyName As String, ByVal keyValue As String)
        Dim strSQL As String
        strSQL = "INSERT INTO " & tableName & "(" & keyName & ") VALUES ('" & keyValue & "');"
        Call execCommandSQL(strSQL)
End Sub

Public Sub insertNumKey(ByVal tableName As String, ByVal keyName As String, ByVal keyValue As Double)
        Dim strSQL As String
        strSQL = "INSERT INTO " & tableName & "(" & keyName & ") VALUES (" & keyValue & ");"
        Call execCommandSQL(strSQL)
End Sub


Public Sub updateStringValueStringKey(ByVal tableName As String, _
        ByVal colName As String, ByVal inputVal As String, _
        ByVal keyName As String, ByVal keyValue As String)

    Dim strSQL As String
    
    If checkStringKeyExists(tableName, keyName, keyValue) = False Then
        MsgBox ("[modDBInterface:updateStringValueStringKey] Warning: la chiave <" & keyValue & "> non e' presente nella tabella " & _
        tableName & "(" & keyName & ")")
    End If
    
    Dim cmdPrep1 As ADODB.command
    Set cmdPrep1 = New ADODB.command
    Dim prm1 As ADODB.Parameter
    Dim prm2 As ADODB.Parameter
    
    Set cmdPrep1.ActiveConnection = connMyUtil
    cmdPrep1.CommandType = adCmdText
    cmdPrep1.CommandText = "UPDATE " & tableName & " SET " & colName & "=? WHERE " & keyName & "=?;"
        
    Set prm1 = cmdPrep1.CreateParameter("datavalue", ADODB.adVarWChar, ADODB.adParamInput, 500)
    cmdPrep1.Parameters.Append prm1
    Set prm2 = cmdPrep1.CreateParameter("keyvalue", ADODB.adVarWChar, ADODB.adParamInput, 500)
    cmdPrep1.Parameters.Append prm2
    
    
    cmdPrep1("datavalue") = inputVal
    cmdPrep1("keyvalue") = keyValue
    
    cmdPrep1.Execute

End Sub


Public Sub updateStringValueStringKey_OLD(ByVal tableName As String, _
                                        ByVal colName As String, ByVal inputVal As String, _
                                        ByVal keyName As String, ByVal keyValue As String)
                                        
    Dim strSQL As String
    
    If checkStringKeyExists(tableName, keyName, keyValue) = False Then
        MsgBox ("[modDBInterface:updateStringValueStringKey] Warning: la chiave <" & keyValue & "> non e' presente nella tabella " & _
                tableName & "(" & keyName & ")")
    End If
    
    strSQL = "UPDATE " & tableName & " SET " & colName & "='" & inputVal & _
             "' WHERE " & keyName & "='" & keyValue & "';"
    Call execCommandSQL(strSQL)

End Sub


Public Sub updateStringValueNumKey(ByVal tableName As String, _
        ByVal colName As String, ByVal inputVal As String, _
        ByVal keyName As String, ByVal keyValue As Double)

    Dim strSQL As String
    
    If checkNumKeyExists(tableName, keyName, keyValue) = False Then
        MsgBox ("[modDBInterface:updateStringValueNumKey] Warning: la chiave <" & keyValue & "> non e' presente nella tabella " & _
                tableName & "(" & keyName & ")")
    End If
    
    strSQL = "UPDATE " & tableName & " SET " & colName & "='" & inputVal & _
             "' WHERE " & keyName & "=" & keyValue & ";"
    Call execCommandSQL(strSQL)

End Sub

Public Sub updateNumValueNumKey(ByVal tableName As String, _
        ByVal colName As String, ByVal inputVal As Double, _
        ByVal keyName As String, ByVal keyValue As Double)

    Dim strSQL As String
    
    If checkNumKeyExists(tableName, keyName, keyValue) = False Then
        MsgBox ("[modDBInterface:updateNumValueNumKey] Warning: la chiave <" & keyValue & "> non e' presente nella tabella " & _
                tableName & "(" & keyName & ")")
    End If
    
    strSQL = "UPDATE " & tableName & " SET " & colName & "=" & inputVal & _
             " WHERE " & keyName & "=" & keyValue & ";"
    Call execCommandSQL(strSQL)

End Sub

Public Sub updateNumValueNumKey_noDelim(ByVal tableName As String, _
        ByVal colName As String, ByVal inputFormula As String, _
        ByVal keyName As String, ByVal keyValue As Double)

    Dim strSQL As String
    
    If checkNumKeyExists(tableName, keyName, keyValue) = False Then
        MsgBox ("[modDBInterface:updateNumValueNumKey_noDelim] Warning: la chiave <" & keyValue & "> non e' presente nella tabella " & _
                tableName & "(" & keyName & ")")
    End If
    
    strSQL = "UPDATE " & tableName & " SET " & colName & "=" & inputFormula & _
             " WHERE " & keyName & "=" & keyValue & ";"
    Call execCommandSQL(strSQL)

End Sub


Public Sub updateDateValueStringKey(ByVal tableName As String, _
                ByVal colName As String, ByVal inputVal As Date, _
                ByVal keyName As String, ByVal keyValue As String)

    Dim strSQL As String
    
    If checkStringKeyExists(tableName, keyName, keyValue) = False Then
        MsgBox ("[modDBInterface:updateDateValueStringKey] Warning: la chiave <" & keyValue & "> non e' presente nella tabella " & _
                tableName & "(" & keyName & ")")
    End If
    
    strSQL = "UPDATE " & tableName & " SET " & colName & "='" & Format(inputVal, "yyyy-mm-dd") & _
             "' WHERE " & keyName & "='" & keyValue & "';"
    Call execCommandSQL(strSQL)
    
End Sub


Public Sub updateDateValueNumKey(ByVal tableName As String, _
                ByVal colName As String, ByVal inputVal As Date, _
                ByVal keyName As String, ByVal keyValue As Long)

    Dim strSQL As String
    
    If checkNumKeyExists(tableName, keyName, keyValue) = False Then
        MsgBox ("[modDBInterface:updateDateValueIntKey] Warning: la chiave <" & keyValue & "> non e' presente nella tabella " & _
                tableName & "(" & keyName & ")")
    End If
    
    strSQL = "UPDATE " & tableName & " SET " & colName & "='" & Format(inputVal, "yyyy-mm-dd") & _
             "' WHERE " & keyName & "=" & keyValue & ";"
    Call execCommandSQL(strSQL)
    
End Sub

Public Sub updateNumValueStringKey(ByVal tableName As String, _
                ByVal colName As String, ByVal inputVal As Double, _
                ByVal keyName As String, ByVal keyValue As String) '
        
    Dim strSQL As String
    
    If checkStringKeyExists(tableName, keyName, keyValue) = False Then
        MsgBox ("[modDBInterface:updateNumValueStringKey] Warning: la chiave <" & keyValue & "> non e' presente nella tabella " & _
                tableName & "(" & keyName & ")")
    End If
            
    strSQL = "UPDATE " & tableName & " SET " & colName & "=" & inputVal & " WHERE " & _
              keyName & "='" & keyValue & "';"
    Call execCommandSQL(strSQL)

End Sub


Public Sub deleteStringKey(ByVal tableName As String, _
                ByVal keyName As String, ByVal keyValue As String, Optional doErrorCheck As Boolean = False)
        
    Dim strSQL As String
    
    strSQL = "DELETE FROM " & tableName & " WHERE " & keyName & "='" & keyValue & "';"
    Call execCommandSQL(strSQL)

End Sub


Public Sub deleteNumKey(ByVal tableName As String, _
                ByVal keyName As String, ByVal keyValue As Double, Optional doErrorCheck As Boolean = False)
        
    Dim strSQL As String
    
    strSQL = "DELETE FROM " & tableName & " WHERE " & keyName & "=" & keyValue & ";"
    Call execCommandSQL(strSQL)

End Sub


Public Function checkStringKeyExists(ByVal tableName As String, _
                ByVal keyName As String, ByVal keyValue As String) As Boolean
        Dim strSQL As String
        
    strSQL = "SELECT count(*) FROM " & tableName & " WHERE " & keyName & "='" & keyValue & "';"
    If execScalarSQL(strSQL) > 0 Then
        checkStringKeyExists = True
    Else
        checkStringKeyExists = False
    End If

End Function

Public Function checkNumKeyExists(ByVal tableName As String, _
                ByVal keyName As String, ByVal keyValue As Double) As Boolean
        Dim strSQL As String
        
    strSQL = "SELECT count(*) FROM " & tableName & " WHERE " & keyName & "=" & keyValue & ";"
    If execScalarSQL(strSQL) > 0 Then
        checkNumKeyExists = True
    Else
        checkNumKeyExists = False
    End If

End Function


Public Function getScalarStringKey(ByVal tableName As String, _
                ByVal valueColumnName As String, ByVal keyColumnName As String, _
                ByVal keyValue As String)
                
    getScalarStringKey = execScalarSQL("select " & valueColumnName & " from " & tableName & _
                                       " where " & keyColumnName & "='" & keyValue & "';")
End Function

Public Function getScalarNumKey(ByVal tableName As String, _
                ByVal valueColumnName As String, ByVal keyColumnName As String, _
                ByVal keyValue As Double)
                
    getScalarNumKey = execScalarSQL("select " & valueColumnName & " from " & tableName & _
                                       " where " & keyColumnName & "=" & keyValue & ";")
End Function




Public Sub logToFile(message As String)

    Dim dateTime, fileExcelName, finalErrorMsg As String
    
    dateTime = Now
    fileExcelName = ThisWorkbook.Name
    
    finalErrorMsg = Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & message
    
    ' Write to txt logFile on disk.
    Dim logFileName As String
    logFileName = ThisWorkbook.Path & "\" & fileExcelName & ".log"
    
    Dim fs As Variant
    Dim a As Variant
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.OpenTextFile(logFileName, 8, True, 0)
    a.WriteLine (finalErrorMsg)
    a.Close
    
End Sub
