Attribute VB_Name = "mod_btn_GetData"
Option Explicit


' execute a SQL query, retrieves data in the recordset and gives the record_count
Public Function get_rsAndCount(ByVal strSQL As String, ByRef rs As ADODB.Recordset, ByRef numRowsFound As Long) As Boolean
    On Error GoTo errFunction
    get_rsAndCount = False
    Call DB_Utilities.execTableSQL_withRS(strSQL, rs)
    numRowsFound = getRowsFoundCount()
    get_rsAndCount = True
    On Error GoTo 0
        Exit Function
    
errFunction:
    On Error GoTo 0
    MsgBox "Error: [mod_btn_GetData:get_rsAndCount]" & vbCrLf & "SQL: " & strSQL & vbCrLf & vbCrLf & Err.Description
    get_rsAndCount = False
    
End Function


' executes the SQL query, retrieves the data and record count, resizes the range,
' and outputs the data to the range
' optionally, writes the header to the rh_header range.
Public Function putRsIntoRange(ByVal strSQL As String, ByRef rg As Range, Optional ByRef rg_header As Range = Nothing) As Boolean
    On Error GoTo errFunction


    Dim rs As ADODB.Recordset
    Dim nRows As Long
    Dim res As Boolean
    
    res = get_rsAndCount(strSQL, rs, nRows)
    
    If Not rs Is Nothing Then
        Call range_resizeRows(rg, nRows)
        rg.ClearContents
        Call rg.CopyFromRecordset(rs)
        
        If Not rg_header Is Nothing Then
            Call mod_helper.writeRecordsetColumnHeader(rg_header, rs)
        End If
    
    End If
    
    putRsIntoRange = True
    
    On Error GoTo 0
    Exit Function
    
errFunction:
    On Error GoTo 0
    MsgBox "[mod_btn_GetData:putRsIntoRange] Error: " & vbCrLf & vbCrLf & "SQL: " & strSQL & Err.Description, vbExclamation, "Error"
    putRsIntoRange = False
End Function




' executes the SQL query, retrieves the data and record count, DOES NOT resizes the range,
' and outputs the data to the range
' optionally, writes the header to the rg_header range.
Public Function putRsIntoRange_noResize(ByVal strSQL As String, ByRef rg As Range, Optional ByRef rg_header As Range = Nothing) As Boolean
    On Error GoTo errFunction


    Dim rs As ADODB.Recordset
    Dim nRows As Long
    Dim nCols As Long
    Dim res As Boolean
    Dim rg1 As Range, rg2 As Range, rg3 As Range
    Dim ws1 As Worksheet
    
    res = get_rsAndCount(strSQL, rs, nRows)
    
    
    If Not rs Is Nothing Then
        'Call range_resizeRows(rg, nRows)
        nCols = rs.Fields.count
        
        Set ws1 = rg.Parent
        Set rg1 = rg.Cells(1, 1)
        Set rg2 = rg1.Cells(nRows, nCols)
        Set rg3 = ws1.Cells.Range(rg1, rg2)
        
        Set rg = rg3
        
        If Not rg_header Is Nothing Then
            Set rg_header = ws1.Cells.Range(rg_header.Cells(1, 1), rg_header.Cells(1, nCols))
        End If
        
        Set rg1 = Nothing
        Set rg2 = Nothing
        Set rg3 = Nothing
        
        
        rg.ClearContents
        Call rg.CopyFromRecordset(rs)
        
        If Not rg_header Is Nothing Then
            Call mod_helper.writeRecordsetColumnHeader(rg_header, rs)
        End If
    
    End If
    
    putRsIntoRange_noResize = True
    
    On Error GoTo 0
    Exit Function
    
errFunction:
    On Error GoTo 0
    MsgBox "[mod_btn_GetData:putRsIntoRange] Error: " & vbCrLf & vbCrLf & "SQL: " & strSQL & Err.Description, vbExclamation, "Error"
    putRsIntoRange_noResize = False
End Function




' executes the SQL query, retrieves the data and record count, resizes the range,
' and outputs the data to the range
' insertion of rows is performed on the worksheet rows
Public Function putRsIntoRange_addByRow(ByVal strSQL As String, ByRef rg As Range) As Long
    On Error GoTo errFunction

    Dim rs As ADODB.Recordset
    Dim nRows As Long
    Dim res As Boolean
    
    res = get_rsAndCount(strSQL, rs, nRows)
    
    If Not rs Is Nothing Then
        Call range_insertRows_worksheet(rg, nRows + 1)
        Call rg.CopyFromRecordset(rs)
    End If
    
    putRsIntoRange_addByRow = nRows
    
    On Error GoTo 0
    Exit Function
    
errFunction:
    On Error GoTo 0
    If rg Is Nothing Then
        MsgBox "[mod_btn_GetData:putRsIntoRange_addByRow] Null Range! "
        Exit Function
    End If
    MsgBox "[mod_btn_GetData:putRsIntoRange_addByRow] Error: " & vbCrLf & vbCrLf & "SQL: " & strSQL & Err.Description, vbExclamation, "Error"
    putRsIntoRange_addByRow = 0
End Function


' get asset_list to the specified range
Public Function getAssetList_short(ByRef rg As Range) As Boolean
    Dim strSQL As String
    strSQL = "SELECT strCode, strNick, strName, strCcy, strAssetType from tblAsset order by strAssetType, strName;"
    getAssetList_short = putRsIntoRange(strSQL, rg)
End Function

' get asset_portfolio list to the specified range
Public Function getAssetPortfolio_list(ByRef rg As Range, ByVal portfolioName As String) As Boolean
    Dim strSQL As String
    strSQL = "CALL prcGetAssetPortfolio_byport('" & portfolioName & "');"
    getAssetPortfolio_list = putRsIntoRange(strSQL, rg)
End Function


' get movements list to the specified range
Public Function getMovementsList_byFundId(ByRef rg As Range, ByVal intFund As Long) As Boolean
    Dim strSQL As String
    strSQL = "call prcGetMovementsByfund(" & intFund & ")"
    getMovementsList_byFundId = putRsIntoRange(strSQL, rg)
End Function















