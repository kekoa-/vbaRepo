Attribute VB_Name = "ExportToCsv"
Option Explicit



'*******************************************************************************************************
' Sub: ExportToCsvELTCatrader
' Param: recordsetELTCatrader: è il recordset contenente la ELT da catrader
' Param: strExportFile: percorso del file su cui scrivere
' Descr: exports the recordsetELTCatrader recordset into a Csv
Sub ExportToCsvELTCatrader(ByRef recordsetELTCatrader As ADODB.Recordset, ByVal strExportFile As String, _
                           ByVal intCondition As Long)
    Dim intFileNum As Long
    Dim varData As String, strDelimiter As String
    
    ' uses TAB field delimiter
    strDelimiter = vbTab
    'get file handle and opens for output
    intFileNum = FreeFile()
    Open strExportFile For Output As #intFileNum
    
    recordsetELTCatrader.MoveFirst
    'CYCLE over the recordset
    Do While Not recordsetELTCatrader.EOF
        ' gets data row
        varData = ""
        varData = varData & intCondition & strDelimiter
        varData = varData & recordsetELTCatrader("intYear") & strDelimiter
        varData = varData & recordsetELTCatrader("intEvent") & strDelimiter
        varData = varData & recordsetELTCatrader("intModel") & strDelimiter
        varData = varData & recordsetELTCatrader("contractLoss") & strDelimiter
        varData = varData & recordsetELTCatrader("dblLossPerc")
        ' writes data to file
        Print #intFileNum, varData
        recordsetELTCatrader.MoveNext
    Loop
    'close the file
    Close #intFileNum
End Sub


Sub RemoveFile(ByVal fileName As String)
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
 
If fso.FileExists(fileName) = False Then
    MsgBox "<" & fileName & "> not found, exit"
    Exit Sub
End If

fso.deletefile fileName

End Sub


'*******************************************************************************************************
' Sub: ExportToCsvELTRMS
' Param: recordsetELTRMS: è il recordset contenente la ELT da RMS
' Param: strExportFile: percorso del file su cui scrivere
' Descr: exports the recordsetELTRMS recordset into a Csv
Sub ExportToCsvELTRMS(ByRef recordsetELTRMS As ADODB.Recordset, ByVal strExportFile As String)
    Dim intFileNum, intColumnCount, i As Long
    Dim varData As String, strDelimiter As String
    
    strDelimiter = vbTab
    ' get file handle and open for output
    intFileNum = FreeFile()
    ' open file
    Open strExportFile For Append As #intFileNum
    
    If recordsetELTRMS Is Nothing Then
        MsgBox "ExportToCsvELTRMS" & vbCrLf & vbCrLf & "Empty Recordset, exit sub"
        Exit Sub
    End If
   
    'gets recordset column count
    intColumnCount = recordsetELTRMS.Fields.count
    
    recordsetELTRMS.MoveFirst
    'CYCLE over the recordset
    Do While Not recordsetELTRMS.EOF
        varData = ""
        
        For i = 0 To intColumnCount - 2
            varData = varData & recordsetELTRMS(i) & strDelimiter
        Next
        varData = varData & recordsetELTRMS(i - 1)
        ' writes data to file
        Print #intFileNum, varData
        recordsetELTRMS.MoveNext
    Loop
    'close the file
    Close #intFileNum
End Sub


' get SQL string for importing the .csv file containing the RMS ELT, to the db
Public Function getSQLloadCsvELT(ByVal fileName As String)
    Dim strSQL As String
    fileName = Replace(fileName, "\", "/")
    strSQL = _
        " LOAD DATA LOCAL INFILE '" & fileName & "'" & vbCrLf & _
        " INTO TABLE lossdb.tbleltrms" & vbCrLf & _
        " fields terminated by '\t' " & vbCrLf & _
        " lines terminated by '\n' " & vbCrLf & _
        " (intRMS, strAreaName, strPeril, intEventId, dblRate, dblPerspvalue, dblStdDevTot," & _
        " dblExpvalue, dblStddevi2, dblStddevc) ;"
    getSQLloadCsvELT = strSQL
End Function


' get SQL string for importing the .csv file containing the CATRADER ELT, to the db
Public Function getSQLloadCsvELT_Catrader(ByVal fileName As String)
    Dim strSQL As String
    fileName = Replace(fileName, "\", "/")
    strSQL = _
        " LOAD DATA LOCAL INFILE '" & fileName & "'" & vbCrLf & _
        " INTO TABLE lossdb.tblELTCatrader" & vbCrLf & _
        " fields terminated by '\t' " & vbCrLf & _
        " lines terminated by '\n' " & vbCrLf & _
        " (intCondition, intYear, intEventId, intModel, dblLoss, dblPercLoss ) ;"
    getSQLloadCsvELT_Catrader = strSQL
End Function



















