Attribute VB_Name = "CatraderImport"


'*******************************************************************************************************
' Sub: getCatraderELT
' Param: connCatrader: open connection to the Catrader SQL Server DB
' Param: recordsetELTCatrader: output recordset
' Param: strGuidCondition: guid for the condition (key of tblCondition in Catrader AIRCT2Exp DB)
' Descr: generates the Catrader ELT (with dblPercLoss) with the right format to import to NOAH-MySQL
Public Sub getCatraderELT(ByRef connCatrader As ADODB.Connection, _
ByRef recordsetELTCatrader As ADODB.Recordset, _
ByVal strGUIDCondition As String, _
Optional showWarnings As Boolean = True)

Dim dbloccret, dblocclmt, dblaggret, dblagglmt As Double
Dim fltCoinsurance, dblmaxloss As Double

    '************************************************************************************************
    'gets Condition data from Catrader
    strSQL = "SELECT a.strName, a.dblOccRet,a.dblocclmt, a.dblaggret, a.dblagglmt, " & _
                " a.fltCoinsurance as contractName " & _
                " FROM airct2exp..tblCondition as a " & _
                " WHERE guidCondition = " & strGUIDCondition
    Call DB_Utilities.execTableSQL_withConn(strSQL, connCatrader)
        
    dbloccret = rsUtil(1).value
    dblocclmt = rsUtil(2).value
    dblaggret = rsUtil(3).value
    dblagglmt = rsUtil(4).value
    fltCoinsurance = rsUtil(5).value
        
    If dblagglmt > 0 Then
            dblmaxloss = dblagglmt
        Else
            dblmaxloss = dblocclmt
    End If

    If dblmaxloss = 0 Then
        If showWarnings = True Then
            MsgBox "[CatraderImport:getCatraderELT]" & vbCrLf & "Warning, the Condition <" & rsUtil(0).value & "> has no Occurrence and no Aggregate Limits. This ELT will not be imported."
        End If
'        dblmaxloss = 1
        Set recordsetELTCatrader = Nothing
        Exit Sub
    End If
    
    Set recordsetELTCatrader = New ADODB.Recordset
    recordsetELTCatrader.CursorLocation = adUseServer
    
    ' gets ELT for this Condition, from Catrader
    strSQL = "SELECT intYear, intEvent, intModel, ROUND(SUM(dblTotal),1) as contractLoss, " & _
             " guidCondition, ROUND(SUM(dblTotal),1)/" & (dblmaxloss * fltCoinsurance) & " as dblLossPerc " & _
             " FROM AirCT2Loss..TblConditionLoss " & _
             " WHERE guidcondition = " & strGUIDCondition & _
             " AND guidEventSet=0x00000000000200500071600000000010" & _
             " AND intModel <> 0 " & _
             " GROUP BY intEvent, intModel, intYear, guidCondition " & _
             " ORDER BY intyear, intevent "
    
    'the result recordset is returned via the "recordsetELTCatrader" variable
    recordsetELTCatrader.Open strSQL, connCatrader

End Sub


' imports one ELT sotred in the CSV, to the DB
' using CSV
Sub ImportELTCatrader_CSV(ByVal fileName As String, ByRef conn As ADODB.Connection)
    Dim strSQL As String
    strSQL = getSQLloadCsvELT_Catrader(fileName)
    conn.Execute strSQL
End Sub




' insertCatraderAnalysis
' crea una nuova Condition (catrader) in tblCondition
' associa il nuovo id, e restituisce newID

Public Sub insertCatraderCondition(ByVal strName As String, ByRef newID As Long)
    Call modDBInterface.insertStringKey("tblCondition", "strName", strName)
    newID = DB_Utilities.execScalarSQL("select last_insert_id();")
End Sub




' gets basic Condition info from Catrader (limit, deductible, etc.
Public Sub getConditionDataFromCT(ByVal GUIDCondition As String, ByRef rs As Recordset)
Dim connCatrader As ADODB.Connection
Dim strSQL As String

Set connCatrader = getSQLServerConnectionCatrader()

strSQL = "SELECT a.strName, dblOccLmt, dblOccRet, dblAggLmt, dblAggRet, fltCoinsurance, b.strViewCurrency as strCcy , a.intReinstNumber  " & _
                    " FROM airct2exp..tblCondition as a INNER JOIN airct2exp..tblContract as b ON a.guidContract=b.guidContract WHERE guidCondition = " & GUIDCondition
Call DB_Utilities.execTableSQL_withConn2(strSQL, connCatrader, rs)
        
        
End Sub

' updates Condition info, from Catrader
' Condition info are stored in the "rs" parameter
Public Sub updateConditionDataToLossdb(ByVal GUIDCondition As String, ByRef rs As Recordset)
    If rs Is Nothing Then
        MsgBox "Condition <" & GUIDCondition & "> not found on Catrader, exit."
        End
    End If
    
    rs.MoveFirst
    Call modDBInterface.updateNumValueStringKey("tblCondition", "dblOccLmt", rs("dblOccLmt"), "strguidcondition", GUIDCondition)
    Call modDBInterface.updateNumValueStringKey("tblCondition", "dblOccRet", rs("dblOccRet"), "strguidcondition", GUIDCondition)
    Call modDBInterface.updateNumValueStringKey("tblCondition", "dblAggLmt", rs("dblAggLmt"), "strguidcondition", GUIDCondition)
    Call modDBInterface.updateNumValueStringKey("tblCondition", "dblAggRet", rs("dblAggRet"), "strguidcondition", GUIDCondition)
    Call modDBInterface.updateNumValueStringKey("tblCondition", "dblCoinsurance", rs("fltCoinsurance"), "strguidcondition", GUIDCondition)
    Call modDBInterface.updateNumValueStringKey("tblCondition", "intReinstNumber", rs("intReinstNumber"), "strguidcondition", GUIDCondition)
    Call modDBInterface.updateStringValueStringKey("tblCondition", "strCcy", rs("strCcy"), "strguidcondition", GUIDCondition)
End Sub

' updates the condition data from Catrader
Public Sub updateConditionData(ByVal GUIDCondition As String)
    Dim rs As Recordset
    ' gets data
    Call getConditionDataFromCT(GUIDCondition, rs)
    ' updates data
    Call updateConditionDataToLossdb(GUIDCondition, rs)
End Sub



' importa la ELT da Catrader, intCond è l'ID da associare in LOSSDB (key [intID] of tblCondition)
'(deve essere già presente un record in tblCondition
' usa i CSV
Public Sub import_ELT_Catrader(ByVal intCond As Long, ByVal GUIDCondition As String, Optional showWarnings As Boolean = True)

    Dim rs As ADODB.Recordset
    Dim connCatrader As ADODB.Connection
    Dim fileName As String
    
    ' open Catrader connection
    Set connCatrader = getSQLServerConnectionCatrader()
    
    If showWarnings = True Then
        MsgBox "Importing GUID Condition <" & GUIDCondition & "> to the analysis ID: <" & intCond & ">"
    End If
    
    'associa la guidCondition all'ID (intCond)
    Call modDBInterface.updateStringValueNumKey("tblCondition", "strGuidCondition", GUIDCondition, "intId", intCond)
    
    'retrieve ELT from Catrader
    Call getCatraderELT(connCatrader, rs, GUIDCondition, showWarnings)
    
    ' rs could be empty!
    If check_Recordset_NotEmpty(rs) = False Then
        If showWarnings = True Then
            Call MsgBox("Condition <" & intCond & "> got an empty ELT from Catrader, no points are imported for this ELT." & vbCrLf & _
                   "Check that the analysis has been run on Catrader with <Use Saved Results> and using the <World Perils 10K Time Dep Hybrid> Event Set, and that " & vbCrLf & _
                   "Occurrence Limit and Aggregate Limit are not both ZERO.", vbExclamation)
        End If
        Exit Sub
    End If
    
    ' CSV temporary filename
    fileName = ActiveWorkbook.Path & "\" & "ELT_Catrader_" & intCond & ".csv"
    
    'export to csv
    Call ExportToCsvELTCatrader(rs, fileName, intCond)
    
    'copy to repo
    Dim fso As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    ' copies the CSV to the CSV_repo
    Call fso.CopyFile(fileName, "K:\Shared Modeling\ANALISI FUNZIONALI\NOAH\Database\CSV_repo\ELT_Catrader_" & intCond & ".csv")
    
    ' imports the CSV to the DB
    Call ImportELTCatrader_CSV(fileName, connMyUtil)
    
    ' deletes the temporary CSV
    Call RemoveFile(fileName)
    

End Sub






