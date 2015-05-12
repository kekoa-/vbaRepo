Attribute VB_Name = "modDBInterface_AIRData"
Option Explicit

' insert_AIRAnalysis
' inserisce una analysis/ELT nel DB MySQL


'Public Sub insert_AIRAnalysis(ByVal strGUIDCondition As String, ByRef analysisId As Long)
'    Dim error1 As Boolean
'    error1 = False
'    Call DB_Utilities.execCommandSQL("insert into tblCondition(strGUIDCondition) values ('" & strGUIDCondition & "');", error1)
'    If error1 = True Then
'        MsgBox ("Unable to retrieve an idCondition for " & strGUIDCondition & "..")
'    Else
'        analysisId = DB_Utilities.execScalarSQL("select last_insert_id();")
'    End If
'End Sub



' assoc_CatSwapProgram_AIRcondition
' associa un (CatSwap)Program alla condition
'Public Sub assoc_CatSwapProgram_AIRcondition(ByVal programUMR As String, ByVal analysisId As Long)
'    Dim strSQL As String
'    If checkNumKeyExists("tblCondition", "intID", analysisId) = False Then
'        MsgBox ("Warning: la chiave <" & analysisId & "> non e' presente nella tabella tblCondition(intID)")
'    End If
'    If checkStringKeyExists("tblCatSwapProgram", "strUMR", programUMR) = False Then
'        MsgBox ("Warning: la chiave <" & programUMR & "> non e' presente nella tabella tblCatSwapProgram(strUMR)")
'    End If
'    strSQL = "INSERT INTO tblProgramCondition(intCondition, strUMR) VALUES (" & _
'            analysisId & ",'" & programUMR & "');"
'    DB_Utilities.execCommandSQL (strSQL)
'End Sub



' select_AIRAnalysis_forProgram
' fa select degli ID associati con il CatSwap program
'Public Sub select_AIRAnalysis_forProgram(ByVal programUMR As String, ByRef rs As ADODB.Recordset)
'    Dim strSQL As String
'
'    strSQL = "SELECT intCondition, strGUIDCondition from tblProgramCondition " & _
'             " inner join tblCondition ON tblProgramCondition.intCondition=tblCondition.intID " & _
'            "  WHERE strUMR='" & programUMR & "';"
'    Call DB_Utilities.execTableSQL_withRS(strSQL, rs)
'End Sub



' update CatSwap Layer, sets the right selected analysis ID
'Public Sub update_CatSwapLayer_AIRAnalysis(strLayerName As String, inputVal As Double)
'    Call updateNumValueStringKey("tblCatSwapLayer", "intCondition", inputVal, "strLayerName", strLayerName)
'End Sub

' update CatSwap Layer, sets the "analysis type" for sheet "AIR"
'Public Sub update_CatSwapLayer_AIRAnalysisType(strLayerName As String, inputVal As String)
'    Call updateStringValueStringKey("tblCatSwapLayer", "strAIRAnalysisType", inputVal, "strLayerName", strLayerName)
'End Sub




Public Function getAnalysesCount_byOwner(ByVal ownerCode As String, Optional ByVal ownerType As String = "") As Long
On Error GoTo exitZero

getAnalysesCount_byOwner = DB_Utilities.execScalarSQL("SELECT count(*) FROM tblCondition WHERE strOwner='" & _
                                                    ownerCode & "' ;")

Exit Function

exitZero:
On Error GoTo 0
getAnalysesCount_byOwner = 0

End Function



Public Function getAnalysisList_byOwner(ByRef rs As Recordset, ByVal ownerCode As String) As Boolean
On Error GoTo exitZero

Call DB_Utilities.execTableSQL_withRS("SELECT intId, strName from tblCondition WHERE strOwner='" & ownerCode & "' ;", rs)
If rs Is Nothing Then
    getAnalysisList_byOwner = False
Else
    getAnalysisList_byOwner = True
End If

Exit Function

exitZero:
On Error GoTo 0
getAnalysisList_byOwner = False

End Function










