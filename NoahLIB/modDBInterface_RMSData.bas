Attribute VB_Name = "modDBInterface_RMSData"
Option Explicit


' insertRMSAnalysis
' inserisce un analysis (group) nel DB MySQL
' strGroupName è il testo "Group Description" che l'utente inserisce nel form RMS_import
' analysisId è l'ID numerico, che viene associato al group

Public Sub insert_RMSAnalysis(ByVal strGroupName As String, ByRef analysisId As Long)
    Dim error1 As Boolean
    error1 = False
    
    Call DB_Utilities.execCommandSQL("insert into tblRMSList(strGroupName) values ('" & strGroupName & "');", error1)
    
    If error1 = True Then
        MsgBox ("Unable to retrieve an analysisId for " & strGroupName & "..")
    Else
        analysisId = DB_Utilities.execScalarSQL("select last_insert_id();")
    End If

End Sub

'Public Sub insertRMSAnalysis(ByVal strGroupName As String, ByRef analysisId As Long)
'    Call insert_RMSAnalysis(strGroupName, analysisId)
'End Sub


' assoc_CatSwapProgram_RMSanalysis
' associa un (CatSwap)Program al RMS group

Public Sub assoc_CatSwapProgram_RMSanalysis(ByVal programUMR As String, ByVal analysisId As Long)
    Dim strSQL As String
    
    If checkNumKeyExists("tblRMSList", "intID", analysisId) = False Then
        MsgBox ("Warning: la chiave <" & analysisId & "> non e' presente nella tabella tblRMSList(intID)")
    End If
    
    If checkStringKeyExists("tblCatSwapProgram", "strUMR", programUMR) = False Then
        MsgBox ("Warning: il CatSwap Program <" & programUMR & "> non e' presente nella tabella tblCatSwapProgram(strUMR)")
    End If
    
    strSQL = "INSERT INTO tblProgramRMSAnalysis(intRMSAnalysisID, strUMR) VALUES (" & _
            analysisId & ",'" & programUMR & "');"
    
    DB_Utilities.execCommandSQL (strSQL)
    

    'updates strOwner in tblrmslist
    Call modDBInterface.updateStringValueNumKey("tblrmslist", "strowner", programUMR, "intId", analysisId)

End Sub


' select_RMSAnalysis_forProgram
' fa select degli ID associati con il CatSwap program

Public Sub select_RMSAnalysis_forProgram(ByVal programUMR As String, ByRef rs As ADODB.Recordset)
    Dim strSQL As String
    
    strSQL = "SELECT intRMSAnalysisID, strGroupname from tblProgramRMSAnalysis " & _
             " inner join tblRMSList ON tblProgramRMSAnalysis.intRMSAnalysisID=tblRMSList.intID " & _
            "  WHERE strUMR='" & programUMR & "';"
     
    Call DB_Utilities.execTableSQL_withRS(strSQL, rs)

End Sub


' select_RMS_ELT_byID
' dato un Analysis ID, restituisce la ELT con questo ID
'Public Sub select_RMS_ELT_byID(ByVal analysisId As Long, ByRef rs As ADODB.Recordset)
'    Dim strSQL As String
'    strSQL = "select * from tbleltrms where intRMS=" & analysisId & ";"
'    Call DB_Utilities.execTableSQL_withRS(strSQL, rs)
'End Sub


' select_RMS_ELT_byLayerName
' dato un LayerName, trova l' RMSAnalysisID associato, e restituisce la ELT con questo ID
'Public Sub select_RMS_ELT_byLayerName(ByVal strLayerName As String, ByRef rs As ADODB.Recordset)
'    Dim strSQL As String
'    Dim analysisId As Long
'    strSQL = "SELECT intRMSAnalysisID FROM tblCatSwapLayer where strLayerName='" & strLayerName & "';"
'    analysisId = DB_Utilities.execScalarSQL(strSQL)
'    If checkStringKeyExists("tblCatSwapLayer", "strLayerName", strLayerName) = False Then
'        MsgBox ("Warning: la chiave <" & strLayerName & "> non e' presente nella tabella tblCatSwapLayer(strLayerName)")
'    End If
'    If analysisId > 0 Then
'        strSQL = "select * from tbleltrms where intRMS=" & analysisId & ";"
'        Call DB_Utilities.execTableSQL_withRS(strSQL, rs)
'    End If
'End Sub



' update CatSwap Layer, sets the right selected analysis ID
'Public Sub update_CatSwapLayer_RMSAnalysis(strLayerName As String, inputVal As Double)
'    Call updateNumValueStringKey("tblCatSwapLayer", "intRMSAnalysisID", inputVal, "strLayerName", strLayerName)
'End Sub

' update CatSwap Layer, sets the "analysis type" for sheet "RMS"
'Public Sub update_CatSwapLayer_RMSAnalysisType(strLayerName As String, inputVal As String)
'    Call updateStringValueStringKey("tblCatSwapLayer", "strRMSAnalysisType", inputVal, "strLayerName", strLayerName)
'End Sub


' get the number of RMS analyses associated with one Program (by UMR)
Public Function getAnalysesCount_byUMR(ByVal UMR As String) As Long
    On Error GoTo exitZero
    getAnalysesCount_byUMR = DB_Utilities.execScalarSQL("SELECT count(*) FROM tblprogramrmsanalysis WHERE strUMR='" & _
                                                    UMR & "' ;")
    Exit Function
exitZero:
    On Error GoTo 0
    getAnalysesCount_byUMR = 0
End Function








