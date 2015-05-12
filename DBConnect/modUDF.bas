Attribute VB_Name = "modUDF"
Option Explicit

' getTable is the generic table-retrieve function.
' this function is called from excel, puts one request to the request-listener, and exits
' the listener will carry the task of executiong the function, and write the data to Excel.

Public Function getTable(ByVal datatype As String, _
                Optional ByVal key As String = "", _
                Optional ByVal param1 As String = "", _
                Optional ByVal param2 As String = "", _
                Optional ByVal param3 As String = "", _
                Optional ByVal output_byRow As Boolean = False) As String

'    Application.Volatile
    Dim i As Integer
    
    ' check if app is up and running
    If (oApp Is Nothing) Or (CONTINUE_loop = False) Then
        getTable = "Link closed? Click AddIns->DB Connect->Refresh Data"
        Exit Function
    End If
        
    ' this object is used to store the request info/data
    Dim oRequest As cRequest
    Set oRequest = New cRequest
    Call oRequest.init
        
    i = 1
    ' caller range
    Set oRequest.rg_caller = Application.Caller.Cells(i, 1)
    i = i + 1

    If (SHOW_logs = True) Then
        ' log range
        Set oRequest.rg_run_log = Application.Caller.Cells(i, 1)
        i = i + 1
    End If
    
    If (SHOW_RUN_details = True) Then
        ' run_details range
        Set oRequest.rg_run_details = Application.Caller.Cells(i, 1)
        i = i + 1
    End If

    ' TARGET RANGE
    Set oRequest.rg_target = Application.Caller.Cells(i, 1)
    


'    oRequest.request = ""
    ' PASS DATA to the request object
    oRequest.datatype = datatype
    oRequest.key = key
    oRequest.param1 = param1
    oRequest.param2 = param2
    oRequest.param3 = param3
'    oRequest.param4 = param4
    
    ' sets if output is by row
    oRequest.output_byRow = output_byRow
    If output_byRow = True Then
        Set oRequest.rg_target = Application.Caller.Cells(1, 2)
    End If
 
'    submit data to the listener
    Call oApp.cRequestList.Add(oRequest)
    
    ' writes something in the caller cell
    getTable = "" & Application.Caller.Formula
    '& " .. param =  " & datatype & ";" & key
    
End Function


' retrieve the YLT for the specified asset
Public Function getYLT(ByVal assetCode As String, _
                    Optional ByVal aggregateOccurrence As String = "AGG", _
                    Optional ByVal modeler As String = "AIR")
    getYLT = getTable("YLT", assetCode, aggregateOccurrence, modeler)
End Function


' retrieve the ELT for the specified asset
Public Function getELT(ByVal assetCode As String, _
                    Optional ByVal modeler As String = "AIR", _
                    Optional ByVal perspective As String = "NET_LAYER")
    getELT = getTable("ELT", assetCode, modeler, perspective)
End Function

' retrieves the portfolio composition
Public Function getPortfolio(ByVal portfolioName As String)
    getPortfolio = getTable("PORT", portfolioName)
End Function






