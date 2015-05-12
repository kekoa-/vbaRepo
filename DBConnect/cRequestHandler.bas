VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRequestHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Option Compare Text


' ok this is the MAIN object of the app
' cycles over the REQUEST list, and actually RETRIEVES the data using NOAH_LIB
' and actually PUTS the data to the Excel worksheets


Public Sub init()
    
End Sub


' CYCLE over the request list,
' removes completed requests
Public Sub freeCompletedRequests(ByRef oList As Collection)
    Dim i As Integer, n As Integer
    
    n = oList.count
    
    ' CYCLE over requests
    For i = n To 1 Step -1
        ' if caller is nothing, then request has been handled
        If (oList.Item(i).rg_caller Is Nothing) Then
            oList.Remove (i)
        End If
        
        On Error Resume Next
        
        If (SHOW_DEBUG = True) Then Cells(1, 1).value = "RequestList length = " & oList.count
        On Error GoTo 0
    Next

End Sub


' CYCLE over the request list,
' HANDLES the requests
Public Sub handleRequestList(ByRef oList As Collection)
    Dim o As cRequest
    Dim rs As Recordset
    
    Dim datatype As String, key As String
    
    
    ' CYCLE over requests
    For Each o In oList
        
        ' if rg_target is missing, skip&delete this request
        If o.rg_target Is Nothing Then
            o.destruct
            o.handled = True
            'o.highlighted = False
            'GoTo nextIteration1
        End If
        
        
        ' this is for: keep the green highlight for 3 cycles, then undo highlight
        Call o.checkHighlight
    
        ' skip request if already handled
        If o.handled = True Then GoTo nextIteration1
        
        ' clear output range before refreshing data
        If o.outputCleaned = False Then
            ' TODO: fix this!
            o.col_target_count = 1
            If o.datatype = "ELT" Then o.col_target_count = 9
            If o.datatype = "OEP" Then o.col_target_count = 3
            If o.datatype = "FIND" Then o.col_target_count = 3
            If o.datatype = "YLT" Then o.col_target_count = 1
            If o.datatype = "EL" Then o.col_target_count = 1
            If o.datatype = "PORT" Then o.col_target_count = 6
            If o.datatype = "CB_ANAGRAFICA" Then o.col_target_count = 75
            If o.datatype = "ASSET_TABLE" Then o.col_target_count = 20
            
            
            ' cleans the output range -
            ' LIGHT BLUE highlight on the caller range
            Call o.cleanOutput
            If Not o.rg_run_log Is Nothing Then
                o.rg_run_log.Cells(1, 1) = "Cleaning.."
            End If
            ' continue this request in the next HANDLE_REQUEST call..
            GoTo nextIteration1
        End If
    
        
        
    '########################################################################################
    '########################################################################################
    '########################################################################################
    ' MAIN WORK: OUTPUT DATA
    ' output results
    ' HERE do the main work
        
        datatype = o.datatype
        key = o.key
        
        If datatype = "ELT" Then
            
            If o.param1 = "AIR" Then
            
                Call db_utilities.execCommandSQL("select 0 into @key1;")
                Call db_utilities.execCommandSQL("select intcondition into @key1 from tblasset where strcode='" & key & "';")
                Call db_utilities.execTableSQL_withRS("select * from tbleltcatrader where intcondition=@key1;", rs)
                
                Call mod_helper.writeRecordsetColumnHeader(o.rg_target.Cells(1, 1), rs)
                Call safeCopyFromRecordset(o.rg_target.Cells(2, 1), rs)
            
            ElseIf o.param1 = "RMS" Then
                
                If o.param2 = "NET_LAYER" Then
                    
                    Call db_utilities.execCommandSQL("select 0 into @key1;")
                    Call db_utilities.execCommandSQL("select intRMSAnalysisToLayer into @key1 from tblCatSwapLayer where strcode='" & key & "';")
                    Call db_utilities.execTableSQL_withRS("select strAreaName, strPeril, intEventId, dblRate, dblPerspvalue, " & _
                        "dblStdDevTot, dblExpvalue, dblStddevi2, dblStddevc from tblELTRMS where intRMS=@key1;", rs)
                    Call mod_helper.writeRecordsetColumnHeader(o.rg_target.Cells(1, 1), rs)
                    Call safeCopyFromRecordset(o.rg_target.Cells(2, 1), rs)
                Else
                ' in any other case, retrieve net_to_program
                'If o.param2 = "NET_PROGRAM" Then
                
                    Call db_utilities.execCommandSQL("select 0 into @key1;")
                    Call db_utilities.execCommandSQL("select intRMSAnalysisToProgram into @key1 from tblCatSwapLayer where strcode='" & key & "';")
                    Call db_utilities.execTableSQL_withRS("select strAreaName, strPeril, intEventId, dblRate, dblPerspvalue, " & _
                        "dblStdDevTot, dblExpvalue, dblStddevi2, dblStddevc from tblELTRMS where intRMS=@key1;", rs)
                    Call mod_helper.writeRecordsetColumnHeader(o.rg_target.Cells(1, 1), rs)
                    Call safeCopyFromRecordset(o.rg_target.Cells(2, 1), rs)
                
                End If
                
            
            Else
                '...
            End If
            
        End If
        
        If datatype = "OEP" Then
            Call db_utilities.execTableSQL_withRS("call prcGetAssetYLT('" & key & "');", rs)
            Call mod_helper.writeRecordsetColumnHeader(o.rg_target.Cells(1, 1), rs)
            Call safeCopyFromRecordset(o.rg_target.Cells(2, 1), rs)
        End If
        
        If datatype = "YLT" Then
            If (o.param1 = "OCC") Or (o.param1 = "O") Then
                Call db_utilities.execTableSQL_withRS("call prcGetAssetYLT_Occurrence('" & key & "');", rs)
            End If
            If (o.param1 = "AGG") Or (o.param1 = "A") Then
                Call db_utilities.execTableSQL_withRS("call prcGetAssetYLT_Aggregate('" & key & "');", rs)
            End If
            Call mod_helper.writeRecordsetColumnHeader(o.rg_target.Cells(1, 1), rs)
            Call safeCopyFromRecordset(o.rg_target.Cells(2, 1), rs)
        End If
        
        
        If datatype = "PORT" Then
            Call db_utilities.execTableSQL_withRS("call prcGetAssetPortfolio_byport('" & key & "');", rs)
            Call mod_helper.writeRecordsetColumnHeader(o.rg_target.Cells(1, 1), rs)
            Call safeCopyFromRecordset(o.rg_target.Cells(2, 1), rs)
        End If
        
        
        If datatype = "CB_LIST" Then
            Call db_utilities.execTableSQL_withRS("select strCode from tblasset where strAssetType='CB' order by strName;", rs)
            Call safeCopyFromRecordset(o.rg_target.Cells(1, 1), rs)
        End If
        
        If datatype = "RE_LIST" Then
            Call db_utilities.execTableSQL_withRS("select strCode from tblasset where strAssetType='RE' order by strName;", rs)
            Call safeCopyFromRecordset(o.rg_target.Cells(1, 1), rs)
        End If
        
        If datatype = "ASSET_LIST" Then
            Call db_utilities.execTableSQL_withRS("select strCode from tblasset order by strAssetType, strName;", rs)
            Call safeCopyFromRecordset(o.rg_target.Cells(1, 1), rs)
        End If
        
        If datatype = "PORTFOLIO_LIST" Then
            Call db_utilities.execTableSQL_withRS("select strName from tblPortfolio order by strName;", rs)
            Call safeCopyFromRecordset(o.rg_target.Cells(1, 1), rs)
        End If
        
        If datatype = "TRADE_LIST" Then
            Call getTradeList(o.rg_target.Cells(2, 1), o.rg_target.Cells(1, 1), o.key)


        End If
        
        If datatype = "CB_ANAGRAFICA" Then
            Call getCBAnagraficaTable(o.rg_target.Cells(2, 1), o.rg_target.Cells(1, 1))
        End If
        
        If datatype = "ASSET_TABLE" Then
            Call getAssetTable(o.rg_target.Cells(2, 1), o.rg_target.Cells(1, 1))
        End If
                
        If (datatype = "FIND") And (key = "ASSET") Then
            Call getAssetFind(o.rg_target.Cells(2, 1), o.rg_target.Cells(1, 1), o.param1)
        End If


    
    ' log run details
    If Not o.rg_run_log Is Nothing Then
        o.rg_run_log.Cells(1, 1).value = "DONE"
    End If
    
    ' log run details
    If Not o.rg_run_details Is Nothing Then
        o.rg_run_details.Cells(1, 1).value = "UPDATED at : <" & Format(Now, "yyyy-mm-dd") & "> <" & Format(Now, "hh:mm:ss") & ">"
    End If
    
    ' GREEN - highlight run details
    Call o.doHighlight
        
    ' this has been handled
    o.handled = True
    
    ' calls destructor
    Call o.destruct
        
    ' END OF MAIN WORK
    '########################################################################################
    '########################################################################################
    '########################################################################################

            
        
nextIteration1:
    Next
    
    
    'free unused requests
    Call freeCompletedRequests(oList)
    
End Sub







