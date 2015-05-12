Attribute VB_Name = "NOAH_SQLquery"
Option Explicit
'##############################################################################################
' SQL RELATED FUNCTION AND SQL QUERIES
'   [1]  strConn                              (to be checked)
'   [2]  sqlAIR_AvailCompany                  (to be checked)
'   [3]  sqlRMS_AvailAnalyses                 (to be checked)
'   [4]  sqlAIR_AvailContract                 (to be checked)
'   [5]  sqlRMS_AvailProgramme                (to be checked)
'   [6]  sqlRMS_QueryELT                      (to be checked)
'##############################################################################################



Function strConn(ByVal serverType As String, _
                 ByVal servername As String, _
                 ByVal DBname As String) _
         As String
'==============================================================================================
' [1] FUNC: TO DEFINE THE PROPER CONNECTION STRING
'==============================================================================================
    Select Case serverType
    
       Case "RMS"
            strConn = "Provider=SQLOLEDB;" & _
                      "Data Source=" & servername & ";" & _
                      "Initial Catalog=" & DBname & ";" & _
                      "Integrated Security=SSPI;"
          
       Case "AIR"
            strConn = "Provider=SQLOLEDB;" & _
                      "Data Source=" & servername & ";" & _
                      "Initial Catalog=" & DBname & ";" & _
                      "User Id=catrader;" & _
                      "Password=catrader;"
       Case Else
            strConn = "ERROR: provided serverType is invalid"
          
    End Select

End Function
'==============================================================================================




Function sqlAIR_AvailCompany(Optional ByVal BondOnly As Boolean = False, _
                             Optional ByVal SwapOnly As Boolean = False) _
         As String
'==============================================================================================
' [2] FUNC: TO CREATE SQL QUERY TO LIST AVAILABLE COMPANY (ON THE AIR DB)
'==============================================================================================
' Input:    #BondOnly  :  optional parameter, if true the function returns the SQL query
'                         to retrive the company names relevant only to Cat Bond
'           #SwapOnly  :  optional parameter, if true the function returns the SQL query
'                         to retrive the company names relevant only to Cat Swap
' Output:   a string containing the propert SQL query to be executed on the AIR Catrader DB
' Descr:    based on the provided parameter, it returns the proper SQL query
' Vers:
' 24.07.2014 - IL
'----------------------------------------------------------------------------------------------
        If BondOnly = True And SwapOnly = False Then
            ' Defines the SQL string in the case:
            ' CAT BOND ONLY
            sqlAIR_AvailCompany = "SELECT AirCT2Exp.dbo.airfn_varbintohexstr(guidCompany) as idCompany," & _
                                           "[strName]" & _
                                   " FROM AirCT2Exp..[TblCompany]" & _
                                   " WHERE [strName] LIKE '%*%'" & _
                                   " ORDER BY [strName];"
        ElseIf BondOnly = False And SwapOnly = True Then
            ' Defines the SQL string in the case:
            ' CAT SWAP ONLY
            sqlAIR_AvailCompany = "SELECT AirCT2Exp.dbo.airfn_varbintohexstr(guidCompany) as idCompany," & _
                                           "[strName]" & _
                                   " FROM AirCT2Exp..[TblCompany]" & _
                                   " WHERE [strName] NOT LIKE '%*%'" & _
                                   " ORDER BY [strName];"
        Else
            ' Defines the SQL string in the case:
            ' ALL OTHER CASES
            sqlAIR_AvailCompany = "SELECT AirCT2Exp.dbo.airfn_varbintohexstr(guidCompany) as idCompany," & _
                                           "[strName]" & _
                                   " FROM AirCT2Exp..[TblCompany]" & _
                                   " ORDER BY [strName];"
        End If

End Function
'==============================================================================================



Function sqlRMS_AvailAnalyses() As String
'==============================================================================================
' [3] FUNC: TO CREATE SQL QUERY TO LIST AVAILABLE ANALYSES (IN THE RMS DATABASE)
'==============================================================================================
        ' Defines the SQL string
        sqlRMS_AvailAnalyses = "SELECT [ID]," & _
                                   "[NAME]," & _
                                   "[DESCRIPTION]," & _
                                   "[CURR]," & _
                                   "[PERIL]," & _
                                   "[REGION]" & _
                           " FROM [rdm_analysis];"

End Function
'==============================================================================================



Function sqlAIR_AvailContract(guidCompany As String) As String
'==============================================================================================
' [4] FUNC: TO CREATE SQL QUERY TO LIST AVAILABLE PROGRAMME/CONDITION (ON THE AIR DB)
'==============================================================================================
        ' Defines the SQL string
        sqlAIR_AvailContract = "SELECT  AirCT2Exp.dbo.airfn_varbintohexstr(a.[guidCondition]) as conditionID," & _
                                      " a.strName as conditionName," & _
                                      " b.strName as contractName," & _
                                      " c.strName as companyName" & _
                              " FROM [AirCT2Exp].[dbo].[TblCondition] as a" & _
                              " INNER JOIN " & _
                                    " [AirCT2Exp].[dbo].[TblContract] as b" & _
                                    " ON a.guidContract=b.guidContract" & _
                              " INNER JOIN " & _
                                    " [AirCT2Exp].[dbo].[TblCompany] as c" & _
                                    " ON b.guidCompany=c.guidCompany" & _
                              " WHERE c.strName = '" & guidCompany & "';"
                              
End Function
'==============================================================================================


Function sqlAIR_ALLContract() As String
'==============================================================================================
' retrieve ALL conditions
'==============================================================================================
        ' Defines the SQL string
        sqlAIR_ALLContract = "SELECT  AirCT2Exp.dbo.airfn_varbintohexstr(a.[guidCondition]) as conditionID," & _
                                      " a.strName as conditionName," & _
                                      " b.strName as contractName," & _
                                      " c.strName as companyName" & _
                              " FROM [AirCT2Exp].[dbo].[TblCondition] as a" & _
                              " INNER JOIN " & _
                                    " [AirCT2Exp].[dbo].[TblContract] as b" & _
                                    " ON a.guidContract=b.guidContract" & _
                              " INNER JOIN " & _
                                    " [AirCT2Exp].[dbo].[TblCompany] as c" & _
                                    " ON b.guidCompany=c.guidCompany"
                              '" WHERE c.strName = '" & guidCompany & "';"
                              
End Function
'==============================================================================================



Function sqlRMS_AvailProgramme() As String
'==============================================================================================
' [5] FUNC: TO CREATE SQL QUERY TO LIST AVAILABLE DATABASE (IN THE RMS DATABASE)
'==============================================================================================
        ' Defines the SQL string
        sqlRMS_AvailProgramme = "SELECT name" & _
                            " FROM sysdatabases" & _
                            " WHERE filename NOT LIKE '%MSSQL%'" & _
                            " ORDER BY name;"

End Function
'==============================================================================================



Public Function sqlRMS_QueryELT(ByVal idELT As Long, _
                                ByVal idArea As String, _
                                ByVal idPeril As String, _
                                ByVal idAnalysis As Long, _
                                ByVal perspcode As String) _
       As String
'==============================================================================================
' [6] FUNC: TO CREATE SQL QUERY TO RETRIEVE THE DESIRED RMS ELT
'==============================================================================================
Dim strSQL As String

    ' It uses the parameters 'idELT', 'idArea', 'idPeril' to define the proper sql query
    strSQL = "SELECT " & idELT & " as idELT," & _
                     "'" & idArea & "' as idArea," & _
                     "'" & idPeril & "'" & " as idPeril," & _
            "tab1.EVENTID," & _
            "tab1.RATE, " & _
            "SUM(tab1.PERSPVALUE) as PERSPVALUE," & _
            "SUM(sqrt(tab1.STDDEVI*tab1.STDDEVI)+tab1.stddevc) as STDDEVTOT," & _
            "SUM(tab1.expvalue) as EXPVALUE," & _
            "SUM(tab1.STDDEVI*tab1.STDDEVI) as STDDEVI2," & _
            "SUM(tab1.STDDEVC) As STDDEVC" & _
        " FROM"
      
    strSQL = strSQL & "(" & _
                        "SELECT A.ANLSID," & _
                                "A.[EVENTID]," & _
                                "A.[PERSPCODE]," & _
                                "A.[PERSPVALUE]," & _
                                "A.[STDDEVI]," & _
                                "A.[STDDEVC]," & _
                                "A.[EXPVALUE]," & _
                                "A.PERIL," & _
                                "B.RATE" & _
                        " FROM"
      
    strSQL = strSQL & "(" & _
                        "SELECT C.ANLSID," & _
                               "C.[EVENTID]," & _
                               "C.[PERSPCODE]," & _
                               "C.[PERSPVALUE]," & _
                               "C.[STDDEVI]," & _
                               "C.[STDDEVC]," & _
                               "C.[EXPVALUE]," & _
                               "D.[PERIL]" & _
                        " FROM [rdm_port] AS C"
      
    strSQL = strSQL & " INNER JOIN [rdm_analysis] AS D" & _
                      " ON D.ID=C.ANLSID" & _
                      " ) AS A"
      
    strSQL = strSQL & " INNER JOIN [rdm_anlsevent] AS B" & _
                      " ON A.EVENTID=B.EVENTID" & _
                      " AND a.ANLSID=b.ANLSID"
    
    ' It uses the parameter 'idAnalysis' to define the proper sql query
    strSQL = strSQL & " WHERE A.[PERSPCODE]='" & perspcode & "' AND" & _
                      "(A.[ANLSID]=" & idAnalysis & ")" & _
                      ") as tab1"
    
    strSQL = strSQL & " GROUP BY tab1.EVENTID," & _
                      "tab1.Rate"
    
    sqlRMS_QueryELT = strSQL

End Function
'==============================================================================================



' SQL query to get the first analysis into the temptable
Public Function sqlRMS_QueryELT_forHelper_createTemp(ByVal idELT As Long, ByVal idArea As String, _
                                ByVal idPeril As String, ByVal idAnalysis As Long, _
                                ByVal perspcode As String) As String
Dim strSQL As String

    strSQL = " SELECT " & idELT & " as idELT," & _
                     "'" & idArea & "' as REGION," & _
                     "'" & idPeril & "'" & " as PERIL," & _
            " tab1.EVENTID," & _
            " tab1.RATE, " & _
            " tab1.PERSPVALUE ," & _
            " tab1.expvalue ," & _
            " tab1.STDDEVI ," & _
            " tab1.STDDEVC " & _
            " INTO ##ttbl_RMSExport " & _
            " FROM " & _
            "(" & _
                "SELECT A.ANLSID," & _
                        "A.[EVENTID]," & _
                        "A.[PERSPCODE]," & _
                        "A.[PERSPVALUE]," & _
                        "A.[STDDEVI]," & _
                        "A.[STDDEVC]," & _
                        "A.[EXPVALUE]," & _
                        "A.PERIL," & _
                        "B.RATE" & _
                " FROM"

    strSQL = strSQL & "(" & _
                        "SELECT C.ANLSID," & _
                               "C.[EVENTID]," & _
                               "C.[PERSPCODE]," & _
                               "C.[PERSPVALUE]," & _
                               "C.[STDDEVI]," & _
                               "C.[STDDEVC]," & _
                               "C.[EXPVALUE]," & _
                               "D.[PERIL]" & _
                        " FROM [rdm_port] AS C"
      
    strSQL = strSQL & " INNER JOIN [rdm_analysis] AS D" & _
                      " ON D.ID=C.ANLSID" & _
                      " ) AS A"
      
    strSQL = strSQL & " INNER JOIN [rdm_anlsevent] AS B" & _
                      " ON A.EVENTID=B.EVENTID" & _
                      " AND a.ANLSID=b.ANLSID"
    
    strSQL = strSQL & " WHERE A.[PERSPCODE]='" & perspcode & "' AND" & _
                      "(A.[ANLSID]=" & idAnalysis & ")" & _
                      ") as tab1 ;"
    
    sqlRMS_QueryELT_forHelper_createTemp = strSQL

End Function
'==============================================================================================





' SQL query to get the first analysis into the temptable
Public Function sqlRMS_QueryELT_forHelper_insertTemp(ByVal idELT As Long, ByVal idArea As String, _
                                ByVal idPeril As String, ByVal idAnalysis As Long, _
                                ByVal perspcode As String) As String
Dim strSQL As String

    strSQL = "insert INTO ##ttbl_RMSExport " & _
            " SELECT " & idELT & " as idELT," & _
                     "'" & idArea & "' as REGION," & _
                     "'" & idPeril & "'" & " as PERIL," & _
            " tab1.EVENTID," & _
            " tab1.RATE, " & _
            " tab1.PERSPVALUE ," & _
            " tab1.expvalue ," & _
            " tab1.STDDEVI ," & _
            " tab1.STDDEVC " & _
            " FROM " & _
            "(" & _
                "SELECT A.ANLSID," & _
                        "A.[EVENTID]," & _
                        "A.[PERSPCODE]," & _
                        "A.[PERSPVALUE]," & _
                        "A.[STDDEVI]," & _
                        "A.[STDDEVC]," & _
                        "A.[EXPVALUE]," & _
                        "A.PERIL," & _
                        "B.RATE" & _
                " FROM"

    strSQL = strSQL & "(" & _
                        "SELECT C.ANLSID," & _
                               "C.[EVENTID]," & _
                               "C.[PERSPCODE]," & _
                               "C.[PERSPVALUE]," & _
                               "C.[STDDEVI]," & _
                               "C.[STDDEVC]," & _
                               "C.[EXPVALUE]," & _
                               "D.[PERIL]" & _
                        " FROM [rdm_port] AS C"
      
    strSQL = strSQL & " INNER JOIN [rdm_analysis] AS D" & _
                      " ON D.ID=C.ANLSID" & _
                      " ) AS A"
      
    strSQL = strSQL & " INNER JOIN [rdm_anlsevent] AS B" & _
                      " ON A.EVENTID=B.EVENTID" & _
                      " AND a.ANLSID=b.ANLSID"
    
    strSQL = strSQL & " WHERE A.[PERSPCODE]='" & perspcode & "' AND" & _
                      "(A.[ANLSID]=" & idAnalysis & ")" & _
                      ") as tab1 ;"
    
    sqlRMS_QueryELT_forHelper_insertTemp = strSQL

End Function
'==============================================================================================


Public Function sqlRMS_QueryELT_forHelper_packTemp() As String
    Dim strSQL As String
    
    strSQL = "select " & _
            " idELT, " & _
            " MAX(tab1.REGION) as REGION, " & _
            " MAX(tab1.PERIL) as PERIL, " & _
            " tab1.EVENTID, " & _
            " tab1.RATE, " & _
            " SUM(tab1.PERSPVALUE) as PERSPVALUE, " & _
            " SQRT(SUM(tab1.STDDEVI*tab1.STDDEVI))+SUM(tab1.stddevc) as STDDEVTOT, " & _
            " SUM(tab1.EXPVALUE) As EXPVALUE, " & _
            " SQRT(SUM(tab1.STDDEVI*tab1.STDDEVI)) as SQRT_STDDEVI2, " & _
            " SUM(tab1.stddevc) as STDDEVC " & _
            " into ##ttbl_RMSExport2 from ##ttbl_RMSExport as tab1  " & _
            " GROUP BY idELT, tab1.EVENTID, tab1.RATE;"
                
    sqlRMS_QueryELT_forHelper_packTemp = strSQL
            
End Function















