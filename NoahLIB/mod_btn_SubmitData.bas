Attribute VB_Name = "mod_btn_SubmitData"
Option Explicit

Public Sub updateAssetPortfolio(ByRef rg As Range, ByVal intID As Long)
    If rg Is Nothing Then Exit Sub
    
    rg.Select
    
    Dim strAssetCode, strCcy As String
    Dim dblExposure As Double
    
    
    ' updates intAsset
    If checkSingleRangeType(rg.Cells(1, 2), "STR") Then
        strAssetCode = rg.Cells(1, 2)
        Call modDBInterface.updateNumValueNumKey_noDelim("tblAssetPortfolio", "intAsset", "funGetAssetId('" & strAssetCode & "')", "intID", intID)
    End If
        
    ' updates strCcy
    If checkSingleRangeType(rg.Cells(1, 5), "STR") Then
        strCcy = rg.Cells(1, 5)
        Call modDBInterface.updateStringValueNumKey("tblAssetPortfolio", "strCcy", strCcy, "intID", intID)
    End If
    
    ' updates dblExposure
    If checkSingleRangeType(rg.Cells(1, 6), "DOUBLE") Then
        dblExposure = rg.Cells(1, 6)
        Call modDBInterface.updateNumValueNumKey("tblAssetPortfolio", "dblExposure", dblExposure, "intID", intID)
    End If

    
    
End Sub

   


Public Function check_beforeUpdateMovement(ByRef rg As Range, ByVal intID As Long) As Boolean
    check_beforeUpdateMovement = False
    If rg Is Nothing Then Exit Function
    
        ' updates strCcy
    If checkSingleRangeType(rg.Cells(1, 5), "STR") Then
        strCcy = rg.Cells(1, 5)
        Call modDBInterface.updateStringValueNumKey("tblMovements", "strCcy", strCcy, "intID", intID)
    End If
    ' updates dblTradeSize
    If checkSingleRangeType(rg.Cells(1, 6), "DOUBLE") Then
        dblTradeSize = rg.Cells(1, 6)
        Call modDBInterface.updateNumValueNumKey("tblMovements", "dblTradeSize", dblTradeSize, "intID", intID)
    End If
    ' updates dblTradePrice
    If checkSingleRangeType(rg.Cells(1, 7), "DOUBLE") Then
        dblTradePrice = rg.Cells(1, 7)
        Call modDBInterface.updateNumValueNumKey("tblMovements", "dblTradePrice", dblTradePrice, "intID", intID)
    End If
    ' updates datTradeDate
    If checkSingleRangeType(rg.Cells(1, 8), "DATE") Then
        datTradeDate = rg.Cells(1, 8)
        Call modDBInterface.updateDateValueNumKey("tblMovements", "datTradeDate", datTradeDate, "intID", intID)
    End If
    ' updates datValueDate
    If checkSingleRangeType(rg.Cells(1, 9), "DATE") Then
        datValueDate = rg.Cells(1, 9)
        Call modDBInterface.updateDateValueNumKey("tblMovements", "datValueDate", datValueDate, "intID", intID)
    End If
    ' updates strBrokerageHouse
    If checkSingleRangeType(rg.Cells(1, 10), "STR") Then
        strBrokerageHouse = rg.Cells(1, 10)
        Call modDBInterface.updateStringValueNumKey("tblMovements", "strBrokerageHouse", strBrokerageHouse, "intID", intID)
    End If
    ' updates strBroker
    If checkSingleRangeType(rg.Cells(1, 11), "STR") Then
        strBroker = rg.Cells(1, 11)
        Call modDBInterface.updateStringValueNumKey("tblMovements", "strBroker", strBroker, "intID", intID)
    End If
    
    ' updates FUND
    If checkSingleRangeType(rg.Cells(1, 12), "STR") Then
        strFundName = rg.Cells(1, 12)
        Call DB_Utilities.execCommandSQL("CALL prc_MovementsSetFund(" & intID & " , '" & strFundName & "');")
    End If


End Function
 

Public Sub updateMovement(ByRef rg As Range, ByVal intID As Long)
    If rg Is Nothing Then Exit Sub
    
    rg.Select
    rg.Cells(1, 1).Interior.Color = 65535
    
    Dim strCcy, strBrokerageHouse, strBroker, strFundName As String
    Dim dblTradeSize, dblTradePrice As Double
    Dim datTradeDate, datValueDate As Date
    Dim startDate, endDate As Date
    Dim strcode As String
    
    Dim rollbackInsertion As Boolean
    strcode = rg.Cells(1, 2)
    startDate = DB_Utilities.execScalarSQL("select datstartdate from tblasset where strcode='" & strcode & "';")
    endDate = DB_Utilities.execScalarSQL("select datenddate from tblasset where strcode='" & strcode & "';")
    
    ' updates strCcy
    If checkSingleRangeType(rg.Cells(1, 5), "STR") Then
        strCcy = rg.Cells(1, 5)
        Call modDBInterface.updateStringValueNumKey("tblMovements", "strCcy", strCcy, "intID", intID)
    End If
    ' updates dblTradeSize
    If checkSingleRangeType(rg.Cells(1, 6), "DOUBLE") Then
        dblTradeSize = rg.Cells(1, 6)
        Call modDBInterface.updateNumValueNumKey("tblMovements", "dblTradeSize", dblTradeSize, "intID", intID)
    End If
    ' updates dblTradePrice
    If checkSingleRangeType(rg.Cells(1, 7), "DOUBLE") Then
        dblTradePrice = rg.Cells(1, 7)
        Call modDBInterface.updateNumValueNumKey("tblMovements", "dblTradePrice", dblTradePrice, "intID", intID)
    End If
    ' updates datTradeDate
    If checkSingleRangeType(rg.Cells(1, 8), "DATE") Then
        datTradeDate = rg.Cells(1, 8)
        Call modDBInterface.updateDateValueNumKey("tblMovements", "datTradeDate", datTradeDate, "intID", intID)
    End If
    ' updates datValueDate
    If checkSingleRangeType(rg.Cells(1, 9), "DATE") Then
        datValueDate = rg.Cells(1, 9)
        Call modDBInterface.updateDateValueNumKey("tblMovements", "datValueDate", datValueDate, "intID", intID)
    End If
    ' updates strBrokerageHouse
    If checkSingleRangeType(rg.Cells(1, 10), "STR") Then
        strBrokerageHouse = rg.Cells(1, 10)
        Call modDBInterface.updateStringValueNumKey("tblMovements", "strBrokerageHouse", strBrokerageHouse, "intID", intID)
    End If
    ' updates strBroker
    If checkSingleRangeType(rg.Cells(1, 11), "STR") Then
        strBroker = rg.Cells(1, 11)
        Call modDBInterface.updateStringValueNumKey("tblMovements", "strBroker", strBroker, "intID", intID)
    End If
    
    ' updates FUND
    If checkSingleRangeType(rg.Cells(1, 12), "STR") Then
        strFundName = rg.Cells(1, 12)
        Call DB_Utilities.execCommandSQL("CALL prc_MovementsSetFund(" & intID & " , '" & strFundName & "');")
    End If
    
    
    
    
    ' do some checks..
    If (strCcy = "") Then MsgBox "Checy strCcy"
    If (dblTradeSize = 0) Then MsgBox "Checy dblTradeSize"
    If (dblTradePrice = 0) Then MsgBox "Checy dblTradePrice"
    If (datTradeDate = 0) Then MsgBox "Checy datTradeDate"
    If (datValueDate = 0) Then MsgBox "Checy datValueDate"
    If (strFundName = "") Then MsgBox "Checy strFundName"
    If (datValueDate < datTradeDate) Then MsgBox "datValueDate must be >= than datTradeDate"
    If (datTradeDate < startDate) Then MsgBox "datTradeDate must be >= than startDate, startDate is " & Format(startDate, "yyyy-mm-dd")
    If (datTradeDate > endDate) Then MsgBox "datTradeDate must be >=  endDate, endDate is " & Format(endDate, "yyyy-mm-dd")
    
    ' do some checks..
    If (strCcy = "") Or (dblTradeSize = 0) Or (dblTradePrice = 0) Or (datTradeDate = 0) Or (datValueDate = 0) _
            Or (strFundName = "") _
            Or (datValueDate < datTradeDate) _
            Or (datTradeDate < startDate) Or (datTradeDate > endDate) Then
            ' if checks fail, then cancel the operation
            Call modDBInterface.rollbackTransaction
            rg.Cells(1, 1).Interior.Color = 255
            rg.Select
            Call MsgBox("Invalid data found, data has not been inserted.", vbExclamation)
    Else
        rg.Cells(1, 1).Interior.Color = 5296274
    
    End If
    
    
    




End Sub
