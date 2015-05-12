Attribute VB_Name = "NOAH_Function"




'==============================================================================================
' GLOBAL VARIABLES
'==============================================================================================
Public wb1 As Workbook
Public sh1 As Worksheet
'==============================================================================================





'##############################################################################################
' ADDITIONAL FUNCTIONS USED WITHIN THE WORKBOOK
'   [1]  ConcatenateRange                   (to be checked)
'   [2]  ConcatenateRangeIF                 (to be reviewed)
'   [3]  CountVoidCells                     (to be reviewed)
'   [4]  Sub_RemoveValidationList           (to be reviewed)
'   [5]  Sub_CreateValidationList           (to be reviewed)
'##############################################################################################



Public Function ConcatenateRange(ByRef rng As Range) As String
'==============================================================================================
' FUNC: TO CONCATENATE CELLS
'==============================================================================================
' Input:   #rng  :  the range whose cells have to be concatenate
' Output:  a string containing all non nil values in #rng separated by semicolons
' Descr:
'----------------------------------------------------------------------------------------------

    ConcatenateRange = ""
    For Each r In rng.Cells
        If r <> "" Then
            ConcatenateRange = r & ";" & ConcatenateRange
        End If
    Next
    If ConcatenateRange <> "" Then
        ConcatenateRange = Left(ConcatenateRange, Len(ConcatenateRange) - 1)
    End If
    
End Function
'==============================================================================================



Public Function ConcatenateRangeIF(ByRef rngConc As Range, _
                            ByRef rngCond As Range, _
                            ByVal strCond As String) _
                As String
'==============================================================================================
' FUNC: TO CONCATENATE CELLS UNDER A STATED CONDITION
'==============================================================================================
' Input:   #rngConc  :  the range whose cells have to be concatenate
'          #rngCond  :  the range over which the condition #strCond has to be checked
'          #strCond  :  the condition to be checked
' Output:  a string containing the concatenated values, separated by semicolumns
' Descr:   loops over the #rngConc and concatenates all non nil values which verify #strCond
'          in range #rngCond, separated by semicolons
'----------------------------------------------------------------------------------------------
        
Dim i As Integer

    ConcatenateRangeIF = ""
    For i = 1 To rngConc.Rows.count
        If rngConc.Cells(i, 1) <> "" Then
            If Evaluate(rngCond.Cells(i, 1) & strCond) = True Then
                ConcatenateRangeIF = ConcatenateRangeIF & rngConc.Cells(i, 1) & ";"
            End If
        End If
    Next
    If ConcatenateRangeIF <> "" Then
        ConcatenateRangeIF = Left(ConcatenateRangeIF, Len(ConcatenateRangeIF) - 1)
    End If

End Function
'==============================================================================================



Public Function CountVoidCells(myRng As Range) As Integer
'==============================================================================================
' FUNC: TO COUNT VOID CELLS
'==============================================================================================

    CountVoidCells = 0
    For Each c In myRng.Cells
        If c.value = "" Then CountVoidCells = CountVoidCells + 1
    Next
    
End Function
'==============================================================================================



Sub Sub_RemoveValidationList(ByRef myWb As Workbook)
'==============================================================================================
' SUB: TO REMOVE VALIDATION LISTS
'==============================================================================================
' Input:   #wb  :  the workbook which the validation lists have to be removed from
' Output:  -
' Descr:   it deletes all the Validation Lists within the workbook since they
'          can cause an error during future opening of the file
'----------------------------------------------------------------------------------------------
Dim rngName(1 To 8) As String
Dim i As Integer

    ' lists in the Summary sheet
    rngName(1) = "rng_Currency"
    rngName(2) = "rng_Broker_House"
    ' lists in the RMS sheet
    rngName(3) = "rng_RMS_DBname"
    rngName(4) = "rng_RMS_LayerGroup"
    ' lists in the AIR sheet
    rngName(5) = "rng_AIR_CompanyName"
    rngName(6) = "rng_AIR_LayerGroup"
    ' lists in the HR sheet
    rngName(7) = "rng_HR_Bucket"
    rngName(8) = "rng_HR_Rating"

    ' Removes all the Validation Lists within the #wbName workbook
    For i = 1 To UBound(rngName)
        myWb.Activate
        With Range(myWb.Names(rngName(i))).Validation
            .Delete
        End With
    Next

End Sub
'==============================================================================================



Sub Sub_CreateValidationList(ByRef myWb As Workbook)
'==============================================================================================
' SUB: TO CREATE VALIDATION LISTS
'==============================================================================================
' Input:   #wb  :  the workbook which the validation lists have to be created in
' Output:  -
' Descr:   it creates all the Validation Lists
'----------------------------------------------------------------------------------------------
Dim shtName As String
Dim rngName As String
Dim valList As String
       
    ' SUMMARY sheet
    ' Set the validation list for the rng_Currency range
    rngName = "rng_Currency"
    valList = "AUD,BDT,CAD,CHF,CLF,CRC,DKK,DOP,EUR,GBP,INR,JPY,KRW,PGK,RMB,SAR,TWD,USD,VEF,ZAR"
    rngName = myWb.Names(rngName).Name
    myWb.Activate
    With Range(rngName).Validation
        .Delete
        .add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlEqual, _
             Formula1:=valList
    End With
    
    ' SUMMARY sheet
    ' Set the validation list for the rng_BrokerHouse range
    rngName = "rng_Broker_House"
    valList = "AGcover, AGJ-Execution, Amwins, AON, Beach, BMS, Cogent, CooperGay, GuyCarp, " & _
              "HannoverRe, ISGRP, JLT Tower, Marsh, Miller, RFIB, TigerRisk, Tullett Prebon, Usre, Willis"
    rngName = myWb.Names(rngName).Name
    myWb.Activate
    With Range(rngName).Validation
        .Delete
        .add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlEqual, _
             Formula1:=valList
    End With
    
    ' RMS sheet
    ' Set the validation list for the rng_RMS_LayerGroup range
    rngName = "rng_RMS_LayerGroup"
    valList = Range("rng_RMS_AnalysesID").Name
    rngName = myWb.Names(rngName).Name
    myWb.Activate
    With Range(rngName).Validation
        .Delete
        .add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlEqual, _
             Formula1:=valList
    End With
   
    ' AIR sheet
    ' Set the validation list for the rng_AIR_LayerGroup range
    rngName = "rng_AIR_LayerGroup"
    valList = Range("rng_AIR_AnalysesID").Name
    rngName = myWb.Names(rngName).Name
    myWb.Activate
    With Range(rngName).Validation
        .Delete
        .add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlEqual, _
             Formula1:=valList
    End With
    
    ' HR sheet
    ' Set the validation list for the rng_HRrating range
    rngName = "rng_HR_Rating"
    valList = "0,1,2"
    rngName = myWb.Names(rngName).Name
    myWb.Activate
    With Range(rngName).Validation
        .Delete
        .add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlEqual, _
             Formula1:=valList
    End With

    ' HR sheet
    ' Set the validation list for the rng_HRbucket range
    rngName = "rng_HR_Bucket"
    valList = "Africa, Asia China / Hkg Nat Cat, Asia India, Asia JP EQ 1st Event, Asia JP Wind, " & _
              "Asia Korea, Asia Middle East, Asia Other Nat Cat, Asia Taiwan Nat Cat, AUS EQ, NZ EQ, " & _
              "AUS/NZ Wind, Aviation excl. War & Terror, Aviation War & Terror, Crop US, " & _
              "Energy non-elemental, EU Eastern Europe & Austria  EQ, EU Other EQ, EU UK Flood, " & _
              "EU Wind - Northern Path 1st event (UK, F, Benelux, DE, Scandinavia, PL), " & _
              "EU Wind - Northern Path 2nd event (UK, F, Benelux, DE, Scandinavia, PL), " & _
              "EU Wind - Southern Path 1st Event (UK, F, Benelux, DE, CZ, AU), " & _
              "EU Wind - Southern Path 2nd event (UK, F, Benelux, DE, CZ, AU), " & _
              "Latin America Nat Cat Carribbean / South America, Latin America Nat Cat Carribbean / Florida, " & _
              "Latin America Nat Cat Carribbean / Texas, Latin America Nat Cat Mexico, Latin America Nat Cat Chile, Latin America Brazil, Latin America Nat Cat other South America," & _
              "North America EQ California , North America EQ Pacific NW, North America EQ Canada, " & _
              "North America EQ New Madrid, North America Hurricane US - FL to Texas + GOM + Mexico Wind 1st Event," & _
              "North America Hurricane US - FL to Texas + GOM + Mexico Wind 2nd Event," & _
              "North America Hurricane US - FL-NC 1st Event, North America Hurricane US - FL-NC 2nd Event, " & _
              "North America Hurricane US - VA-ME , North America Hurricane US - VA-ME 2nd event, " & _
              "North America Tornado, Terror - Denmark, Terror - Latin America ex Mexico, " & _
              "Terror - Latin America Mexico, Terror - UK, Terror - US, Terror - India, Terror - WW, " & _
              "Turkey Catastrophe 1st Event, Turkey Catastrophe 2ndEvent, Rest - no separate bucket"
    rngName = myWb.Names(rngName).Name
    myWb.Activate
    With Range(rngName).Validation
        .Delete
        .add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlEqual, _
             Formula1:=valList
    End With

End Sub
'==============================================================================================



Public Sub Set_FileOnOpening(myWb As Workbook, mySht As Worksheet)

End Sub


Public Sub Sub_ClearNamesFromSheet(ByRef myWb As Workbook, _
                                   ByRef mySht As Worksheet)
'==============================================================================================
' SUB: TO REMOVE ALL NAMED RANGES FROM A SHEET
'==============================================================================================
' Input:   #myWb   :  the workwook from which the named ranges should be removed
'          #mySht  :  the worksheet from which the named ranges should be removed
' Output:  -
' Descr:   it loops among all named ranges in the workbook and removes those located in the
'          worksheet provided as input
'----------------------------------------------------------------------------------------------

Dim u As Name
    
    ' Avoid allerts
    On Error Resume Next
    
    ' Loops among all named ranges
    For Each u In myWb.Names
        ' If the sheet containing the named range is equal to selected one
        If Range(u).Worksheet.Name = mySht.Name Then
            ' Removes the named range
            u.Delete
        End If
    Next
    
    ' Reintroduce allerts
    On Error GoTo 0

End Sub
'==============================================================================================
