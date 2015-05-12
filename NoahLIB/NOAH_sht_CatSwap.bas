Attribute VB_Name = "NOAH_sht_CatSwap"
'##############################################################################################
' ROUTINES, FORMS AND FUNCTIONS RELATIVE TO THE SUMMARY SHEET (CAT SWAP CASE)
'   [1]  Sub_CatSwap_SubmitData          (to be developed)
'   [2]  Sub_CatSwap_RetrieveData        (to be developed)
'   [3]  Sub_CatSwap_DeleteData          (to be developed)
'##############################################################################################



Public Sub Sub_CatSwap_SubmitData(ByRef myWb As Workbook, _
                                  ByRef mySht As Worksheet)
'==============================================================================================
' SUB: TO SUBMIT DATA INTO THE MySQL DB
'==============================================================================================

'    MsgBox "Routine 'Sub_CatSwap_SubmitData' is Under Construction"
' implementation
Call SubmitDataIntoDatabase_CatSwapProgram(myWb)

End Sub
'==============================================================================================


Public Sub Sub_CatSwap_RetrieveData(ByRef myWb As Workbook, _
                                    ByRef mySht As Worksheet)
'==============================================================================================
' SUB: TO RETRIEVE DATA INTO THE MySQL DB
'==============================================================================================

    MsgBox "Routine 'Sub_CatSwap_RetrieveData' is Under Construction"
    
End Sub
'==============================================================================================


Public Sub Sub_CatSwap_DeleteData(ByRef myWb As Workbook, _
                                  ByRef mySht As Worksheet)
'==============================================================================================
' SUB: TO DELETE DATA INTO THE MySQL DB
'==============================================================================================

'    MsgBox "Routine 'Sub_CatSwap_DeleteData' is Under Construction"

    
    'implementation
    Call DeleteFromDatabase_CatSwap_Program(myWb)

End Sub
'==============================================================================================

