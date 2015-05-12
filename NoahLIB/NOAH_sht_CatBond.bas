Attribute VB_Name = "NOAH_sht_CatBond"
'##############################################################################################
' ROUTINES, FORMS AND FUNCTIONS RELATIVE TO THE SUMMARY SHEET (CAT BOND CASE)
'   [1]  Sub_CatBond_SubmitData          (to be developed)
'   [2]  Sub_CatBond_RetrieveData        (to be developed)
'   [3]  Sub_CatBond_DeleteData          (to be developed)
'##############################################################################################



Public Sub Sub_CatBond_SubmitData(ByRef myWb As Workbook, _
                                  ByRef mySht As Worksheet)
'==============================================================================================
' SUB: TO SUBMIT DATA INTO THE MySQL DB
'==============================================================================================

'    MsgBox "Routine 'Sub_CatBond_SubmitData' is Under Construction"

    Call SubmitDataIntoDatabase_CatBond(myWb)
    
End Sub
'==============================================================================================


Public Sub Sub_CatBond_RetrieveData(ByRef myWb As Workbook, _
                                    ByRef mySht As Worksheet)
'==============================================================================================
' SUB: TO RETRIEVE DATA INTO THE MySQL DB
'==============================================================================================
    
    MsgBox "Routine 'Sub_CatBond_RetrieveData' is Under Construction"

End Sub
'==============================================================================================


Public Sub Sub_CatBond_DeleteData(ByRef myWb As Workbook, _
                                  ByRef mySht As Worksheet)
'==============================================================================================
' SUB: TO DELETE DATA INTO THE MySQL DB
'==============================================================================================

'    MsgBox "Routine 'Sub_CatBond_DeleteData' is Under Construction"
    ' implementation
    Call DeleteFromDatabase_CatBond(myWb)

End Sub
'==============================================================================================

