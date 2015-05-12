VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userForm_AIR_Analysis 
   Caption         =   "AIR Analysis"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   OleObjectBlob   =   "userForm_AIR_Analysis.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "userForm_AIR_Analysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






' retrieves asset list, and puts them into the appropriate range
Private Sub CommandButton1_Click()

    Call getAssetList_AIR(myWb:=ActiveWorkbook, _
                          mySht:=ActiveSheet)
    
End Sub




Private Sub CommandButton2_Click()

    Call AIR_LinkAnalysisToAsset(ActiveWorkbook, ActiveWorkbook.ActiveSheet)

End Sub

Private Sub CommandButton3_Click()
    Call clearRiskMeasuresAir(myWb:=ActiveWorkbook, _
                              mySht:=ActiveSheet)

End Sub


'estimate risk button
Private Sub CommandButton4_Click()


End Sub



Private Sub CommandButton5_Click()
    Dim o As Object
    
    Dim rg As Range
    Set o = ActiveSheet.OLEObjects("btn_AIR_SubmitOEP").Object
    o.Caption = "Work with ELTs"
    
    ' if doesn't exist, create it
    Call updateRange_ELAggregateAIR(ActiveSheet)
    Call updateRange_ExhProbAggregateAIR(ActiveSheet)
    ActiveSheet.Columns("M:O").ColumnWidth = 11
    Call updateRange_AttProbAggregateAIR(ActiveSheet)

    Call updateRange_Format_AggregateAIR(ActiveSheet)
End Sub

Private Sub CommandButton6_Click()
    userForm_AIR_Analysis.Hide
    Call Form_AIR_ImportELT_ALL(ActiveWorkbook, ActiveWorkbook.Worksheets("AIR"))
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

End Sub























