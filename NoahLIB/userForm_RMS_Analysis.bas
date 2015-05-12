VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userForm_RMS_Analysis 
   Caption         =   "UserForm1"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   OleObjectBlob   =   "userForm_RMS_Analysis.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "userForm_RMS_Analysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub CommandButton1_Click()
    Call getAssetList_RMS
End Sub

Private Sub CommandButton2_Click()

Dim i, nickName
Dim rg_asset_nick, rg_analysis_id As Range
Dim analysisId As Long

If MsgBox("Associate the selected analyses to the assets?", vbQuestion + vbYesNo) <> vbYes Then
    MsgBox "Canceled, exit"
    Exit Sub
End If


Set rg_asset_nick = sh1.Range("rng_RMS_LayerName")
Set rg_analysis_id = sh1.Range("rng_RMS_LayerGroup")

For i = 1 To rg_asset_nick.Rows.count
    nickName = rg_asset_nick.Cells(i, 1)
    analysisId = rg_analysis_id.Cells(i, 1)
    If modDBInterface.checkStringKeyExists("tblCatSwapLayer", "strLayerName", nickName) Then
        Call modDBInterface.updateNumValueStringKey("tblCatSwapLayer", "intRmsAnalysisToProgram", analysisId, "strLayerName", nickName)
        
    End If
Next

MsgBox "Done"

End Sub

Private Sub CommandButton3_Click()
Call clearRiskMeasuresRms
'sh1.Range("rng_RMS_LayerName").ClearContents
'sh1.Range("rng_RMS_LayerGroup").ClearContents


End Sub

'  "update buttons" button
Private Sub CommandButton5_Click()
Dim o As Object
Set o = ActiveSheet.OLEObjects("btn_RMS_SubmitOEP").Object
o.Caption = "Work with ELTs"

Call updateRange_ELAggregateRMS(ActiveSheet)

End Sub


'  "Generate YET" button
Private Sub CommandButton6_Click()

End Sub























