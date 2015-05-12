VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userForm_Find 
   Caption         =   "Find.."
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   OleObjectBlob   =   "userForm_Find.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "userForm_Find"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rg_target As Range

Const ASSET_HELP = "Asset"
Const EXCEL_COMMAND_HELP = "Excel Command"
Const TRADE_HELP = "Trade"
'
'



Private Sub ComboBox1_Enter()
    ComboBox1.DropDown
End Sub





Private Sub Label1_Click()

End Sub

Private Sub TextBox1_Change()
    Dim searchString As String
    
    If Len(TextBox1.value) >= 1 Then
        searchString = Replace(TextBox1.value, " ", "%")
    
        If ComboBox1.value = ASSET_HELP Then
            Call cleanTarget_down(rg_target, 4)
            Call getAssetFind(rg_target.Cells(2, 1), rg_target.Cells(1, 1), searchString)
        ElseIf ComboBox1.value = TRADE_HELP Then
            '..
        ElseIf ComboBox1.value = EXCEL_COMMAND_HELP Then
            Call cleanTarget_down(rg_target, 8)
            Call getCommandFind(rg_target.Cells(2, 1), rg_target.Cells(1, 1), searchString)
        Else
            '..
        End If
    End If
        
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub ComboBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub UserForm_Initialize()
    Set rg_target = ActiveCell.Cells(1, 1)
    ComboBox1.AddItem ASSET_HELP
    ComboBox1.AddItem TRADE_HELP
    ComboBox1.AddItem EXCEL_COMMAND_HELP
    ComboBox1.value = EXCEL_COMMAND_HELP
    
    
    
End Sub
