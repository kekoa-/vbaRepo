VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userForm_AIR_ImportELT 
   Caption         =   "Import ELT"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10350
   OleObjectBlob   =   "userForm_AIR_ImportELT.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "userForm_AIR_ImportELT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub btn_Exit_Click()
      
    ' Closes the ImportELT form
    Unload Me

End Sub



Private Sub btn_SubmitELT_Click()

    Call Sub_AIR_ImportELT(myWb:=wb1, _
                           mySht:=sh1, _
                           myForm:=Me)

End Sub

