VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userForm_SelectBucket 
   Caption         =   "Bucket selection"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10920
   OleObjectBlob   =   "userForm_SelectBucket.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "userForm_SelectBucket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Exit_Click()

    Unload Me

End Sub

Private Sub btn_rngDescr_Click()

    ' Calls the routine to set the proper HR bucket
    ' within the relevant sheet
    Call Sub_SetBucketHR(myForm:=Me)

End Sub
