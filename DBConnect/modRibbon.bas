Attribute VB_Name = "modRibbon"



Option Explicit


' add Menu sub
Sub AddMenus()

Dim cMenu1 As CommandBarControl
Dim cbMainMenuBar As CommandBar
Dim iHelpMenu As Integer
Dim cbcCustomMenu As CommandBarControl
Dim cbcCustomMenu2 As CommandBarControl
Dim myCBtn1 As CommandBarButton
    
    On Error Resume Next
    ' remove old menus
    Application.CommandBars("Worksheet Menu Bar").Controls("&DB Connect").Delete
    Application.CommandBars("Worksheet Menu Bar").Controls("&Bloomberg-Like").Delete

    Set cbMainMenuBar = Application.CommandBars("Worksheet Menu Bar")
    iHelpMenu = cbMainMenuBar.Controls("Help").Index

    Set cbcCustomMenu = cbMainMenuBar.Controls.Add(Type:=msoControlPopup, Before:=iHelpMenu)
    
    ' add menu
    cbcCustomMenu.Caption = "&DB Connect"
    
    
    'add buttons
    With cbcCustomMenu.Controls.Add(Type:=msoControlButton)
        .Caption = "&Open Link "
        .OnAction = "openLink_fromButton"
    End With
    With cbcCustomMenu.Controls.Add(Type:=msoControlButton)
        .Caption = "&Refresh Data"
        .OnAction = "refreshData_fromButton"
    End With
    With cbcCustomMenu.Controls.Add(Type:=msoControlButton)
        .Caption = "&Close Link"
        .OnAction = "closeLink_fromButton"
    End With
    With cbcCustomMenu.Controls.Add(Type:=msoControlButton)
        .Caption = "Fix Links"
        .OnAction = "fixLinks"
    End With
    
    With cbcCustomMenu.Controls.Add(Type:=msoControlButton)
        .Caption = "&Search/Help"
        .OnAction = "open_userForm_Find"
    End With
    
    With cbcCustomMenu.Controls.Add(Type:=msoControlButton)
        .Caption = "About"
        .OnAction = "showVersion"
    End With
    
    
'    Set myCBtn1 = myCB.Controls.Add(Type:=msoControlButton)
'    Set myCBtn1 = cbMainMenuBar.Controls.Add(Type:=msoControlButton)
'    With myCBtn1
'        .Caption = "Search/Help"
'        .OnAction = "open_userForm_Find"
'        .Picture = picPicture'
'     .Style = msoButtonCaption   '<- force caption text to show on your button
'    End With
    
    On Error GoTo 0
End Sub



