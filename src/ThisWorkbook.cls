VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    AutoRegisterShortcut
End Sub

Private Sub Workbook_AddinInstall()
    AutoRegisterShortcut
End Sub

Private Sub Workbook_Activate()
    AutoRegisterShortcut
End Sub

Private Sub Workbook_WindowActivate(ByVal Wn As Window)
    AutoRegisterShortcut
End Sub

