Attribute VB_Name = "modDuplicateManager"
Option Explicit

' ==========================================================
'  GLOBAL VARIABLES
' ==========================================================
Public DM_DarkMode As Boolean     ' For toggle dark/light mode


' ==========================================================
'  FULL SHEET BACKUP TO HIDDEN SHEET (Add-in Safe)
' ==========================================================
Public Sub MakeSheetBackup(ws As Worksheet)

    Dim bk As Worksheet
    On Error Resume Next
    Set bk = ThisWorkbook.Sheets("__DM_BACKUP")
    On Error GoTo 0
    
    ' If nothing exists, create hidden backup sheet
    If bk Is Nothing Then
        Set bk = ThisWorkbook.Worksheets.Add
        bk.Name = "__DM_BACKUP"
        bk.Visible = xlSheetVeryHidden
    Else
        bk.Cells.Clear
    End If
    
    ' Save original sheet & workbook name
    bk.Range("A1").Value = ws.Name                 ' Sheet name
    bk.Range("B1").Value = ws.Parent.Name          ' Workbook name  <<< FIX
    
    ' Detect size of data
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' If the sheet is completely empty, don't error.
    If lastRow < 1 Or lastCol < 1 Then Exit Sub    ' <<< SAFETY
    
    ' Store data into backup sheet starting row A2
    bk.Range("A2").Resize(lastRow, lastCol).Value = _
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value

End Sub


' ==========================================================
'  RESTORE SHEET FROM HIDDEN BACKUP
' ==========================================================
Public Sub Restore_Last_State()

    Dim bk As Worksheet
    On Error Resume Next
    Set bk = ThisWorkbook.Sheets("__DM_BACKUP")
    On Error GoTo 0
    
    If bk Is Nothing Then
        MsgBox "No backups found.", vbExclamation, "Restore Failed"
        Exit Sub
    End If
    
    Dim sheetName As String, wbName As String
    sheetName = CStr(bk.Range("A1").Value)
    wbName = CStr(bk.Range("B1").Value)
    
    If sheetName = "" Or wbName = "" Then
        MsgBox "Invalid backup data.", vbCritical, "Restore Failed"
        Exit Sub
    End If
    
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(wbName)          ' <<< make sure to go back to the same workbook
    On Error GoTo 0
    
    If wb Is Nothing Then
        MsgBox "Original workbook '" & wbName & "' is not open.", vbCritical, "Restore Failed"
        Exit Sub
    End If
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Original sheet '" & sheetName & "' not found in workbook '" & wbName & "'.", _
               vbCritical, "Restore Failed"
        Exit Sub
    End If
    
    ' Detect backup size: data starts at row 2, col 1
    Dim lastRow As Long, lastCol As Long
    lastRow = bk.Cells(bk.Rows.Count, 1).End(xlUp).Row - 1   ' minus header row
    If lastRow <= 0 Then
        MsgBox "Backup is empty.", vbExclamation, "Restore Failed"
        Exit Sub
    End If
    
    lastCol = bk.Cells(2, bk.Columns.Count).End(xlToLeft).Column
    If lastCol <= 0 Then
        MsgBox "Backup has no columns.", vbExclamation, "Restore Failed"
        Exit Sub
    End If
    
    ' Wipe sheet clean
    ws.Cells.Clear
    
    ' Restore data
    ws.Range("A1").Resize(lastRow, lastCol).Value = _
        bk.Range("A2").Resize(lastRow, lastCol).Value
    
    MsgBox "Restore successful!", vbInformation, "Restore Complete"

End Sub


' ==========================================================
'  HELPER: Column String → Number
' ==========================================================
Public Function ColStringToNumber(colStr As String) As Long
    Dim i As Long, ch As String, letters As String, result As Long
    
    colStr = UCase(Trim(colStr))
    
    For i = 1 To Len(colStr)
        ch = Mid(colStr, i, 1)
        If ch >= "A" And ch <= "Z" Then letters = letters & ch
    Next i
    
    If letters = "" Then
        ColStringToNumber = 0
        Exit Function
    End If
    
    For i = 1 To Len(letters)
        result = result * 26 + (Asc(Mid(letters, i, 1)) - 64)
    Next i
    
    ColStringToNumber = result
End Function


' ==========================================================
'  ENTRY POINT (Open UserForm)
' ==========================================================
Public Sub DuplicateManager()
    frmDuplicateManager.Show
End Sub


' ==========================================================
'  TOGGLE DARK/LIGHT MODE (CTRL+SHIFT+M)
' ==========================================================
Public Sub ToggleDarkMode()
    
    ' Flip mode
    DM_DarkMode = Not DM_DarkMode
    
    ' Update UI if form is open
    Dim frm As Object
    For Each frm In VBA.UserForms
        If TypeName(frm) = "frmDuplicateManager" Then
            On Error Resume Next
            frm.ApplyTheme      ' make sure there is this Sub in the form
            On Error GoTo 0
            Exit For
        End If
    Next frm
    
    If DM_DarkMode Then
        MsgBox "Dark Mode Enabled", vbInformation
    Else
        MsgBox "Light Mode Enabled", vbInformation
    End If

End Sub


' ==========================================================
'  REGISTER SHORTCUTS
' ==========================================================
Public Sub AutoRegisterShortcut()

    ' CTRL + SHIFT + D → Duplicate Manager
    Application.OnKey "^+D", "DuplicateManager"
    
    ' CTRL + SHIFT + R → Restore Last Backup
    Application.OnKey "^+R", "Restore_Last_State"

    ' CTRL + SHIFT + M → Toggle Theme
    Application.OnKey "^+M", "ToggleDarkMode"

End Sub
