VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDuplicateManager 
   Caption         =   "Duplicate Manager by @ilhamrzr"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmDuplicateManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDuplicateManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private colStart As Long
Private colEnd As Long
Private rowStart As Long
Private lastRow As Long
Private dupCount As Long
Private scanned As Boolean
Private headerRow As Long

' NEW: remember the scanned workbook and sheet
Private scannedBookName As String
Private scannedSheetName As String

' ==========================================================
'  INIT + APPLY THEME + LOAD HEADER DROPDOWN
' ==========================================================
Private Sub UserForm_Initialize()

    ' Apply theme (Dark/Light)
    Call ApplyTheme

    ' Disable delete button before scan
    cmdDelete.Enabled = False
    lblSummary.Caption = ""

    ' Load header from active sheet using headerRow textbox
    Dim ws As Worksheet
    Dim lastCol As Long, c As Long

    Set ws = ActiveSheet

    ' Get headerRow from textbox
    headerRow = Val(txtHeaderRow.Text)
    If headerRow < 1 Then headerRow = 1

    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    cboHeader.Clear
    For c = 1 To lastCol
        If Trim(ws.Cells(headerRow, c).Value) <> "" Then
            cboHeader.AddItem ws.Cells(headerRow, c).Value
        End If
    Next c

    If cboHeader.ListCount > 0 Then cboHeader.ListIndex = 0
    End Sub

' ==========================================================
'  APPLY THEME (DARK / LIGHT MODE)
' ==========================================================
Public Sub ApplyTheme()

    If DM_DarkMode Then
        ' DARK MODE
        Me.BackColor = RGB(25, 25, 25)

        StyleControlDark lblTitle, True
        StyleControlDark lblHeader
        StyleControlDark lblRow
        StyleControlDark lblSummary
        StyleControlDark lblHeaderRow

        StyleButtonDark cmdScan
        StyleButtonDark cmdDelete
        StyleButtonDark cmdReset

        txtRowStart.BackColor = RGB(40, 40, 40)
        txtRowStart.ForeColor = vbWhite

        cboHeader.BackColor = RGB(40, 40, 40)
        cboHeader.ForeColor = vbWhite
        
        txtHeaderRow.BackColor = RGB(40, 40, 40)
        txtHeaderRow.ForeColor = vbWhite

    Else
        ' LIGHT MODE
        Me.BackColor = RGB(240, 240, 240)

        StyleControlLight lblTitle, True
        StyleControlLight lblHeader
        StyleControlLight lblRow
        StyleControlLight lblSummary
        StyleControlLight lblHeaderRow

        StyleButtonLight cmdScan
        StyleButtonLight cmdDelete
        StyleButtonLight cmdReset

        txtRowStart.BackColor = vbWhite
        txtRowStart.ForeColor = vbBlack

        cboHeader.BackColor = vbWhite
        cboHeader.ForeColor = vbBlack
        
        txtHeaderRow.BackColor = vbWhite
        txtHeaderRow.ForeColor = vbBlack
    End If

End Sub


' ==========================================================
'  DARK MODE STYLES
' ==========================================================
Private Sub StyleControlDark(ctrl As Control, Optional isTitle As Boolean = False)
    On Error Resume Next
    ctrl.BackColor = RGB(25, 25, 25)
    If isTitle Then
        ctrl.ForeColor = RGB(200, 230, 255)
        ctrl.Font.Bold = True
    Else
        ctrl.ForeColor = RGB(220, 220, 220)
    End If
End Sub

Private Sub StyleButtonDark(btn As MSForms.CommandButton)
    btn.BackColor = RGB(50, 50, 50)
    btn.ForeColor = vbWhite
End Sub


' ==========================================================
'  LIGHT MODE STYLES
' ==========================================================
Private Sub StyleControlLight(ctrl As Control, Optional isTitle As Boolean = False)
    On Error Resume Next
    ctrl.BackColor = RGB(240, 240, 240)
    If isTitle Then
        ctrl.ForeColor = RGB(0, 60, 130)
        ctrl.Font.Bold = True
    Else
        ctrl.ForeColor = vbBlack
    End If
End Sub

Private Sub StyleButtonLight(btn As MSForms.CommandButton)
    btn.BackColor = RGB(230, 230, 230)
    btn.ForeColor = vbBlack
End Sub


' ==========================================================
'  SCAN DUPLICATES
' ==========================================================
Private Sub cmdScan_Click()

    Dim ws As Worksheet
    Dim dict As Object
    Dim key As String
    Dim r As Long, c As Long
    Dim headerName As String
    
    If cboHeader.ListIndex < 0 Then
        MsgBox "Select the column header first.", vbExclamation
        Exit Sub
    End If
    
    Set ws = ActiveSheet
    Set dict = CreateObject("Scripting.Dictionary")
    
    headerName = cboHeader.Value
    
    ' Find column position
    colStart = 0
    For c = 1 To ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
        If CStr(ws.Cells(headerRow, c).Value) = headerName Then
            colStart = c
            Exit For
        End If
    Next c
    
    If colStart = 0 Then
        MsgBox "Header not found in row 1.", vbCritical
        Exit Sub
    End If
    
    colEnd = colStart
    
    rowStart = Val(txtRowStart.Text)
    If rowStart < 1 Then rowStart = headerRow + 1

    ' make sure it is always below the headerRow
    If rowStart <= headerRow Then
        rowStart = headerRow + 1
        txtRowStart.Text = rowStart
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, colStart).End(xlUp).Row
    If lastRow < rowStart Then
        MsgBox "No data could be scanned.", vbExclamation
        Exit Sub
    End If
    
    ' Clear old highlights
    ws.Range(ws.Cells(rowStart, 1), ws.Cells(lastRow, ws.Columns.Count)).Interior.ColorIndex = xlNone
    
    dupCount = 0
    
    ' Scan animation
    cmdScan.Caption = "Scanning..."
    cmdScan.Enabled = False
    DoEvents
    
    For r = rowStart To lastRow
        
        key = LCase(Trim(ws.Cells(r, colStart).Text))
        
        If dict.Exists(key) Then
            ws.Rows(r).Interior.Color = RGB(206, 27, 27)
            dupCount = dupCount + 1
        Else
            dict.Add key, 1
        End If
    Next r
    
    cmdScan.Caption = "Scan Duplicates"
    cmdScan.Enabled = True
    
    lblSummary.Caption = "Duplicate: " & dupCount & _
                         " | Column: " & headerName & _
                         " | Rows " & rowStart & " - " & lastRow
    scanned = True
    cmdDelete.Enabled = (dupCount > 0)
    
    ' REMEMBER workbook and sheet to delete later
    scannedBookName = ws.Parent.Name
    scannedSheetName = ws.Name
    
    If dupCount > 0 Then Beep

End Sub


' ==========================================================
'  DELETE + BACKUP
' ==========================================================
Private Sub cmdDelete_Click()

    If Not scanned Then
        MsgBox "Scan first before deleting.", vbExclamation
        Exit Sub
    End If
    
    If dupCount = 0 Then
        MsgBox "No duplicates to delete.", vbInformation
        Exit Sub
    End If

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim delDict As Object
    Dim key As String
    Dim r As Long
    Dim t As String
    Dim rowsToDelete As Collection
    Dim i As Long
    
    ' Make sure the original workbook is still open.
    On Error Resume Next
    Set wb = Workbooks(scannedBookName)
    On Error GoTo 0
    
    If wb Is Nothing Then
        MsgBox "Original workbook '" & scannedBookName & "' is not open.", _
               vbCritical, "Delete Failed"
        Exit Sub
    End If
    
    On Error Resume Next
    Set ws = wb.Worksheets(scannedSheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Original sheet '" & scannedSheetName & "' not found.", _
               vbCritical, "Delete Failed"
        Exit Sub
    End If
    
    ' === BACKUP SHEET BK_... ===
    t = Format(Now, "yymmdd_hhnnss")
    
    On Error Resume Next
    ws.Copy After:=ws
    Err.Clear
    ActiveSheet.Name = "BK_" & t
    If Err.Number <> 0 Then
        ' If the name already exists, add a random suffix
        ActiveSheet.Name = "BK_" & t & "_" & Format(Timer, "0")
        Err.Clear
    End If
    On Error GoTo 0
    
    ' === BACKUP HIDDEN FOR RESTORE ===
    Call MakeSheetBackup(ws)
    
    ' === DELETE DUPLICATES (KEEP THE FIRST LINE) ===
    Set delDict = CreateObject("Scripting.Dictionary")
    Set rowsToDelete = New Collection
    
    cmdDelete.Caption = "Deleting..."
    cmdDelete.Enabled = False
    DoEvents
    
    ' Loop from top to bottom: first appears → saved, rest → marked for delete
    For r = rowStart To lastRow
        
        key = LCase(Trim(ws.Cells(r, colStart).Text))
        
        If delDict.Exists(key) Then
            rowsToDelete.Add r
        Else
            delDict.Add key, 1
        End If
    Next r
    
    ' Delete from bottom to top so that the index does not shift.
    For i = rowsToDelete.Count To 1 Step -1
        ws.Rows(rowsToDelete(i)).Delete
    Next i
    
    ' Clean color
    ws.Cells.Interior.ColorIndex = xlNone
    ws.Cells.Font.ColorIndex = xlAutomatic
    
    cmdDelete.Caption = "Delete + Backup"
    cmdDelete.Enabled = False   ' need to rescan if you want to delete again
    scanned = False
    
    lblSummary.Caption = "Duplicate removed. Backup sheet: BK_" & t
    Beep
    
    MsgBox "Duplicates successfully removed!" & vbCrLf & _
           "Backup sheet: BK_" & t, _
           vbInformation, "Done"

End Sub


Private Sub cmdReset_Click()
    Call ResetSheet
End Sub
Private Sub ResetSheet()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Remove all interior highlights
    ws.Cells.Interior.ColorIndex = xlNone

    ' Reset font color
    ws.Cells.Font.ColorIndex = xlAutomatic

    ' Reset summary info
    lblSummary.Caption = ""

    ' Turn off the delete button
    cmdDelete.Enabled = False

    ' Reset state
    dupCount = 0
    scanned = False

    Beep
    MsgBox "Highlight has been reset.", vbInformation, "Reset Complete"

End Sub

Private Sub txtHeaderRow_Change()
    On Error Resume Next
    ' Don't call the full Show/Initialize, just reload the header
    Dim ws As Worksheet
    Dim lastCol As Long, c As Long

    Set ws = ActiveSheet

    headerRow = Val(txtHeaderRow.Text)
    If headerRow < 1 Then headerRow = 1

    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    cboHeader.Clear
    For c = 1 To lastCol
        If Trim(ws.Cells(headerRow, c).Value) <> "" Then
            cboHeader.AddItem ws.Cells(headerRow, c).Value
        End If
    Next c

    If cboHeader.ListCount > 0 Then cboHeader.ListIndex = 0
End Sub