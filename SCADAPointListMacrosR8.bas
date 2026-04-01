Attribute VB_Name = "modPointsListGenerator"
' ============================================================================
' SCADA POINTS LIST GENERATOR
' ============================================================================
'
' PURPOSE:
'   Automates the creation and maintenance of a SCADA points database by
'   joining three source tables:
'
'     DEVICE_LIST  --[I/O TEMPLATE]-->  TEMPLATE_TABLE
'     TEMPLATE_TABLE  --[POINT CATEGORY]-->  CATEGORY_TABLE
'
'   Each device in Device List has an I/O TEMPLATE assignment.  That template
'   defines one or more SCADA points (trip, breaker, status, etc.) in the
'   Template Table.  Each template point has a POINT CATEGORY that maps to
'   the Alarm Categories table for HMI, alarm, and historian settings.
'
'   The output is a flat "Points List" - one row per device-point combination
'   with all attributes merged from the three source tables.
'
' ARCHITECTURE:
'   The module uses a shared-context pattern:
'     - LoadContext() reads all tables once and returns a Dictionary with
'       every array, dictionary, and column index needed by the engine.
'     - BuildExpectedArray() uses that context to generate the expected
'       output for any set of devices (new only, or all).
'     - RunCompareEngine() reuses the same expected output to compare or
'       update the existing Points List.
'   This avoids duplicating table-reading and mapping logic across the
'   four public macros.
'
' PERFORMANCE STRATEGY:
'   - All sheet reads are single bulk operations into Variant arrays.
'   - All processing happens in memory - no cell-by-cell interaction.
'   - Lookups use Scripting.Dictionary for O(1) key access.
'   - ScreenUpdating, Calculation, and Events are disabled during execution.
'   - Sheet writes are single bulk operations from pre-built arrays.
'
' COLUMN MAPPING:
'   Instead of hardcoding column positions, headers are read at runtime
'   and matched by name between the source tables and the Points List.
'   This means columns can be reordered, added, or removed in any table
'   and the code adapts automatically.  An alias system handles cases
'   where the same data has different header names across tables
'   (e.g. "TAG SUFFIX" in Points List vs "IED TAG SUFFIX" in Templates).
'
' MAIN MACROS (visible in the Macro dialog):
'   GeneratePointsList           - Append new device points to POINTS_LIST
'   ClearPointsList              - Remove all data rows from POINTS_LIST
'   ComparePointsList            - Compare existing vs expected, highlight diffs
'   FilterReviewedPoints         - Filter Points List to tags in Compare/Orphaned Points (run again to clear)
'   ShowAllRows                  - Unhide all rows and clear any active filter
'   UpdatePointsList             - Overwrite only differing cells with expected values
'   UpdateCategoryFromSelection  - Refresh category data for selected rows
'   FindOrphanedPoints           - Highlight rows whose device or template no longer exists
'   RemoveOrphanedPoints         - Delete selected orphaned rows (select on Orphaned Points sheet)
'   GenerateCompletionReport     - Report complete/incomplete devices grouped by IED AREA
'   InsertMissingPoints          - Insert rows that are in the template but not the Points List
'   BackupWorkbook               - Save a timestamped copy of this workbook
'
' ============================================================================
Option Explicit

' ---------------------------------------------------------------------------
' CONSTANTS
' ---------------------------------------------------------------------------
' All table names, sheet names, and key headers are defined here as constants
' so that if a name changes in the workbook, only one line needs updating.
' ---------------------------------------------------------------------------

' Sheet names - must match the worksheet tab names exactly
Private Const SH_DEVICES     As String = "Device List"
Private Const SH_TEMPLATES   As String = "Templates"
Private Const SH_CATEGORIES  As String = "Alarm Categories"
Private Const SH_POINTS      As String = "Points List"
Private Const SH_COMPARE     As String = "Compare"
Private Const SH_ORPHANS     As String = "Orphaned Points"
Private Const SH_COMPLETION  As String = "Completion Report"

' ListObject (Excel Table) names - must match the table names in each sheet.
' Using named tables rather than raw ranges gives us automatic expansion,
' structured references, and reliable data body access.
Private Const TBL_DEVICES    As String = "DEVICE_LIST"
Private Const TBL_TEMPLATES  As String = "TEMPLATE_TABLE"
Private Const TBL_CATEGORIES As String = "CATEGORY_TABLE"
Private Const TBL_POINTS     As String = "POINTS_LIST"

' Key column headers used for joining tables.
Private Const HDR_IO_TEMPLATE   As String = "I/O TEMPLATE"    ' Device -> Template join key
Private Const HDR_POINT_CAT     As String = "POINT CATEGORY"  ' Template -> Category join key
Private Const HDR_TAG_NAME      As String = "TAG NAME"        ' Unique identifier per point (derived)
Private Const HDR_TAG_SUFFIX    As String = "TAG SUFFIX"      ' Point suffix from template
Private Const HDR_IED_NAME      As String = "IED NAME"        ' Unique device identifier
Private Const HDR_IED_AREA      As String = "IED AREA"        ' Area/zone grouping for devices
Private Const HDR_DESCRIPTION   As String = "DESCRIPTION"     ' Point description (has placeholders)
Private Const HDR_AESO_REQ      As String = "AESO REQUIRED"   ' Always ignored in compare
Private Const HDR_AESO_DESC     As String = "AESO DESCRIPTION" ' Always ignored in compare

' HEADER ALIASES
' The Points List may use different column names than the source tables for
' the same data.  Format: semicolon-separated "TARGET_HEADER|SOURCE_HEADER" pairs.
'
' "TAG SUFFIX" in Points List maps to "IED TAG SUFFIX" in Template Table.
' "DESCRIPTION" in Points List maps to "TEMP TAG DESCRIPTION" in Template Table.
Private Const HEADER_ALIASES As String = _
    "TAG SUFFIX|IED TAG SUFFIX;" & _
    "DESCRIPTION|TEMP TAG DESCRIPTION"

' REPLACEMENT TOKENS
' The TEMP TAG DESCRIPTION field can contain literal placeholder text that is
' replaced with device-specific values at generation time.
' Listed longest-first to prevent partial matches (e.g. RPLIED before RPL1).
Private Const RPL_TOKENS As String = _
    "RPLIED;RPLBKR1;RPLBKR2;RPLBUS1;RPLBUS2;RPLTX1;RPLNGR1;RPLRLY1;RPLRLY2;" & _
    "86PL1;86PL2;86PL3;86PL4;86PL5;86PL6;" & _
    "RPLPT1;RPLPT2;RPLPT3;RPLCT1;RPLCT2;RPLCT3;RPLCT4;RPLGCT1;RPLGCT2"

' TAG NAME SEPARATOR
' TAG NAME is a derived field: IED_NAME + separator + TAG_SUFFIX
Private Const TAG_NAME_SEP As String = "_"

' HIGHLIGHT COLOURS for the Compare and Orphan functions
Private Const CLR_ROW_YELLOW As Long = 65535       ' RGB(255, 255,   0) - row has differences
Private Const CLR_CELL_RED   As Long = 8421631     ' RGB(255, 150, 150) - specific differing cell
Private Const CLR_ORPHAN     As Long = 42495        ' RGB(255, 165,   0) - orphaned row (orange)

' COMPARE ENGINE MODES
Private Const MODE_COMPARE As Long = 1  ' Read-only: highlight + log differences
Private Const MODE_UPDATE  As Long = 2  ' Read-write: overwrite differing cells + log


' ============================================================================
'  PUBLIC: GeneratePointsList
' ============================================================================
' Reads every device in DEVICE_LIST, finds matching template rows via the
' I/O TEMPLATE key, and writes combined rows into POINTS_LIST.
'
' SKIP LOGIC:
'   - Devices already in the Points List (matched by IED NAME) are skipped
'     so running Generate multiple times won't create duplicates.
'   - Devices with a blank I/O TEMPLATE are skipped (nothing to expand).
'
' New rows are APPENDED below any existing data so previously generated
' points are preserved.  Use ClearPointsList first for a full rebuild.
' ============================================================================
Public Sub GeneratePointsList()

    Dim currentStep  As String
    Dim msg          As String
    Dim ctx          As Object
    Dim dictExisting As Object
    Dim devicesAdded As Collection
    Dim loPoints     As ListObject
    Dim wsPoints     As Worksheet
    Dim outArr()     As Variant
    Dim totalRows    As Long
    Dim skippedExisting As Long
    Dim skippedNoTpl    As Long
    Dim numPtsCols   As Long
    Dim writeStartRow As Long
    Dim devItem      As Variant

    On Error GoTo ErrHandler

    ClearPointsFilter

    If Not AutoBackup() Then Exit Sub

    ' Disable screen repainting, automatic recalculation, and event firing.
    ' This is the single most impactful performance optimisation.
    PerformanceOn

    ' ------------------------------------------------------------------
    ' Load all table data, build dictionaries, and map columns.
    ' LoadContext returns Nothing if a required column is missing.
    ' ------------------------------------------------------------------
    currentStep = "GEN: Loading context"
    Set ctx = LoadContext()
    If ctx Is Nothing Then GoTo CleanExit
    If CLng(ctx("devRows")) = 0 Then MsgBox "DEVICE_LIST is empty.", vbExclamation: GoTo CleanExit
    If CLng(ctx("tmpRows")) = 0 Then MsgBox "TEMPLATE_TABLE is empty.", vbExclamation: GoTo CleanExit

    ' ------------------------------------------------------------------
    ' Build a hash set of IED NAMEs already present in the Points List.
    ' Exists() is O(1) so skipping already-generated devices is cheap.
    ' ------------------------------------------------------------------
    currentStep = "GEN: Building skip set"
    Set dictExisting = BuildExistingDeviceDict()

    ' ------------------------------------------------------------------
    ' Generate the output array containing only NEW devices.
    ' A Collection is passed to collect the names of devices that were added.
    ' ------------------------------------------------------------------
    currentStep = "GEN: Building output array"
    Set devicesAdded = New Collection
    BuildExpectedArray ctx, dictExisting, outArr, totalRows, skippedExisting, skippedNoTpl, devicesAdded

    If totalRows = 0 Then
        msg = "No new points to generate."
        If skippedExisting > 0 Then msg = msg & vbNewLine & "  Devices already in list: " & skippedExisting
        If skippedNoTpl > 0 Then msg = msg & vbNewLine & "  Devices without I/O Template: " & skippedNoTpl
        MsgBox msg, vbExclamation
        GoTo CleanExit
    End If

    ' ------------------------------------------------------------------
    ' Write the output array to the Points List sheet in one bulk operation.
    ' ------------------------------------------------------------------
    currentStep = "GEN: Writing to sheet"
    Application.StatusBar = "Writing " & totalRows & " rows to Points List..."

    Set loPoints = ThisWorkbook.Worksheets(SH_POINTS).ListObjects(TBL_POINTS)
    Set wsPoints = ThisWorkbook.Worksheets(SH_POINTS)
    numPtsCols = CLng(ctx("numPtsCols"))

    ' Determine the first empty row to write to.
    ' DataBodyRange Is Nothing (or ListRows.Count = 0) means the table is empty.
    If loPoints.DataBodyRange Is Nothing Then
        writeStartRow = loPoints.HeaderRowRange.Row + 1
    ElseIf loPoints.ListRows.Count = 0 Then
        writeStartRow = loPoints.HeaderRowRange.Row + 1
    Else
        writeStartRow = loPoints.DataBodyRange.Row + loPoints.DataBodyRange.Rows.Count
    End If

    ' Single bulk write - transfers the entire 2D array in one COM call.
    wsPoints.Range( _
        wsPoints.Cells(writeStartRow, loPoints.HeaderRowRange.Column), _
        wsPoints.Cells(writeStartRow + totalRows - 1, loPoints.HeaderRowRange.Column + numPtsCols - 1) _
    ).Value = outArr

    ' Resize the ListObject so structured references and formatting extend to new rows.
    currentStep = "GEN: Resizing table"
    loPoints.Resize wsPoints.Range( _
        loPoints.HeaderRowRange.Cells(1, 1), _
        wsPoints.Cells(writeStartRow + totalRows - 1, loPoints.HeaderRowRange.Column + numPtsCols - 1))

    ' ------------------------------------------------------------------
    ' Summary message
    ' ------------------------------------------------------------------
    msg = "Points List updated." & vbNewLine & _
          "  New points added:          " & totalRows & vbNewLine & _
          "  Devices added:             " & devicesAdded.Count
    If skippedExisting > 0 Then
        msg = msg & vbNewLine & "  Devices skipped (existing): " & skippedExisting
    End If
    If skippedNoTpl > 0 Then
        msg = msg & vbNewLine & "  Devices skipped (no I/O):   " & skippedNoTpl
    End If

    If devicesAdded.Count > 0 Then
        msg = msg & vbNewLine & vbNewLine & "Devices added:"
        For Each devItem In devicesAdded
            msg = msg & vbNewLine & "  - " & CStr(devItem)
        Next devItem
    End If

    MsgBox msg, vbInformation

CleanExit:
    Erase outArr
    Set ctx = Nothing
    Set dictExisting = Nothing
    Set devicesAdded = Nothing
    Set loPoints = Nothing
    Set wsPoints = Nothing
    PerformanceOff
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Debug.Print "*** ERROR at: " & currentStep & " | " & Err.Number & ": " & Err.Description
    MsgBox "Error at: " & currentStep & vbNewLine & vbNewLine & _
           "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit

End Sub


' ============================================================================
'  PUBLIC: ClearPointsList
' ============================================================================
' Deletes all data rows from the POINTS_LIST table, leaving headers intact.
' Requires two-step confirmation (Yes/No + typed "CLEAR") before proceeding.
' ============================================================================
Public Sub ClearPointsList()

    Dim confirmCode As String
    Dim loPoints    As ListObject

    If MsgBox("This will permanently delete all rows from the Points List." & vbNewLine & _
              "Are you sure?", vbQuestion Or vbYesNo, "Clear Points List") = vbNo Then
        Exit Sub
    End If

    confirmCode = InputBox("Type CLEAR to confirm deletion of all Points List data.", "Confirm Clear")
    If UCase(Trim(confirmCode)) <> "CLEAR" Then
        MsgBox "Operation cancelled.", vbExclamation, "Clear Points List"
        Exit Sub
    End If

    If Not AutoBackup() Then Exit Sub

    On Error GoTo ErrHandler
    ClearPointsFilter
    PerformanceOn

    Set loPoints = ThisWorkbook.Worksheets(SH_POINTS).ListObjects(TBL_POINTS)

    If Not loPoints.DataBodyRange Is Nothing Then
        If loPoints.ListRows.Count > 0 Then loPoints.DataBodyRange.Delete
        MsgBox "Points List cleared.", vbInformation
    Else
        MsgBox "Points List is already empty.", vbInformation
    End If

CleanExit:
    Set loPoints = Nothing
    PerformanceOff
    Exit Sub

ErrHandler:
    MsgBox "Error in ClearPointsList: " & Err.Description, vbCritical
    Resume CleanExit

End Sub


' ============================================================================
'  PRIVATE: GetLocalFolder
' ============================================================================
' Returns the local file-system path for the workbook's containing folder.
' When a workbook is stored on OneDrive, Office 365 sometimes returns a cloud
' URL (https://...) from ThisWorkbook.Path instead of a local path.  This
' function detects that case and attempts to map the URL back to the local
' OneDrive sync folder using the OneDrive* environment variables.
' ============================================================================
Private Function GetLocalFolder(ByVal wb As Workbook) As String

    Dim sPath     As String
    Dim sRel      As String
    Dim aRoots(2) As String
    Dim sTry      As String
    Dim sCandidate As String
    Dim nSlash    As Long
    Dim i         As Integer
    Dim j         As Integer

    sPath = wb.Path

    ' If it's already a local path (no "://"), return it unchanged.
    If InStr(1, sPath, "://") = 0 Then
        GetLocalFolder = sPath
        Exit Function
    End If

    ' Strip the protocol and host to get just the URL path.
    ' e.g. "https://d.docs.live.net/CID/Work/SCADA Macros" -> "\Work\SCADA Macros"
    sRel = Mid(sPath, InStr(sPath, "://") + 3)  ' remove "https://"
    sRel = Mid(sRel, InStr(sRel, "/"))           ' remove hostname
    sRel = Replace(sRel, "/", "\")               ' convert to backslashes

    ' Try each known OneDrive local root in turn.
    aRoots(0) = Environ("OneDriveCommercial")
    aRoots(1) = Environ("OneDrive")
    aRoots(2) = Environ("OneDriveConsumer")

    For i = 0 To 2
        If Len(aRoots(i)) > 0 Then
            ' Strip leading URL segments one at a time until Dir finds a match.
            sTry = sRel
            For j = 0 To 6
                sCandidate = aRoots(i) & sTry
                If Dir(sCandidate, vbDirectory) <> "" Then
                    GetLocalFolder = sCandidate
                    Exit Function
                End If
                nSlash = InStr(2, sTry, "\")
                If nSlash = 0 Then Exit For
                sTry = Mid(sTry, nSlash)
            Next j
        End If
    Next i

    GetLocalFolder = sPath  ' Could not resolve - return original

End Function


' ============================================================================
'  PRIVATE: AutoBackup
' ============================================================================
' Silent backup called automatically before any operation that modifies the
' Points List.  Returns True if the backup succeeded (safe to proceed) or
' False if it failed (caller should abort).
'
' Special case: if the workbook has never been saved (Path = ""), the backup
' is silently skipped and True is returned so the operation can continue.
' ============================================================================
Private Function AutoBackup() As Boolean

    Dim folder       As String
    Dim baseName     As String
    Dim ext          As String
    Dim backupFolder As String
    Dim dotPos       As Long

    AutoBackup = False

    folder = GetLocalFolder(ThisWorkbook)
    If Len(folder) = 0 Then
        AutoBackup = True   ' Unsaved workbook - skip backup
        Exit Function
    End If

    baseName = ThisWorkbook.Name
    dotPos = InStrRev(baseName, ".")
    If dotPos > 0 Then
        ext = Mid(baseName, dotPos)
        baseName = Left(baseName, dotPos - 1)
    End If

    backupFolder = folder & "\Backups"

    On Error GoTo BackupFail
    If Dir(backupFolder, vbDirectory) = "" Then MkDir backupFolder
    ThisWorkbook.SaveCopyAs backupFolder & "\" & baseName & " Backup " & _
                             Format(Now, "yyyy-mm-dd hh-mm") & ext
    AutoBackup = True
    Exit Function

BackupFail:
    MsgBox "Auto-backup failed. Operation cancelled." & vbNewLine & vbNewLine & _
           "Error #" & Err.Number & ": " & Err.Description & vbNewLine & vbNewLine & _
           "Path attempted: " & backupFolder, vbCritical, "Backup Failed"

End Function


' ============================================================================
'  PUBLIC: BackupWorkbook
' ============================================================================
' Saves a copy of the current workbook into a "Backups" subfolder in the
' same directory as the workbook, with a timestamp appended to the filename.
'
' Example: "SCADA Points.xlsm" -> "Backups\SCADA Points Backup 2026-03-07 14-35.xlsm"
'
' SaveCopyAs is used instead of SaveAs so the active workbook remains bound
' to its original path and its saved/unsaved state is unaffected.
' ============================================================================
Public Sub BackupWorkbook()

    Dim folder       As String
    Dim baseName     As String
    Dim ext          As String
    Dim ts           As String
    Dim backupFolder As String
    Dim backupName   As String
    Dim dotPos       As Long

    folder = GetLocalFolder(ThisWorkbook)
    baseName = ThisWorkbook.Name

    ' Split extension from base name so we can insert the suffix before it.
    dotPos = InStrRev(baseName, ".")
    If dotPos > 0 Then
        ext = Mid(baseName, dotPos)           ' e.g. ".xlsm"
        baseName = Left(baseName, dotPos - 1) ' e.g. "SCADA Points"
    End If

    ' Build timestamp: yyyy-mm-dd hh-mm  (hyphens replace colons for Windows filenames)
    ts = Format(Now, "yyyy-mm-dd hh-mm")

    backupFolder = folder & "\Backups"
    backupName = backupFolder & "\" & baseName & " Backup " & ts & ext

    On Error GoTo ErrHandler
    If Dir(backupFolder, vbDirectory) = "" Then MkDir backupFolder
    ThisWorkbook.SaveCopyAs backupName
    MsgBox "Backup saved:" & vbNewLine & backupName, vbInformation, "Backup Workbook"
    Exit Sub

ErrHandler:
    MsgBox "Backup failed." & vbNewLine & vbNewLine & _
           "Error #" & Err.Number & ": " & Err.Description, vbCritical, "Backup Workbook"

End Sub


' ============================================================================
'  PUBLIC: ComparePointsList
' ============================================================================
' Regenerates the FULL expected output (all devices, no skip) and compares
' it against the existing Points List.  Differences are:
'   - Highlighted in the Points List (yellow row, red cells)
'   - Logged to a "Compare" worksheet with full details
'
' This is a read-only operation - no data in the Points List is modified.
' Use UpdatePointsList to apply the corrections.
' ============================================================================
Public Sub ComparePointsList()
    ClearPointsFilter
    RunCompareEngine MODE_COMPARE
End Sub


' ============================================================================
'  PUBLIC: UpdatePointsList
' ============================================================================
' Same comparison logic as ComparePointsList, but OVERWRITES each differing
' cell with the expected value instead of highlighting it.
'
' Only cells that differ are touched - unchanged rows are left completely
' untouched.  After updating, a fresh Compare pass is run so the user can
' see what was changed (and confirm no remaining differences).
' ============================================================================
Public Sub UpdatePointsList()
    If Not AutoBackup() Then Exit Sub
    ClearPointsFilter
    RunCompareEngine MODE_UPDATE
    RunCompareEngine MODE_COMPARE
End Sub


' ============================================================================
'  PUBLIC: FilterReviewedPoints
' ============================================================================
' Toggles an AutoFilter on the Points List (POINTS_LIST table) to show only
' rows whose TAG NAME appears in the Compare sheet or the Orphaned Points sheet.
'
'   First call  - collects all tag names from the Compare and Orphaned Points
'                 sheets and applies an xlFilterValues AutoFilter on the TAG NAME
'                 column, hiding every unmatched row.
'   Second call - clears the filter and restores all rows.
'
' PREREQUISITES:
'   Run ComparePointsList to populate the Compare sheet, and/or
'   run FindOrphanedPoints to populate the Orphaned Points sheet.
'   Both sources are combined; either alone is sufficient.
' ============================================================================
Public Sub FilterReviewedPoints()

    Dim loPoints     As ListObject
    Dim wsPoints     As Worksheet
    Dim wsTest       As Worksheet
    Dim tagNameCol   As Long
    Dim dictTags     As Object
    Dim tagArr       As Variant
    Dim r            As Long
    Dim lastRow      As Long
    Dim tagVal       As String
    Dim numPtsCols   As Long
    Dim ptsHeaders() As String
    Dim c            As Long

    On Error GoTo ErrHandler

    Set wsPoints = ThisWorkbook.Worksheets(SH_POINTS)
    Set loPoints = wsPoints.ListObjects(TBL_POINTS)

    If loPoints.DataBodyRange Is Nothing Or loPoints.ListRows.Count = 0 Then
        MsgBox "Points List is empty.", vbExclamation, "Filter Points List"
        Exit Sub
    End If

    ' --- Toggle: if a filter is already active, clear it and exit ---
    If wsPoints.FilterMode Then
        wsPoints.ShowAllData
        Application.StatusBar = False
        Exit Sub
    End If

    ' --- Find TAG NAME column index (1-based, relative to table) ---
    numPtsCols = loPoints.HeaderRowRange.Columns.Count
    ReDim ptsHeaders(1 To numPtsCols)
    For c = 1 To numPtsCols
        ptsHeaders(c) = CStr(loPoints.HeaderRowRange.Cells(1, c).Value)
    Next c
    tagNameCol = FindHeader(ptsHeaders, numPtsCols, HDR_TAG_NAME)

    If tagNameCol = 0 Then
        MsgBox "TAG NAME column not found in Points List.", vbExclamation, "Filter Points List"
        Exit Sub
    End If

    ' --- Collect unique tag names from Compare and Orphaned Points sheets ---
    Set dictTags = CreateObject("Scripting.Dictionary")
    dictTags.CompareMode = 1   ' vbTextCompare - case-insensitive

    ' From Compare sheet: TAG NAME in column B, data starts at row 3
    On Error Resume Next
    Set wsTest = ThisWorkbook.Worksheets(SH_COMPARE)
    On Error GoTo ErrHandler
    If Not wsTest Is Nothing Then
        lastRow = wsTest.Cells(wsTest.Rows.Count, 2).End(xlUp).Row
        For r = 3 To lastRow
            tagVal = SafeStrVal(wsTest.Cells(r, 2).Value)
            If tagVal <> "" Then dictTags(tagVal) = 1
        Next r
        Set wsTest = Nothing
    End If

    ' From Orphaned Points sheet: TAG NAME in column B, data starts at row 3
    On Error Resume Next
    Set wsTest = ThisWorkbook.Worksheets(SH_ORPHANS)
    On Error GoTo ErrHandler
    If Not wsTest Is Nothing Then
        lastRow = wsTest.Cells(wsTest.Rows.Count, 2).End(xlUp).Row
        For r = 3 To lastRow
            tagVal = SafeStrVal(wsTest.Cells(r, 2).Value)
            If tagVal <> "" Then dictTags(tagVal) = 1
        Next r
        Set wsTest = Nothing
    End If

    If dictTags.Count = 0 Then
        MsgBox "No tag names found in the Compare or Orphaned Points sheets." & vbNewLine & _
               "Run ComparePointsList or FindOrphanedPoints first.", _
               vbExclamation, "Filter Points List"
        Exit Sub
    End If

    tagArr = dictTags.Keys

    ' --- Apply AutoFilter on the table ---
    PerformanceOn
    loPoints.Range.AutoFilter Field:=tagNameCol, _
                               Criteria1:=tagArr, _
                               Operator:=xlFilterValues
    PerformanceOff
    Application.StatusBar = False
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "Error in FilterReviewedPoints: " & Err.Description, vbCritical
    PerformanceOff

End Sub


' ============================================================================
'  PUBLIC: ShowAllRows
' ============================================================================
' Unhides all rows in the Points List and clears any active AutoFilter
' criteria, restoring the full unfiltered view.
' ============================================================================
Public Sub ShowAllRows()

    Dim loPoints As ListObject
    Dim wsPoints As Worksheet

    On Error GoTo ErrHandler

    Set wsPoints = ThisWorkbook.Worksheets(SH_POINTS)
    Set loPoints = wsPoints.ListObjects(TBL_POINTS)

    If loPoints.DataBodyRange Is Nothing Or loPoints.ListRows.Count = 0 Then
        MsgBox "Points List is empty.", vbExclamation, "Show All Rows"
        Exit Sub
    End If

    PerformanceOn

    ' Clear any active AutoFilter criteria first
    On Error Resume Next
    wsPoints.ShowAllData
    On Error GoTo ErrHandler

    ' Unhide any manually hidden rows
    loPoints.DataBodyRange.EntireRow.Hidden = False

    PerformanceOff
    Exit Sub

ErrHandler:
    MsgBox "Error in ShowAllRows: " & Err.Description, vbCritical
    PerformanceOff

End Sub


' ============================================================================
'  PRIVATE: ClearPointsFilter
' ============================================================================
' Clears any active AutoFilter on the POINTS_LIST table so that all rows are
' visible before an operation reads or modifies the table.
' ============================================================================
Private Sub ClearPointsFilter()
    Dim loPoints As ListObject
    On Error Resume Next
    Set loPoints = ThisWorkbook.Worksheets(SH_POINTS).ListObjects(TBL_POINTS)
    If Not loPoints Is Nothing Then
        If loPoints.AutoFilter.FilterMode Then loPoints.AutoFilter.ShowAllData
    End If
    On Error GoTo 0
End Sub


' ============================================================================
'  PUBLIC: InsertMissingPoints
' ============================================================================
' Inserts rows into the Points List that are expected from the current
' Device List x Template Table join but are not yet present.
'
' ORDERING:
'   Inserted rows are placed in template order.  Each missing row is anchored
'   immediately after its nearest preceding existing row.  If no preceding
'   existing row exists, it is placed just above the first following existing
'   row.  Existing rows are shifted down as needed.
'
' INSERTION STRATEGY:
'   All inserts are performed bottom-to-top so earlier row numbers are not
'   invalidated.  Consecutive missing rows sharing the same anchor are
'   inserted in reverse expected order so the final sequence is correct.
' ============================================================================
Public Sub InsertMissingPoints()

    Dim currentStep  As String
    Dim msg          As String
    Dim ctx          As Object
    Dim loPoints     As ListObject
    Dim wsPoints     As Worksheet
    Dim existArr     As Variant
    Dim existRows    As Long
    Dim ptsHeaders() As String
    Dim numPtsCols   As Long
    Dim tagNameCol   As Long
    Dim dictNone     As Object
    Dim expArr()     As Variant
    Dim expRows      As Long
    Dim dummyA       As Long
    Dim dummyB       As Long
    Dim dictExisting As Object
    Dim dataStartRow As Long
    Dim dataStartCol As Long
    Dim isMissing()  As Boolean
    Dim anchorRow()  As Long
    Dim insertCount  As Long
    Dim lastExistSheet As Long
    Dim nextExistSheet As Long
    Dim anchorArr()  As Long
    Dim expRowArr()  As Long
    Dim rowData()    As Variant
    Dim tagKey       As String
    Dim e            As Long
    Dim r            As Long
    Dim i            As Long
    Dim j            As Long
    Dim ic           As Long
    Dim c            As Long
    Dim tmpL         As Long
    Dim expRowIdx    As Long
    Dim insertAt     As Long
    Dim newLastRow   As Long
    Dim doSwap       As Boolean

    On Error GoTo ErrHandler
    ClearPointsFilter
    PerformanceOn

    ' ------------------------------------------------------------------
    ' Load all table data and build lookup structures
    ' ------------------------------------------------------------------
    currentStep = "INSERT: Loading context"
    Set ctx = LoadContext()
    If ctx Is Nothing Then GoTo CleanExit
    If CLng(ctx("devRows")) = 0 Then MsgBox "DEVICE_LIST is empty.", vbExclamation: GoTo CleanExit
    If CLng(ctx("tmpRows")) = 0 Then MsgBox "TEMPLATE_TABLE is empty.", vbExclamation: GoTo CleanExit

    ' ------------------------------------------------------------------
    ' Read existing Points List
    ' ------------------------------------------------------------------
    currentStep = "INSERT: Reading Points List"
    Set loPoints = ThisWorkbook.Worksheets(SH_POINTS).ListObjects(TBL_POINTS)
    Set wsPoints = ThisWorkbook.Worksheets(SH_POINTS)

    ReadData loPoints, existArr, existRows
    If existRows = 0 Then
        MsgBox "Points List is empty. Use GeneratePointsList first.", vbExclamation
        GoTo CleanExit
    End If

    ' ------------------------------------------------------------------
    ' Locate TAG NAME column in the Points List
    ' ------------------------------------------------------------------
    currentStep = "INSERT: Locating columns"
    ptsHeaders = ctx("ptsHeaders")
    numPtsCols = CLng(ctx("numPtsCols"))
    tagNameCol = FindHeader(ptsHeaders, numPtsCols, HDR_TAG_NAME)
    If tagNameCol = 0 Then
        MsgBox "Points List missing '" & HDR_TAG_NAME & "' column.", vbCritical
        GoTo CleanExit
    End If

    ' ------------------------------------------------------------------
    ' Build full expected array (all devices, template order, no skips)
    ' ------------------------------------------------------------------
    currentStep = "INSERT: Building expected output"
    Application.StatusBar = "Building expected output..."
    Set dictNone = CreateObject("Scripting.Dictionary")
    BuildExpectedArray ctx, dictNone, expArr, expRows, dummyA, dummyB
    If expRows = 0 Then
        MsgBox "No expected points from current source tables.", vbExclamation
        GoTo CleanExit
    End If

    ' ------------------------------------------------------------------
    ' Index existing Points List: TAG NAME -> 1-based row index in existArr
    ' ------------------------------------------------------------------
    currentStep = "INSERT: Indexing existing rows"
    Set dictExisting = CreateObject("Scripting.Dictionary")
    dictExisting.CompareMode = vbTextCompare

    For r = 1 To existRows
        tagKey = SafeStr(existArr, r, tagNameCol)
        If Len(tagKey) > 0 Then
            If Not dictExisting.Exists(tagKey) Then dictExisting.Add tagKey, r
        End If
    Next r

    ' ------------------------------------------------------------------
    ' Plan insertions.
    '
    ' Walk expArr in expected order tracking the sheet row of the last
    ' seen existing row ("lastExistSheet").  Each missing row is anchored
    ' at lastExistSheet (insert-after semantics).
    '
    ' anchorRow = 0 means no existing row precedes this missing row yet.
    ' A backward pass resolves these by anchoring one row above the first
    ' following existing row (insert-before semantics).
    ' ------------------------------------------------------------------
    currentStep = "INSERT: Planning insertions"
    dataStartRow = loPoints.HeaderRowRange.Row + 1
    dataStartCol = loPoints.HeaderRowRange.Column

    ReDim isMissing(1 To expRows)
    ReDim anchorRow(1 To expRows)
    insertCount = 0
    lastExistSheet = 0

    ' Forward pass: identify missing rows and record the preceding anchor
    For e = 1 To expRows
        tagKey = SafeStrVal(expArr(e, tagNameCol))
        If Len(tagKey) = 0 Then GoTo FwdNext
        If dictExisting.Exists(tagKey) Then
            lastExistSheet = dataStartRow + CLng(dictExisting(tagKey)) - 1
        Else
            isMissing(e) = True
            anchorRow(e) = lastExistSheet  ' 0 if no preceding existing row yet
            insertCount = insertCount + 1
        End If
FwdNext:
    Next e

    If insertCount = 0 Then
        Application.StatusBar = False
        MsgBox "No missing rows found. Points List is already complete.", vbInformation
        GoTo CleanExit
    End If

    ' Backward pass: resolve missing rows with no preceding anchor.
    nextExistSheet = 0
    For e = expRows To 1 Step -1
        tagKey = SafeStrVal(expArr(e, tagNameCol))
        If Len(tagKey) = 0 Then GoTo BwdNext
        If dictExisting.Exists(tagKey) Then
            nextExistSheet = dataStartRow + CLng(dictExisting(tagKey)) - 1
        ElseIf isMissing(e) And anchorRow(e) = 0 Then
            If nextExistSheet > 0 Then
                anchorRow(e) = nextExistSheet - 1  ' insert just above the next existing row
            Else
                anchorRow(e) = dataStartRow + existRows - 1  ' append after last data row
            End If
        End If
BwdNext:
    Next e

    ' ------------------------------------------------------------------
    ' Confirm with user before modifying any data
    ' ------------------------------------------------------------------
    Application.StatusBar = False
    msg = insertCount & " missing point(s) will be inserted into the Points List" & vbNewLine & _
          "in template order, shifting existing rows as needed." & vbNewLine & vbNewLine & _
          "Do you want to proceed?"
    If MsgBox(msg, vbQuestion Or vbYesNo, "Insert Missing Points") = vbNo Then GoTo CleanExit
    If Not AutoBackup() Then GoTo CleanExit

    ' ------------------------------------------------------------------
    ' Sort insertions descending by anchorRow, then descending by expRow.
    '
    ' Bottom-to-top order prevents earlier inserts from shifting the row
    ' numbers still needed for later inserts.
    '
    ' For identical anchors, descending expRow means the last expected row
    ' is inserted first at anchor+1, then the previous row is inserted at
    ' anchor+1 (pushing the first down), giving the correct final order.
    ' ------------------------------------------------------------------
    currentStep = "INSERT: Sorting insertion plan"
    ReDim anchorArr(1 To insertCount)
    ReDim expRowArr(1 To insertCount)
    ic = 0
    For e = 1 To expRows
        If isMissing(e) Then
            ic = ic + 1
            anchorArr(ic) = anchorRow(e)
            expRowArr(ic) = e
        End If
    Next e

    ' Selection sort - adequate for typical SCADA point counts
    For i = 1 To insertCount - 1
        For j = i + 1 To insertCount
            doSwap = False
            If anchorArr(i) < anchorArr(j) Then
                doSwap = True
            ElseIf anchorArr(i) = anchorArr(j) And expRowArr(i) < expRowArr(j) Then
                doSwap = True
            End If
            If doSwap Then
                tmpL = anchorArr(i): anchorArr(i) = anchorArr(j): anchorArr(j) = tmpL
                tmpL = expRowArr(i): expRowArr(i) = expRowArr(j): expRowArr(j) = tmpL
            End If
        Next j
    Next i

    ' ------------------------------------------------------------------
    ' Perform insertions bottom-to-top
    ' ------------------------------------------------------------------
    currentStep = "INSERT: Inserting rows"
    Application.StatusBar = "Inserting " & insertCount & " row(s)..."

    ReDim rowData(1 To 1, 1 To numPtsCols)

    For i = 1 To insertCount
        Application.StatusBar = "Inserting row " & i & " of " & insertCount & "..."
        insertAt = anchorArr(i) + 1
        expRowIdx = expRowArr(i)

        ' Insert a blank sheet row - the ListObject expands automatically
        ' when the insertion is within or adjacent to its range.
        wsPoints.Rows(insertAt).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

        ' Write expected data into the new row using a 1-row 2D array.
        For c = 1 To numPtsCols
            rowData(1, c) = expArr(expRowIdx, c)
        Next c
        wsPoints.Range( _
            wsPoints.Cells(insertAt, dataStartCol), _
            wsPoints.Cells(insertAt, dataStartCol + numPtsCols - 1) _
        ).Value = rowData
    Next i

    ' Explicitly resize the ListObject to cover all rows after insertion.
    newLastRow = dataStartRow + existRows + insertCount - 1
    loPoints.Resize wsPoints.Range( _
        loPoints.HeaderRowRange.Cells(1, 1), _
        wsPoints.Cells(newLastRow, dataStartCol + numPtsCols - 1))

    Application.StatusBar = False
    MsgBox insertCount & " row(s) inserted successfully.", vbInformation, "Insert Missing Points"

CleanExit:
    Erase expArr
    Erase isMissing
    Erase anchorRow
    On Error Resume Next
    Erase anchorArr
    Erase expRowArr
    On Error GoTo 0
    Set ctx = Nothing
    Set dictNone = Nothing
    Set dictExisting = Nothing
    Set loPoints = Nothing
    Set wsPoints = Nothing
    PerformanceOff
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Debug.Print "*** ERROR at: " & currentStep & " | " & Err.Number & ": " & Err.Description
    MsgBox "Error at: " & currentStep & vbNewLine & vbNewLine & _
           "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit

End Sub


' ============================================================================
'  PUBLIC: UpdateCategoryFromSelection
' ============================================================================
' Updates the selected rows in the Points List with current values from the
' CATEGORY_TABLE, using each row's POINT CATEGORY as the lookup key.
'
' HOW TO USE:
'   1. Select one or more rows (or cells within rows) in the Points List.
'   2. Run the macro.
'
' WHY THIS EXISTS:
'   When alarm category settings change, this function lets you surgically
'   update just the affected rows without regenerating the entire Points List.
'   It ONLY touches category-sourced columns and leaves everything else
'   (device info, template data, user-entered descriptions) untouched.
' ============================================================================
Public Sub UpdateCategoryFromSelection()

    Dim currentStep  As String
    Dim msg          As String
    Dim wsPoints     As Worksheet
    Dim loPoints     As ListObject
    Dim loCategories As ListObject
    Dim ptsHeaders() As String
    Dim numPtsCols   As Long
    Dim ptsCatCol    As Long
    Dim catHeaders() As String
    Dim catCols      As Long
    Dim catData      As Variant
    Dim catRows      As Long
    Dim dictCatHdr   As Object
    Dim catKeyCol    As Long
    Dim dictCategory As Object
    Dim catColMap()  As Long
    Dim mappedCols   As Long
    Dim dataBody     As Range
    Dim dataFirstRow As Long
    Dim dataLastRow  As Long
    Dim dataFirstCol As Long
    Dim dictSelRows  As Object
    Dim userSel      As Range
    Dim area         As Range
    Dim isect        As Range
    Dim selRowKeys   As Variant
    Dim hdr          As String
    Dim catKey       As String
    Dim ptCatVal     As String
    Dim catVal       As Variant
    Dim r            As Long
    Dim c            As Long
    Dim k            As Long
    Dim rw           As Long
    Dim catRowIdx    As Long
    Dim shRow        As Long
    Dim updatedRows  As Long
    Dim updatedCells As Long
    Dim skippedNoCat As Long
    Dim rowUpdated   As Boolean

    If Not AutoBackup() Then Exit Sub
    On Error GoTo ErrHandler
    ClearPointsFilter
    PerformanceOn

    ' ------------------------------------------------------------------
    ' STEP 1: Validate that we have a Points List table and a selection
    ' ------------------------------------------------------------------
    currentStep = "CATUPD: Getting table reference"
    Set wsPoints = ThisWorkbook.Worksheets(SH_POINTS)
    Set loPoints = wsPoints.ListObjects(TBL_POINTS)

    If loPoints.DataBodyRange Is Nothing Then
        MsgBox "Points List is empty.", vbExclamation
        GoTo CleanExit
    End If
    If loPoints.ListRows.Count = 0 Then
        MsgBox "Points List is empty.", vbExclamation
        GoTo CleanExit
    End If

    ' Capture the current selection - must be on the Points List sheet
    If ActiveSheet.Name <> wsPoints.Name Then
        MsgBox "Please select rows on the '" & SH_POINTS & "' sheet first.", vbExclamation
        GoTo CleanExit
    End If
    Set userSel = Selection

    ' ------------------------------------------------------------------
    ' STEP 2: Read Points List headers and identify the key column
    ' ------------------------------------------------------------------
    currentStep = "CATUPD: Reading headers"
    ReadHeaders loPoints, ptsHeaders, numPtsCols
    ptsCatCol = FindHeader(ptsHeaders, numPtsCols, HDR_POINT_CAT)
    If ptsCatCol = 0 Then
        MsgBox "Points List is missing '" & HDR_POINT_CAT & "' column.", vbCritical
        GoTo CleanExit
    End If

    ' ------------------------------------------------------------------
    ' STEP 3: Read Category Table and build lookup dictionary
    ' ------------------------------------------------------------------
    currentStep = "CATUPD: Reading Category Table"
    Set loCategories = ThisWorkbook.Worksheets(SH_CATEGORIES).ListObjects(TBL_CATEGORIES)
    ReadHeaders loCategories, catHeaders, catCols
    ReadData loCategories, catData, catRows
    If catRows = 0 Then
        MsgBox "CATEGORY_TABLE is empty.", vbExclamation
        GoTo CleanExit
    End If

    Set dictCatHdr = BuildHeaderDict(catHeaders, catCols)

    catKeyCol = 0
    If dictCatHdr.Exists(UCase(HDR_POINT_CAT)) Then
        catKeyCol = CLng(dictCatHdr(UCase(HDR_POINT_CAT)))
    End If
    If catKeyCol = 0 Then
        MsgBox "CATEGORY_TABLE is missing '" & HDR_POINT_CAT & "' column.", vbCritical
        GoTo CleanExit
    End If

    ' Build category lookup: POINT CATEGORY value -> row index in catData
    Set dictCategory = CreateObject("Scripting.Dictionary")
    dictCategory.CompareMode = vbTextCompare
    For r = 1 To catRows
        catKey = SafeStr(catData, r, catKeyCol)
        If Len(catKey) > 0 Then
            If Not dictCategory.Exists(catKey) Then dictCategory.Add catKey, r
        End If
    Next r

    ' ------------------------------------------------------------------
    ' STEP 4: Build column map - which Points List columns come from the
    ' Category Table?  Maps Points List column index -> Category column index.
    ' The POINT CATEGORY key column itself is skipped (we read it, not write it).
    ' ------------------------------------------------------------------
    currentStep = "CATUPD: Building category column map"
    ReDim catColMap(1 To numPtsCols)
    mappedCols = 0

    For c = 1 To numPtsCols
        hdr = UCase(Trim(ptsHeaders(c)))
        If hdr = UCase(HDR_POINT_CAT) Then
            catColMap(c) = 0  ' Skip the key column itself
        ElseIf dictCatHdr.Exists(hdr) Then
            catColMap(c) = CLng(dictCatHdr(hdr))
            mappedCols = mappedCols + 1
        Else
            catColMap(c) = 0
        End If
    Next c

    If mappedCols = 0 Then
        MsgBox "No Points List columns match Category Table headers." & vbNewLine & _
               "Nothing to update.", vbExclamation
        GoTo CleanExit
    End If

    ' ------------------------------------------------------------------
    ' STEP 5: Determine which data rows are selected.
    ' Iterates each Area in the Selection, intersects with the data body,
    ' and collects unique row numbers.
    ' ------------------------------------------------------------------
    currentStep = "CATUPD: Resolving selected rows"
    Set dataBody = loPoints.DataBodyRange
    dataFirstRow = dataBody.Row
    dataLastRow = dataBody.Row + dataBody.Rows.Count - 1
    dataFirstCol = loPoints.HeaderRowRange.Column

    Set dictSelRows = CreateObject("Scripting.Dictionary")
    For Each area In userSel.Areas
        Set isect = Nothing
        On Error Resume Next
        Set isect = Application.Intersect(area, dataBody)
        On Error GoTo ErrHandler
        If Not isect Is Nothing Then
            For rw = isect.Row To isect.Row + isect.Rows.Count - 1
                If Not dictSelRows.Exists(rw) Then dictSelRows.Add rw, True
            Next rw
        End If
    Next area

    If dictSelRows.Count = 0 Then
        MsgBox "No Points List data rows are selected." & vbNewLine & _
               "Please select rows within the Points List table.", vbExclamation
        GoTo CleanExit
    End If

    ' ------------------------------------------------------------------
    ' STEP 6: Update each selected row
    ' ------------------------------------------------------------------
    currentStep = "CATUPD: Updating selected rows"
    Application.StatusBar = "Updating category data for " & dictSelRows.Count & " rows..."

    selRowKeys = dictSelRows.Keys
    updatedRows = 0
    updatedCells = 0
    skippedNoCat = 0

    For k = 0 To UBound(selRowKeys)
        Application.StatusBar = "Updating category: row " & (k + 1) & " of " & dictSelRows.Count

        shRow = CLng(selRowKeys(k))

        ptCatVal = SafeStrVal(wsPoints.Cells(shRow, dataFirstCol + ptsCatCol - 1).Value)
        If Len(ptCatVal) = 0 Then
            skippedNoCat = skippedNoCat + 1
            GoTo NextSelRow
        End If
        If Not dictCategory.Exists(ptCatVal) Then
            skippedNoCat = skippedNoCat + 1
            GoTo NextSelRow
        End If

        catRowIdx = CLng(dictCategory(ptCatVal))
        rowUpdated = False

        For c = 1 To numPtsCols
            If catColMap(c) > 0 Then
                catVal = catData(catRowIdx, catColMap(c))
                If Not IsError(catVal) Then
                    wsPoints.Cells(shRow, dataFirstCol + c - 1).Value = catVal
                    updatedCells = updatedCells + 1
                    rowUpdated = True
                End If
            End If
        Next c

        If rowUpdated Then updatedRows = updatedRows + 1

NextSelRow:
    Next k

    ' ------------------------------------------------------------------
    ' STEP 7: Summary
    ' ------------------------------------------------------------------
    msg = "Category update complete." & vbNewLine & _
          "  Rows selected:             " & dictSelRows.Count & vbNewLine & _
          "  Rows updated:              " & updatedRows & vbNewLine & _
          "  Cells updated:             " & updatedCells
    If skippedNoCat > 0 Then
        msg = msg & vbNewLine & "  Rows skipped (no/unmatched category): " & skippedNoCat
    End If
    MsgBox msg, vbInformation

CleanExit:
    Set dictCategory = Nothing
    Set dictCatHdr = Nothing
    Set dictSelRows = Nothing
    Set loPoints = Nothing
    Set loCategories = Nothing
    Set wsPoints = Nothing
    Set dataBody = Nothing
    PerformanceOff
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Debug.Print "*** ERROR at: " & currentStep & " | " & Err.Number & ": " & Err.Description
    MsgBox "Error at: " & currentStep & vbNewLine & vbNewLine & _
           "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit

End Sub


' ============================================================================
'  PUBLIC: FindOrphanedPoints
' ============================================================================
' Scans the Points List for rows that no longer have a valid source in the
' current Device List and Template Table.  Two orphan types are detected:
'
'   TYPE 1 - Device removed:
'     The IED NAME row does not match any device in the Device List.
'
'   TYPE 2 - Template point removed:
'     The IED NAME is still in the Device List, but the TAG NAME is not
'     produced by the current Device List x Template Table join.
'
' Orphaned rows are highlighted orange and a report is written to the
' "Orphaned Points" sheet.  This is a READ-ONLY diagnostic.
' ============================================================================
Public Sub FindOrphanedPoints()

    Dim currentStep     As String
    Dim msg             As String
    Dim ctx             As Object
    Dim loPoints        As ListObject
    Dim wsPoints        As Worksheet
    Dim wsOrphans       As Worksheet
    Dim existArr        As Variant
    Dim existRows       As Long
    Dim ptsHeaders()    As String
    Dim numPtsCols      As Long
    Dim iedNameCol      As Long
    Dim tagNameCol      As Long
    Dim tagSuffixCol    As Long
    Dim devData         As Variant
    Dim devRows         As Long
    Dim devNameCol      As Long
    Dim dictDevices     As Object
    Dim dictNone        As Object
    Dim dictExpected    As Object
    Dim expArr()        As Variant
    Dim expRows         As Long
    Dim dummyA          As Long
    Dim dummyB          As Long
    Dim logArr()        As Variant
    Dim writeLog()      As Variant
    Dim dataStartRow    As Long
    Dim dataStartCol    As Long
    Dim logRows         As Long
    Dim orphanByDevice  As Long
    Dim orphanByTemplate As Long
    Dim tagKey          As String
    Dim iedName         As String
    Dim tagName         As String
    Dim reason          As String
    Dim dn              As String
    Dim r               As Long
    Dim lr              As Long

    On Error GoTo ErrHandler
    ClearPointsFilter
    PerformanceOn

    ' ------------------------------------------------------------------
    ' Load all table data and build lookup structures
    ' ------------------------------------------------------------------
    currentStep = "ORPHAN: Loading context"
    Set ctx = LoadContext()
    If ctx Is Nothing Then GoTo CleanExit

    ' ------------------------------------------------------------------
    ' Read existing Points List
    ' ------------------------------------------------------------------
    currentStep = "ORPHAN: Reading Points List"
    Set loPoints = ThisWorkbook.Worksheets(SH_POINTS).ListObjects(TBL_POINTS)
    ReadData loPoints, existArr, existRows
    If existRows = 0 Then
        MsgBox "Points List is empty.", vbExclamation
        GoTo CleanExit
    End If

    ' ------------------------------------------------------------------
    ' Locate required columns in the Points List
    ' ------------------------------------------------------------------
    currentStep = "ORPHAN: Locating columns"
    ptsHeaders = ctx("ptsHeaders")
    numPtsCols = CLng(ctx("numPtsCols"))
    iedNameCol = FindHeader(ptsHeaders, numPtsCols, HDR_IED_NAME)
    tagNameCol = FindHeader(ptsHeaders, numPtsCols, HDR_TAG_NAME)
    tagSuffixCol = FindHeader(ptsHeaders, numPtsCols, HDR_TAG_SUFFIX)

    If iedNameCol = 0 Then
        MsgBox "Points List missing '" & HDR_IED_NAME & "' column.", vbCritical
        GoTo CleanExit
    End If
    If tagNameCol = 0 Then
        MsgBox "Points List missing '" & HDR_TAG_NAME & "' column.", vbCritical
        GoTo CleanExit
    End If

    ' ------------------------------------------------------------------
    ' Build hash set of valid IED NAMEs from Device List
    ' ------------------------------------------------------------------
    currentStep = "ORPHAN: Building device set"
    devData = ctx("devData")
    devRows = CLng(ctx("devRows"))
    devNameCol = CLng(ctx("devNameCol"))

    Set dictDevices = CreateObject("Scripting.Dictionary")
    dictDevices.CompareMode = vbTextCompare
    For r = 1 To devRows
        dn = SafeStr(devData, r, devNameCol)
        If Len(dn) > 0 Then
            If Not dictDevices.Exists(dn) Then dictDevices.Add dn, True
        End If
    Next r

    ' ------------------------------------------------------------------
    ' Build hash set of all expected TAG NAMEs from the full generate output.
    ' ------------------------------------------------------------------
    currentStep = "ORPHAN: Building expected point set"
    Application.StatusBar = "Building expected points..."

    Set dictNone = CreateObject("Scripting.Dictionary")
    BuildExpectedArray ctx, dictNone, expArr, expRows, dummyA, dummyB

    Set dictExpected = CreateObject("Scripting.Dictionary")
    dictExpected.CompareMode = vbTextCompare
    If expRows > 0 Then
        For r = 1 To expRows
            tagKey = SafeStrVal(expArr(r, tagNameCol))
            If Len(tagKey) > 0 Then
                If Not dictExpected.Exists(tagKey) Then dictExpected.Add tagKey, True
            End If
        Next r
    End If

    ' ------------------------------------------------------------------
    ' Scan Points List for orphaned rows
    ' ------------------------------------------------------------------
    currentStep = "ORPHAN: Scanning Points List"
    Application.StatusBar = "Scanning for orphaned points..."

    Set wsPoints = ThisWorkbook.Worksheets(SH_POINTS)
    dataStartRow = loPoints.HeaderRowRange.Row + 1
    dataStartCol = loPoints.HeaderRowRange.Column

    ' Clear any previous highlights before scanning
    If Not loPoints.DataBodyRange Is Nothing Then
        loPoints.DataBodyRange.Interior.ColorIndex = xlNone
    End If

    ' Pre-allocate log at maximum possible size; trim later before writing.
    ReDim logArr(1 To existRows, 1 To 4)  ' IED NAME | TAG NAME | TAG SUFFIX | Reason
    logRows = 0
    orphanByDevice = 0
    orphanByTemplate = 0

    For r = 1 To existRows
        iedName = SafeStr(existArr, r, iedNameCol)
        tagName = SafeStr(existArr, r, tagNameCol)

        ' Skip rows with blank IED NAME or TAG NAME - can't check them
        If Len(iedName) = 0 Or Len(tagName) = 0 Then GoTo NextOrphanRow

        reason = ""
        If Not dictDevices.Exists(iedName) Then
            reason = "Device not in Device List"
            orphanByDevice = orphanByDevice + 1
        ElseIf Not dictExpected.Exists(tagName) Then
            reason = "Template point not found"
            orphanByTemplate = orphanByTemplate + 1
        End If

        If Len(reason) > 0 Then
            wsPoints.Range( _
                wsPoints.Cells(dataStartRow + r - 1, dataStartCol), _
                wsPoints.Cells(dataStartRow + r - 1, dataStartCol + numPtsCols - 1) _
            ).Interior.Color = CLR_ORPHAN

            logRows = logRows + 1
            logArr(logRows, 1) = iedName
            logArr(logRows, 2) = tagName
            If tagSuffixCol > 0 Then logArr(logRows, 3) = SafeStr(existArr, r, tagSuffixCol)
            logArr(logRows, 4) = reason
        End If

NextOrphanRow:
    Next r

    ' ------------------------------------------------------------------
    ' Write results to the Orphaned Points sheet
    ' ------------------------------------------------------------------
    currentStep = "ORPHAN: Writing results sheet"
    Set wsOrphans = GetOrCreateSheet(SH_ORPHANS)
    wsOrphans.Cells.Clear

    wsOrphans.Cells(2, 1).Value = "IED NAME"
    wsOrphans.Cells(2, 2).Value = "TAG NAME"
    wsOrphans.Cells(2, 3).Value = "TAG SUFFIX"
    wsOrphans.Cells(2, 4).Value = "Reason"
    With wsOrphans.Range("A2:D2")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    If logRows > 0 Then
        ReDim writeLog(1 To logRows, 1 To 4)
        For lr = 1 To logRows
            writeLog(lr, 1) = logArr(lr, 1)
            writeLog(lr, 2) = logArr(lr, 2)
            writeLog(lr, 3) = logArr(lr, 3)
            writeLog(lr, 4) = logArr(lr, 4)
        Next lr
        wsOrphans.Range("A3").Resize(logRows, 4).Value = writeLog
    End If

    wsOrphans.Columns("A:D").AutoFit

    ' ------------------------------------------------------------------
    ' Summary message
    ' ------------------------------------------------------------------
    Application.StatusBar = False
    If logRows = 0 Then
        MsgBox "No orphaned points found." & vbNewLine & _
               "Points List is consistent with the current source tables.", _
               vbInformation, "Find Orphaned Points"
    Else
        msg = "Orphaned points found: " & logRows & vbNewLine & _
              "  Device not in Device List:  " & orphanByDevice & vbNewLine & _
              "  Template point not found:   " & orphanByTemplate & vbNewLine & vbNewLine & _
              "Orphaned rows highlighted orange in the Points List." & vbNewLine & _
              "Details written to '" & SH_ORPHANS & "' sheet."
        MsgBox msg, vbInformation, "Find Orphaned Points"
    End If

CleanExit:
    Erase expArr
    Erase logArr
    Set ctx = Nothing
    Set dictNone = Nothing
    Set dictExpected = Nothing
    Set dictDevices = Nothing
    Set loPoints = Nothing
    Set wsPoints = Nothing
    PerformanceOff
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Debug.Print "*** ERROR at: " & currentStep & " | " & Err.Number & ": " & Err.Description
    MsgBox "Error at: " & currentStep & vbNewLine & vbNewLine & _
           "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit

End Sub


' ============================================================================
'  PUBLIC: RemoveOrphanedPoints
' ============================================================================
' Permanently deletes the Points List rows corresponding to the rows selected
' on the Orphaned Points sheet, then removes those rows from the report sheet.
'
' HOW TO USE:
'   1. Run FindOrphanedPoints to populate the Orphaned Points sheet.
'   2. Select one or more data rows on that sheet (header excluded).
'   3. Run this macro.
'
' MATCHING:
'   Selected rows are matched to Points List rows by TAG NAME (column 2 of
'   the Orphaned Points sheet).  Only rows whose TAG NAME exists in the
'   Points List are deleted; any stale report rows are still removed from
'   the Orphaned Points sheet.
'
' SAFETY:
'   - Requires two-step confirmation (Yes/No + AutoBackup) before any data
'     is modified.
'   - The default button on the confirmation dialog is "No" to prevent
'     accidental deletion.
' ============================================================================
Public Sub RemoveOrphanedPoints()

    ' The Orphaned Points sheet has a fixed column layout (written by FindOrphanedPoints):
    '   Row 1 = header,  Col 1 = IED NAME,  Col 2 = TAG NAME,  Col 3 = TAG SUFFIX,  Col 4 = Reason
    Const ORPHAN_TAG_NAME_COL As Long = 2
    Const ORPHAN_HEADER_ROWS  As Long = 2   ' row 1 reserved for buttons; row 2 is the header

    Dim currentStep       As String
    Dim msg               As String
    Dim wsOrphans         As Worksheet
    Dim wsPoints          As Worksheet
    Dim loPoints          As ListObject
    Dim userSel           As Range
    Dim area              As Range
    Dim orphanDeleteRange As Range
    Dim ptsHeaders()      As String
    Dim ptsData           As Variant
    Dim dictTagNames      As Object
    Dim tagKey            As String
    Dim numPtsCols        As Long
    Dim tagNameCol        As Long
    Dim ptsRows           As Long
    Dim deleteCount       As Long
    Dim notFoundCount     As Long
    Dim shownCount        As Long
    Dim rw                As Long
    Dim r                 As Long
    Dim tagItem           As Variant

    On Error GoTo ErrHandler
    ClearPointsFilter

    ' ------------------------------------------------------------------
    ' STEP 1: Validate that the selection is on the Orphaned Points sheet
    ' ------------------------------------------------------------------
    currentStep = "RMORPHAN: Validating selection"

    On Error Resume Next
    Set wsOrphans = ThisWorkbook.Worksheets(SH_ORPHANS)
    On Error GoTo ErrHandler

    If wsOrphans Is Nothing Then
        MsgBox "The '" & SH_ORPHANS & "' sheet does not exist." & vbNewLine & _
               "Run FindOrphanedPoints first.", vbExclamation, "Remove Orphaned Points"
        Exit Sub
    End If

    If ActiveSheet.Name <> wsOrphans.Name Then
        MsgBox "Please select rows on the '" & SH_ORPHANS & "' sheet first.", _
               vbExclamation, "Remove Orphaned Points"
        Exit Sub
    End If

    Set userSel = Selection

    ' ------------------------------------------------------------------
    ' STEP 2: Extract TAG NAMEs from the selected data rows.
    ' We use a Dictionary as a hash set - Exists() is O(1) and prevents
    ' duplicates if the user selects the same row multiple times.
    ' ------------------------------------------------------------------
    currentStep = "RMORPHAN: Reading selected TAG NAMEs"

    Set dictTagNames = CreateObject("Scripting.Dictionary")
    dictTagNames.CompareMode = vbTextCompare

    For Each area In userSel.Areas
        For rw = area.Row To area.Row + area.Rows.Count - 1
            If rw <= ORPHAN_HEADER_ROWS Then GoTo NextOrphanSelRow  ' skip header
            tagKey = SafeStrVal(wsOrphans.Cells(rw, ORPHAN_TAG_NAME_COL).Value)
            If Len(tagKey) > 0 Then
                If Not dictTagNames.Exists(tagKey) Then
                    dictTagNames.Add tagKey, True
                End If
            End If
NextOrphanSelRow:
        Next rw
    Next area

    If dictTagNames.Count = 0 Then
        MsgBox "No valid TAG NAMEs found in the selection." & vbNewLine & _
               "Select data rows (not the header) on the '" & SH_ORPHANS & "' sheet.", _
               vbExclamation, "Remove Orphaned Points"
        Exit Sub
    End If

    ' ------------------------------------------------------------------
    ' STEP 3: Confirm with the user - default button is No to prevent
    ' accidental deletion.
    ' ------------------------------------------------------------------
    msg = dictTagNames.Count & " orphaned point(s) will be permanently deleted from the Points List:"
    shownCount = 0
    For Each tagItem In dictTagNames.Keys
        If shownCount < 20 Then
            msg = msg & vbNewLine & "  - " & CStr(tagItem)
        ElseIf shownCount = 20 Then
            msg = msg & vbNewLine & "  ... and " & (dictTagNames.Count - 20) & " more"
        End If
        shownCount = shownCount + 1
    Next tagItem
    msg = msg & vbNewLine & vbNewLine & "This cannot be undone. Proceed?"

    If MsgBox(msg, vbQuestion Or vbYesNo Or vbDefaultButton2, "Remove Orphaned Points") = vbNo Then
        Exit Sub
    End If

    If Not AutoBackup() Then Exit Sub

    ' ------------------------------------------------------------------
    ' STEP 4: Find and delete matching rows in the Points List.
    ' We walk bottom-to-top using ListRows(r).Delete so that row indices
    ' above the deletion point are not shifted during the loop.
    ' ------------------------------------------------------------------
    currentStep = "RMORPHAN: Reading Points List"
    PerformanceOn

    Set wsPoints = ThisWorkbook.Worksheets(SH_POINTS)
    Set loPoints = wsPoints.ListObjects(TBL_POINTS)

    If loPoints.DataBodyRange Is Nothing Or loPoints.ListRows.Count = 0 Then
        MsgBox "Points List is empty.", vbExclamation
        GoTo CleanExit
    End If

    ReadHeaders loPoints, ptsHeaders, numPtsCols
    tagNameCol = FindHeader(ptsHeaders, numPtsCols, HDR_TAG_NAME)
    If tagNameCol = 0 Then
        MsgBox "Points List missing '" & HDR_TAG_NAME & "' column.", vbCritical
        GoTo CleanExit
    End If

    ReadData loPoints, ptsData, ptsRows

    currentStep = "RMORPHAN: Deleting rows from Points List"
    Application.StatusBar = "Removing " & dictTagNames.Count & " point(s) from Points List..."

    deleteCount = 0
    notFoundCount = 0

    ' Walk bottom-to-top so ListRows indices above the current row remain valid.
    For r = ptsRows To 1 Step -1
        tagKey = SafeStr(ptsData, r, tagNameCol)
        If Len(tagKey) > 0 Then
            If dictTagNames.Exists(tagKey) Then
                loPoints.ListRows(r).Delete
                deleteCount = deleteCount + 1
                dictTagNames.Remove tagKey  ' mark as found
            End If
        End If
    Next r

    ' Any keys still in dictTagNames were not found in the Points List
    notFoundCount = dictTagNames.Count

    ' ------------------------------------------------------------------
    ' STEP 5: Remove the corresponding rows from the Orphaned Points sheet.
    ' Build a Union of all rows whose TAG NAME was in the selection,
    ' then delete in a single operation.
    ' ------------------------------------------------------------------
    currentStep = "RMORPHAN: Updating Orphaned Points sheet"
    Application.StatusBar = "Updating Orphaned Points sheet..."

    For Each area In userSel.Areas
        For rw = area.Row To area.Row + area.Rows.Count - 1
            If rw <= ORPHAN_HEADER_ROWS Then GoTo NextOrphanDelRow  ' skip header
            ' Only remove rows whose TAG NAME was actually found (or attempted) -
            ' i.e. rows where TAG NAME was non-blank.  Rows with blank TAG NAME
            ' were never added to dictTagNames and are left in place.
            tagKey = SafeStrVal(wsOrphans.Cells(rw, ORPHAN_TAG_NAME_COL).Value)
            If Len(tagKey) > 0 Then
                If orphanDeleteRange Is Nothing Then
                    Set orphanDeleteRange = wsOrphans.Rows(rw)
                Else
                    Set orphanDeleteRange = Application.Union(orphanDeleteRange, wsOrphans.Rows(rw))
                End If
            End If
NextOrphanDelRow:
        Next rw
    Next area

    If Not orphanDeleteRange Is Nothing Then orphanDeleteRange.Delete

    ' ------------------------------------------------------------------
    ' STEP 6: Summary
    ' ------------------------------------------------------------------
    Application.StatusBar = False
    msg = "Removal complete." & vbNewLine & _
          "  Points deleted from Points List:  " & deleteCount
    If notFoundCount > 0 Then
        msg = msg & vbNewLine & _
              "  TAG NAMEs not found in Points List (already removed?):  " & notFoundCount
    End If
    MsgBox msg, vbInformation, "Remove Orphaned Points"

CleanExit:
    Set dictTagNames = Nothing
    Set loPoints = Nothing
    Set wsPoints = Nothing
    Set wsOrphans = Nothing
    Set orphanDeleteRange = Nothing
    PerformanceOff
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Debug.Print "*** ERROR at: " & currentStep & " | " & Err.Number & ": " & Err.Description
    MsgBox "Error at: " & currentStep & vbNewLine & vbNewLine & _
           "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit

End Sub


' ============================================================================
'  PUBLIC: GenerateCompletionReport
' ============================================================================
' Scans the Points List and classifies each device as Complete (every point
' has a non-blank DESCRIPTION) or Incomplete (one or more points are blank).
' IED AREA is sourced from the Device List (authoritative); the Points List
' IED AREA column is used only as a fallback for devices not in Device List.
'
' OUTPUT SHEET ("Completion Report"):
'   Row 1  - reserved for buttons
'   Row 2  - "SUMMARY BY IED AREA" title
'   Row 3  - Summary header: IED AREA | Complete | Incomplete | Total | %
'   Row 4+ - Summary rows (green/amber/red) + TOTAL footer
'   Gap row
'   Next   - Side-by-side device detail:
'              Left  cols 1-3: COMPLETED DEVICES   (IED NAME | IED AREA | Points)
'              Gap   col  4:   empty
'              Right cols 5-8: INCOMPLETE DEVICES  (IED NAME | IED AREA | Described | Total)
' ============================================================================
Public Sub GenerateCompletionReport()

    ' Layout constants
    Const LEFT_COL  As Long = 1   ' completed list starts here
    Const RIGHT_COL As Long = 5   ' incomplete list starts here (col 4 = gap)
    Const FAR_COL   As Long = 10  ' zero-point list starts here (col 9 = gap)

    Dim currentStep         As String
    Dim loPoints            As ListObject
    Dim loDevices           As ListObject
    Dim wsPoints            As Worksheet
    Dim wsDevices           As Worksheet
    Dim wsReport            As Worksheet
    Dim ptsHeaders()        As String
    Dim devHeaders()        As String
    Dim ptsData             As Variant
    Dim devData             As Variant
    Dim ptsRows             As Long
    Dim devRows             As Long
    Dim numPtsCols          As Long
    Dim numDevCols          As Long
    Dim iedNameCol          As Long
    Dim descCol             As Long
    Dim ptAreaCol           As Long   ' Points List IED AREA (fallback)
    Dim devIedNameCol       As Long
    Dim devAreaCol          As Long
    Dim dictDevTotal        As Object  ' IED NAME -> total point count
    Dim dictDevFilled       As Object  ' IED NAME -> points with description filled
    Dim dictDevArea         As Object  ' IED NAME -> IED AREA (from Device List)
    Dim dictAreaComplete    As Object  ' IED AREA -> complete device count
    Dim dictAreaIncomplete  As Object  ' IED AREA -> incomplete device count
    Dim areaKeys            As Variant
    Dim devNames            As Variant
    Dim devName             As Variant
    Dim areaKey             As Variant
    Dim completedNames()    As String
    Dim completedAreas()    As String
    Dim completedPts()      As Long
    Dim incompleteNames()   As String
    Dim incompleteAreas()   As String
    Dim incompleteFilled()  As Long
    Dim incompleteTotals()  As Long
    Dim leftArr()           As Variant
    Dim rightArr()          As Variant
    Dim iedName             As String
    Dim descVal             As String
    Dim areaVal             As String
    Dim area                As String
    Dim tmpS                As String
    Dim tmpL                As Long
    Dim r                   As Long
    Dim i                   As Long
    Dim j                   As Long
    Dim writeRow            As Long
    Dim detailStartRow      As Long
    Dim completeCount       As Long
    Dim incompleteCount     As Long
    Dim areaTotal           As Long
    Dim totalComplete       As Long
    Dim totalIncomplete     As Long
    Dim totalDevices        As Long
    Dim completedCount      As Long
    Dim incompletedCount    As Long
    Dim detailRows          As Long
    Dim pctComplete         As Double
    Dim doSwap              As Boolean
    Dim dictAreaZero        As Object  ' IED AREA -> zero-point device count
    Dim zeroNames()         As String
    Dim zeroAreas()         As String
    Dim zeroArr()           As Variant
    Dim totalZero           As Long
    Dim zeroCount           As Long
    Dim zeroAreaCount       As Long

    On Error GoTo ErrHandler
    ClearPointsFilter
    PerformanceOn

    ' ------------------------------------------------------------------
    ' Read Device List → seed all known devices with 0 counts and IED AREA
    ' Devices only in the Device List (no points yet) will appear as incomplete.
    ' ------------------------------------------------------------------
    currentStep = "COMPLETION: Reading Device List"
    Set dictDevArea   = CreateObject("Scripting.Dictionary")
    Set dictDevTotal  = CreateObject("Scripting.Dictionary")
    Set dictDevFilled = CreateObject("Scripting.Dictionary")
    dictDevArea.CompareMode   = vbTextCompare
    dictDevTotal.CompareMode  = vbTextCompare
    dictDevFilled.CompareMode = vbTextCompare

    On Error Resume Next
    Set wsDevices = ThisWorkbook.Worksheets(SH_DEVICES)
    Set loDevices = wsDevices.ListObjects(TBL_DEVICES)
    On Error GoTo ErrHandler

    If Not loDevices Is Nothing Then
        If Not loDevices.DataBodyRange Is Nothing Then
            ReadHeaders loDevices, devHeaders, numDevCols
            ReadData loDevices, devData, devRows
            devIedNameCol = FindHeader(devHeaders, numDevCols, HDR_IED_NAME)
            devAreaCol    = FindHeader(devHeaders, numDevCols, HDR_IED_AREA)
            If devIedNameCol > 0 And devAreaCol > 0 Then
                For r = 1 To devRows
                    iedName = SafeStr(devData, r, devIedNameCol)
                    areaVal = SafeStr(devData, r, devAreaCol)
                    If Len(iedName) > 0 And Not dictDevArea.Exists(iedName) Then
                        dictDevArea.Add iedName, areaVal
                        dictDevTotal.Add iedName, 0
                        dictDevFilled.Add iedName, 0
                    End If
                Next r
            End If
        End If
    End If

    ' ------------------------------------------------------------------
    ' Read Points List
    ' ------------------------------------------------------------------
    currentStep = "COMPLETION: Reading Points List"
    Set wsPoints = ThisWorkbook.Worksheets(SH_POINTS)
    Set loPoints = wsPoints.ListObjects(TBL_POINTS)

    If loPoints.DataBodyRange Is Nothing Or loPoints.ListRows.Count = 0 Then
        MsgBox "Points List is empty.", vbExclamation, "Generate Completion Report"
        GoTo CleanExit
    End If

    ReadHeaders loPoints, ptsHeaders, numPtsCols
    ReadData loPoints, ptsData, ptsRows

    ' ------------------------------------------------------------------
    ' Locate required Points List columns
    ' ------------------------------------------------------------------
    currentStep = "COMPLETION: Locating columns"
    iedNameCol = FindHeader(ptsHeaders, numPtsCols, HDR_IED_NAME)
    descCol    = FindHeader(ptsHeaders, numPtsCols, HDR_DESCRIPTION)
    ptAreaCol  = FindHeader(ptsHeaders, numPtsCols, HDR_IED_AREA)  ' fallback only

    If iedNameCol = 0 Then
        MsgBox "Points List missing '" & HDR_IED_NAME & "' column.", vbCritical
        GoTo CleanExit
    End If
    If descCol = 0 Then
        MsgBox "Points List missing '" & HDR_DESCRIPTION & "' column.", vbCritical
        GoTo CleanExit
    End If

    ' ------------------------------------------------------------------
    ' Pass 1: accumulate per-device totals
    '         IED AREA: Device List primary, Points List fallback
    ' ------------------------------------------------------------------
    currentStep = "COMPLETION: Building device statistics"
    Application.StatusBar = "Analysing " & ptsRows & " point(s)..."

    For r = 1 To ptsRows
        iedName = SafeStr(ptsData, r, iedNameCol)
        If Len(iedName) = 0 Then GoTo NextDataRow

        descVal = SafeStr(ptsData, r, descCol)

        If Not dictDevTotal.Exists(iedName) Then
            dictDevTotal.Add iedName, 0
            dictDevFilled.Add iedName, 0
            ' If Device List didn't supply an area, fall back to Points List column
            If Not dictDevArea.Exists(iedName) Then
                areaVal = IIf(ptAreaCol > 0, SafeStr(ptsData, r, ptAreaCol), "")
                dictDevArea.Add iedName, areaVal
            End If
        End If

        dictDevTotal(iedName)  = dictDevTotal(iedName) + 1
        If Len(descVal) > 0 Then dictDevFilled(iedName) = dictDevFilled(iedName) + 1

NextDataRow:
    Next r

    ' ------------------------------------------------------------------
    ' Pass 2: classify devices and build area summary
    ' ------------------------------------------------------------------
    currentStep = "COMPLETION: Building area summary"

    Set dictAreaComplete   = CreateObject("Scripting.Dictionary")
    Set dictAreaIncomplete = CreateObject("Scripting.Dictionary")
    Set dictAreaZero       = CreateObject("Scripting.Dictionary")
    dictAreaComplete.CompareMode   = vbTextCompare
    dictAreaIncomplete.CompareMode = vbTextCompare
    dictAreaZero.CompareMode       = vbTextCompare

    devNames = dictDevTotal.Keys

    For Each devName In devNames
        area = CStr(dictDevArea(devName))
        If Not dictAreaComplete.Exists(area) Then
            dictAreaComplete.Add area, 0
            dictAreaIncomplete.Add area, 0
            dictAreaZero.Add area, 0
        End If
        If CLng(dictDevTotal(devName)) > 0 And CLng(dictDevFilled(devName)) = CLng(dictDevTotal(devName)) Then
            dictAreaComplete(area)   = dictAreaComplete(area) + 1
            totalComplete = totalComplete + 1
        ElseIf CLng(dictDevTotal(devName)) = 0 Then
            dictAreaZero(area)       = dictAreaZero(area) + 1
            totalZero = totalZero + 1
        Else
            dictAreaIncomplete(area) = dictAreaIncomplete(area) + 1
            totalIncomplete = totalIncomplete + 1
        End If
    Next devName

    totalDevices = totalComplete + totalIncomplete + totalZero

    ' ------------------------------------------------------------------
    ' Sort area keys alphabetically
    ' ------------------------------------------------------------------
    areaKeys = dictAreaComplete.Keys
    For i = 0 To UBound(areaKeys) - 1
        For j = i + 1 To UBound(areaKeys)
            If CStr(areaKeys(i)) > CStr(areaKeys(j)) Then
                tmpS = areaKeys(i): areaKeys(i) = areaKeys(j): areaKeys(j) = tmpS
            End If
        Next j
    Next i

    ' ------------------------------------------------------------------
    ' Collect completed and incomplete device arrays for sorting
    ' ------------------------------------------------------------------
    currentStep = "COMPLETION: Collecting device lists"

    If totalComplete > 0 Then
        ReDim completedNames(1 To totalComplete)
        ReDim completedAreas(1 To totalComplete)
        ReDim completedPts(1 To totalComplete)
    End If
    If totalIncomplete > 0 Then
        ReDim incompleteNames(1 To totalIncomplete)
        ReDim incompleteAreas(1 To totalIncomplete)
        ReDim incompleteFilled(1 To totalIncomplete)
        ReDim incompleteTotals(1 To totalIncomplete)
    End If
    If totalZero > 0 Then
        ReDim zeroNames(1 To totalZero)
        ReDim zeroAreas(1 To totalZero)
    End If

    completedCount   = 0
    incompletedCount = 0
    zeroCount        = 0

    For Each devName In devNames
        If CLng(dictDevTotal(devName)) > 0 And CLng(dictDevFilled(devName)) = CLng(dictDevTotal(devName)) Then
            completedCount = completedCount + 1
            completedNames(completedCount) = CStr(devName)
            completedAreas(completedCount) = CStr(dictDevArea(devName))
            completedPts(completedCount)   = CLng(dictDevTotal(devName))
        ElseIf CLng(dictDevTotal(devName)) = 0 Then
            zeroCount = zeroCount + 1
            zeroNames(zeroCount) = CStr(devName)
            zeroAreas(zeroCount) = CStr(dictDevArea(devName))
        Else
            incompletedCount = incompletedCount + 1
            incompleteNames(incompletedCount)  = CStr(devName)
            incompleteAreas(incompletedCount)  = CStr(dictDevArea(devName))
            incompleteFilled(incompletedCount) = CLng(dictDevFilled(devName))
            incompleteTotals(incompletedCount) = CLng(dictDevTotal(devName))
        End If
    Next devName

    ' Sort completed: IED AREA then IED NAME
    For i = 1 To completedCount - 1
        For j = i + 1 To completedCount
            doSwap = False
            If completedAreas(i) > completedAreas(j) Then
                doSwap = True
            ElseIf completedAreas(i) = completedAreas(j) And completedNames(i) > completedNames(j) Then
                doSwap = True
            End If
            If doSwap Then
                tmpS = completedNames(i): completedNames(i) = completedNames(j): completedNames(j) = tmpS
                tmpS = completedAreas(i): completedAreas(i) = completedAreas(j): completedAreas(j) = tmpS
                tmpL = completedPts(i):   completedPts(i)   = completedPts(j):   completedPts(j)   = tmpL
            End If
        Next j
    Next i

    ' Sort incomplete: IED AREA then IED NAME
    For i = 1 To incompletedCount - 1
        For j = i + 1 To incompletedCount
            doSwap = False
            If incompleteAreas(i) > incompleteAreas(j) Then
                doSwap = True
            ElseIf incompleteAreas(i) = incompleteAreas(j) And incompleteNames(i) > incompleteNames(j) Then
                doSwap = True
            End If
            If doSwap Then
                tmpS = incompleteNames(i):  incompleteNames(i)  = incompleteNames(j):  incompleteNames(j)  = tmpS
                tmpS = incompleteAreas(i):  incompleteAreas(i)  = incompleteAreas(j):  incompleteAreas(j)  = tmpS
                tmpL = incompleteFilled(i): incompleteFilled(i) = incompleteFilled(j): incompleteFilled(j) = tmpL
                tmpL = incompleteTotals(i): incompleteTotals(i) = incompleteTotals(j): incompleteTotals(j) = tmpL
            End If
        Next j
    Next i

    ' Sort zero-point: IED AREA then IED NAME
    For i = 1 To zeroCount - 1
        For j = i + 1 To zeroCount
            doSwap = False
            If zeroAreas(i) > zeroAreas(j) Then
                doSwap = True
            ElseIf zeroAreas(i) = zeroAreas(j) And zeroNames(i) > zeroNames(j) Then
                doSwap = True
            End If
            If doSwap Then
                tmpS = zeroNames(i): zeroNames(i) = zeroNames(j): zeroNames(j) = tmpS
                tmpS = zeroAreas(i): zeroAreas(i) = zeroAreas(j): zeroAreas(j) = tmpS
            End If
        Next j
    Next i

    ' ------------------------------------------------------------------
    ' Write report sheet
    ' ------------------------------------------------------------------
    currentStep = "COMPLETION: Writing report"
    Application.StatusBar = "Writing report..."

    Set wsReport = GetOrCreateSheet(SH_COMPLETION)
    wsReport.Cells.Clear
    writeRow = 2  ' Row 1 reserved for buttons

    ' --- Section 1: Summary by IED AREA ---
    With wsReport.Cells(writeRow, 1)
        .Value = "SUMMARY BY IED AREA"
        .Font.Bold = True
        .Font.Size = 12
    End With
    writeRow = writeRow + 1

    wsReport.Cells(writeRow, 1).Value = "IED AREA"
    wsReport.Cells(writeRow, 2).Value = "Complete"
    wsReport.Cells(writeRow, 3).Value = "Incomplete"
    wsReport.Cells(writeRow, 4).Value = "Zero Pts"
    wsReport.Cells(writeRow, 5).Value = "Total"
    wsReport.Cells(writeRow, 6).Value = "% Complete"
    With wsReport.Range(wsReport.Cells(writeRow, 1), wsReport.Cells(writeRow, 6))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    writeRow = writeRow + 1

    For Each areaKey In areaKeys
        area            = CStr(areaKey)
        completeCount   = CLng(dictAreaComplete(area))
        incompleteCount = CLng(dictAreaIncomplete(area))
        zeroAreaCount   = CLng(dictAreaZero(area))
        areaTotal       = completeCount + incompleteCount + zeroAreaCount
        pctComplete     = IIf(areaTotal > 0, completeCount / areaTotal, 0)

        wsReport.Cells(writeRow, 1).Value = area
        wsReport.Cells(writeRow, 2).Value = completeCount
        wsReport.Cells(writeRow, 3).Value = incompleteCount
        wsReport.Cells(writeRow, 4).Value = zeroAreaCount
        wsReport.Cells(writeRow, 5).Value = areaTotal
        wsReport.Cells(writeRow, 6).Value = pctComplete
        wsReport.Cells(writeRow, 6).NumberFormat = "0%"

        With wsReport.Range(wsReport.Cells(writeRow, 1), wsReport.Cells(writeRow, 6))
            If completeCount = areaTotal Then
                .Interior.Color = RGB(198, 239, 206)   ' green
            ElseIf completeCount = 0 Then
                .Interior.Color = RGB(255, 199, 206)   ' red
            Else
                .Interior.Color = RGB(255, 235, 156)   ' amber
            End If
        End With

        writeRow = writeRow + 1
    Next areaKey

    ' Total row
    pctComplete = IIf(totalDevices > 0, totalComplete / totalDevices, 0)
    wsReport.Cells(writeRow, 1).Value = "TOTAL"
    wsReport.Cells(writeRow, 2).Value = totalComplete
    wsReport.Cells(writeRow, 3).Value = totalIncomplete
    wsReport.Cells(writeRow, 4).Value = totalZero
    wsReport.Cells(writeRow, 5).Value = totalDevices
    wsReport.Cells(writeRow, 6).Value = pctComplete
    wsReport.Cells(writeRow, 6).NumberFormat = "0%"
    With wsReport.Range(wsReport.Cells(writeRow, 1), wsReport.Cells(writeRow, 6))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    writeRow = writeRow + 2  ' blank gap row

    ' --- Section 2: Side-by-side detail ---
    ' Title row — left and right titles on same row
    With wsReport.Cells(writeRow, LEFT_COL)
        .Value = "COMPLETED DEVICES (" & completedCount & ")"
        .Font.Bold = True
        .Font.Size = 12
    End With
    With wsReport.Cells(writeRow, RIGHT_COL)
        .Value = "INCOMPLETE DEVICES (" & incompletedCount & ")"
        .Font.Bold = True
        .Font.Size = 12
    End With
    With wsReport.Cells(writeRow, FAR_COL)
        .Value = "ZERO-POINT DEVICES (" & zeroCount & ")"
        .Font.Bold = True
        .Font.Size = 12
    End With
    writeRow = writeRow + 1

    ' Column header row
    wsReport.Cells(writeRow, LEFT_COL).Value     = "IED NAME"
    wsReport.Cells(writeRow, LEFT_COL + 1).Value = "IED AREA"
    wsReport.Cells(writeRow, LEFT_COL + 2).Value = "Points"
    With wsReport.Range(wsReport.Cells(writeRow, LEFT_COL), wsReport.Cells(writeRow, LEFT_COL + 2))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    wsReport.Cells(writeRow, RIGHT_COL).Value     = "IED NAME"
    wsReport.Cells(writeRow, RIGHT_COL + 1).Value = "IED AREA"
    wsReport.Cells(writeRow, RIGHT_COL + 2).Value = "Described"
    wsReport.Cells(writeRow, RIGHT_COL + 3).Value = "Total"
    With wsReport.Range(wsReport.Cells(writeRow, RIGHT_COL), wsReport.Cells(writeRow, RIGHT_COL + 3))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    wsReport.Cells(writeRow, FAR_COL).Value     = "IED NAME"
    wsReport.Cells(writeRow, FAR_COL + 1).Value = "IED AREA"
    With wsReport.Range(wsReport.Cells(writeRow, FAR_COL), wsReport.Cells(writeRow, FAR_COL + 1))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    writeRow = writeRow + 1

    detailStartRow = writeRow

    ' Bulk write left (completed) column
    If completedCount > 0 Then
        ReDim leftArr(1 To completedCount, 1 To 3)
        For i = 1 To completedCount
            leftArr(i, 1) = completedNames(i)
            leftArr(i, 2) = completedAreas(i)
            leftArr(i, 3) = completedPts(i)
        Next i
        With wsReport.Range(wsReport.Cells(detailStartRow, LEFT_COL), _
                            wsReport.Cells(detailStartRow + completedCount - 1, LEFT_COL + 2))
            .Value = leftArr
            .Interior.Color = RGB(198, 239, 206)  ' green
        End With
    Else
        wsReport.Cells(detailStartRow, LEFT_COL).Value = "No completed devices."
        wsReport.Cells(detailStartRow, LEFT_COL).Font.Italic = True
    End If

    ' Bulk write right (incomplete) column
    If incompletedCount > 0 Then
        ReDim rightArr(1 To incompletedCount, 1 To 4)
        For i = 1 To incompletedCount
            rightArr(i, 1) = incompleteNames(i)
            rightArr(i, 2) = incompleteAreas(i)
            rightArr(i, 3) = incompleteFilled(i)
            rightArr(i, 4) = incompleteTotals(i)
        Next i
        With wsReport.Range(wsReport.Cells(detailStartRow, RIGHT_COL), _
                            wsReport.Cells(detailStartRow + incompletedCount - 1, RIGHT_COL + 3))
            .Value = rightArr
            .Interior.Color = RGB(255, 235, 156)  ' amber
        End With
    Else
        wsReport.Cells(detailStartRow, RIGHT_COL).Value = "No incomplete devices."
        wsReport.Cells(detailStartRow, RIGHT_COL).Font.Italic = True
    End If

    ' Bulk write far-right (zero-point) column
    If zeroCount > 0 Then
        ReDim zeroArr(1 To zeroCount, 1 To 2)
        For i = 1 To zeroCount
            zeroArr(i, 1) = zeroNames(i)
            zeroArr(i, 2) = zeroAreas(i)
        Next i
        With wsReport.Range(wsReport.Cells(detailStartRow, FAR_COL), _
                            wsReport.Cells(detailStartRow + zeroCount - 1, FAR_COL + 1))
            .Value = zeroArr
            .Interior.Color = RGB(255, 199, 206)  ' red/pink - no points configured
        End With
    Else
        wsReport.Cells(detailStartRow, FAR_COL).Value = "No zero-point devices."
        wsReport.Cells(detailStartRow, FAR_COL).Font.Italic = True
    End If

    wsReport.Columns("A:K").AutoFit
    wsReport.Activate

    Application.StatusBar = False
    MsgBox "Completion report generated." & vbNewLine & _
           "  Total devices:      " & totalDevices & vbNewLine & _
           "  Complete:           " & totalComplete & vbNewLine & _
           "  Incomplete:         " & totalIncomplete & vbNewLine & _
           "  Zero-point:         " & totalZero, _
           vbInformation, "Generate Completion Report"

CleanExit:
    Set dictDevTotal = Nothing
    Set dictDevFilled = Nothing
    Set dictDevArea = Nothing
    Set dictAreaComplete = Nothing
    Set dictAreaIncomplete = Nothing
    Set dictAreaZero = Nothing
    Set loPoints = Nothing
    Set loDevices = Nothing
    Set wsPoints = Nothing
    Set wsDevices = Nothing
    Set wsReport = Nothing
    PerformanceOff
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Debug.Print "*** ERROR at: " & currentStep & " | " & Err.Number & ": " & Err.Description
    MsgBox "Error at: " & currentStep & vbNewLine & vbNewLine & _
           "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit

End Sub


' ============================================================================
'  PRIVATE: RunCompareEngine
' ============================================================================
' Shared engine for ComparePointsList and UpdatePointsList.
'
' runMode = MODE_COMPARE : highlight differences, don't change data
' runMode = MODE_UPDATE  : overwrite differing cells with expected values
'
' ALGORITHM:
'   1. Load context (all table data and lookup structures)
'   2. Build the full expected output array (no devices skipped)
'   3. Read the existing Points List into an array
'   4. Index expected rows by TAG NAME for O(1) lookup
'   5. For each existing row, find its expected counterpart by TAG NAME
'   6. Compare every cell - if different, highlight/update and log
'   7. Track expected rows not found in existing (missing devices)
'   8. Write the change log to the Compare sheet
'
' WHY SKIP BLANK DESCRIPTIONS:
'   When a template row has no TEMP TAG DESCRIPTION (intentionally blank),
'   the description and all columns to its right are considered user-owned
'   and are excluded from comparison.
' ============================================================================
Private Sub RunCompareEngine(ByVal runMode As Long)

    Dim currentStep  As String
    Dim msg          As String
    Dim modeLabel    As String
    Dim ctx          As Object
    Dim dictNone     As Object
    Dim dictExpected As Object
    Dim loPoints     As ListObject
    Dim wsPoints     As Worksheet
    Dim wsCompare    As Worksheet
    Dim ptsHeaders() As String
    Dim expArr()     As Variant
    Dim existArr     As Variant
    Dim logArr()     As Variant
    Dim writeLog()   As Variant
    Dim missingKeys  As Variant
    Dim numPtsCols   As Long
    Dim expRows      As Long
    Dim existRows    As Long
    Dim dummySE      As Long
    Dim dummySN      As Long
    Dim tagNameCol   As Long
    Dim ptsDescCol   As Long
    Dim ptsCatCol    As Long
    Dim ptsAesoReqCol  As Long
    Dim ptsAesoDescCol As Long
    Dim blankDesc    As Boolean
    Dim dataStartRow As Long
    Dim dataStartCol As Long
    Dim logRows      As Long
    Dim logMax       As Long
    Dim diffRowCount As Long
    Dim diffCellCount As Long
    Dim missingCount As Long
    Dim updatedCount As Long
    Dim tagKey       As String
    Dim existVal     As String
    Dim expVal       As String
    Dim expRow       As Long
    Dim r            As Long
    Dim c            As Long
    Dim k            As Long
    Dim lr           As Long
    Dim rowHasDiff() As Boolean   ' per-row: True if any cell in this row differs
    Dim cellIsDiff() As Boolean   ' per-cell: True if this specific cell differs
    Dim ptsIedNameCol    As Long
    Dim dictExistingDevs As Object
    Dim dictExpTagToDev  As Object
    Dim expIedName       As String
    Dim missingTagName   As String

    If runMode = MODE_UPDATE Then modeLabel = "UPD" Else modeLabel = "CMP"

    On Error GoTo ErrHandler
    PerformanceOn

    ' ------------------------------------------------------------------
    ' Load all table data and build lookup structures
    ' ------------------------------------------------------------------
    currentStep = modeLabel & ": Loading context"
    Set ctx = LoadContext()
    If ctx Is Nothing Then GoTo CleanExit
    If CLng(ctx("devRows")) = 0 Then MsgBox "DEVICE_LIST is empty.", vbExclamation: GoTo CleanExit
    If CLng(ctx("tmpRows")) = 0 Then MsgBox "TEMPLATE_TABLE is empty.", vbExclamation: GoTo CleanExit

    numPtsCols = CLng(ctx("numPtsCols"))
    ptsHeaders = ctx("ptsHeaders")

    ' ------------------------------------------------------------------
    ' Build the FULL expected output - pass an empty Dictionary so no
    ' devices are skipped.
    ' ------------------------------------------------------------------
    currentStep = modeLabel & ": Building expected output"
    Set dictNone = CreateObject("Scripting.Dictionary")
    BuildExpectedArray ctx, dictNone, expArr, expRows, dummySE, dummySN

    If expRows = 0 Then
        MsgBox "No expected points could be generated from current data.", vbExclamation
        GoTo CleanExit
    End If

    ' ------------------------------------------------------------------
    ' Read existing Points List into an array for comparison.
    ' ------------------------------------------------------------------
    currentStep = modeLabel & ": Reading existing Points List"
    Set loPoints = ThisWorkbook.Worksheets(SH_POINTS).ListObjects(TBL_POINTS)
    ReadData loPoints, existArr, existRows
    If existRows = 0 Then
        MsgBox "Points List is empty. Run GeneratePointsList first.", vbExclamation
        GoTo CleanExit
    End If

    ' ------------------------------------------------------------------
    ' Locate the TAG NAME column (comparison key) and DESCRIPTION column.
    ' ------------------------------------------------------------------
    currentStep = modeLabel & ": Locating columns"
    tagNameCol = FindHeader(ptsHeaders, numPtsCols, HDR_TAG_NAME)
    If tagNameCol = 0 Then
        MsgBox "Points List missing '" & HDR_TAG_NAME & "' column.", vbCritical
        GoTo CleanExit
    End If
    ptsDescCol = CLng(ctx("ptsDescCol"))
    ptsCatCol  = CLng(ctx("ptsCatCol"))
    ptsAesoReqCol  = FindHeader(ptsHeaders, numPtsCols, HDR_AESO_REQ)
    ptsAesoDescCol = FindHeader(ptsHeaders, numPtsCols, HDR_AESO_DESC)
    ptsIedNameCol = FindHeader(ptsHeaders, numPtsCols, HDR_IED_NAME)

    ' ------------------------------------------------------------------
    ' Index expected rows by TAG NAME.  First occurrence wins for duplicates.
    ' ------------------------------------------------------------------
    currentStep = modeLabel & ": Indexing expected rows"
    Set dictExpected = CreateObject("Scripting.Dictionary")
    dictExpected.CompareMode = vbTextCompare
    For r = 1 To expRows
        tagKey = SafeStrVal(expArr(r, tagNameCol))
        If Len(tagKey) > 0 Then
            If Not dictExpected.Exists(tagKey) Then dictExpected.Add tagKey, r
        End If
    Next r

    ' ------------------------------------------------------------------
    ' COMPARISON LOOP
    ' ------------------------------------------------------------------
    currentStep = modeLabel & ": Comparing"
    Set wsPoints = ThisWorkbook.Worksheets(SH_POINTS)
    dataStartRow = loPoints.HeaderRowRange.Row + 1
    dataStartCol = loPoints.HeaderRowRange.Column

    ' In Compare mode, clear previous highlights so we start fresh.
    If runMode = MODE_COMPARE Then
        If Not loPoints.DataBodyRange Is Nothing Then
            loPoints.DataBodyRange.Interior.ColorIndex = xlNone
        End If
    End If

    ' Pre-allocate log and diff-tracking arrays.
    ' rowHasDiff / cellIsDiff are filled during the pure in-memory comparison loop.
    ' All sheet writes (highlights / value updates) happen AFTER the loop in a
    ' separate pass, which keeps the loop tight and Excel responsive.
    logMax = existRows * numPtsCols + expRows
    ReDim logArr(1 To logMax, 1 To 5)  ' Sheet Row | TAG NAME | Column | Existing | Expected
    ReDim rowHasDiff(1 To existRows)
    ReDim cellIsDiff(1 To existRows, 1 To numPtsCols)
    logRows = 0
    diffRowCount = 0
    diffCellCount = 0
    missingCount = 0
    updatedCount = 0

    ' ------------------------------------------------------------------
    ' PURE COMPARISON LOOP - no sheet interaction whatsoever.
    '
    ' STATUS BAR THROTTLE: updating the status bar forces a Windows UI
    ' event-loop cycle on every call.  At one call per row this is the
    ' primary cause of Excel appearing to stop responding on large lists.
    ' Updating every STATUS_INTERVAL rows gives smooth progress with
    ' negligible overhead.
    ' ------------------------------------------------------------------
    Const STATUS_INTERVAL As Long = 25

    For r = 1 To existRows
        If r Mod STATUS_INTERVAL = 1 Then
            Application.StatusBar = modeLabel & ": " & ProgressBar(r, existRows) & _
                                     "  Row " & r & " of " & existRows & _
                                     "  (" & diffCellCount & " difference(s) found)"
        End If

        tagKey = SafeStr(existArr, r, tagNameCol)

        ' Skip rows with no TAG NAME, or rows with no expected match
        ' (these might be manually added rows not template-driven).
        If Len(tagKey) = 0 Then GoTo NextCmpRow
        If Not dictExpected.Exists(tagKey) Then GoTo NextCmpRow

        expRow = CLng(dictExpected(tagKey))

        ' Pre-check: is this row's template description blank?
        blankDesc = (ptsDescCol > 0 And Len(SafeStrVal(expArr(expRow, ptsDescCol))) = 0)

        For c = 1 To numPtsCols
            existVal = SafeStr(existArr, r, c)
            expVal = SafeStrVal(expArr(expRow, c))

            ' When the template has no description, also skip POINT CATEGORY.
            If blankDesc And ptsCatCol > 0 And c = ptsCatCol Then GoTo NextCmpCol

            ' Always ignore AESO columns - these are manually maintained.
            If ptsAesoReqCol > 0 And c = ptsAesoReqCol Then GoTo NextCmpCol
            If ptsAesoDescCol > 0 And c = ptsAesoDescCol Then GoTo NextCmpCol

            ' When the template has no description for this point, skip the
            ' description column and all columns to the right.
            If c = ptsDescCol And Len(expVal) = 0 Then Exit For

            If StrComp(existVal, expVal, vbTextCompare) <> 0 Then

                ' Log the difference
                logRows = logRows + 1
                logArr(logRows, 1) = dataStartRow + r - 1
                logArr(logRows, 2) = tagKey
                logArr(logRows, 3) = ptsHeaders(c)
                logArr(logRows, 4) = existVal
                logArr(logRows, 5) = expVal
                diffCellCount = diffCellCount + 1

                ' Record diff position for the post-loop sheet pass.
                rowHasDiff(r) = True
                cellIsDiff(r, c) = True

                ' MODE_UPDATE: modify existArr in place.
                ' The whole array is written back to the sheet in one bulk
                ' operation after the loop - no per-cell COM calls here.
                If runMode = MODE_UPDATE Then
                    If Len(expVal) > 0 Then
                        existArr(r, c) = expArr(expRow, c)
                    Else
                        existArr(r, c) = Empty
                    End If
                    updatedCount = updatedCount + 1
                End If

            End If

NextCmpCol:
        Next c

        If rowHasDiff(r) Then diffRowCount = diffRowCount + 1

        ' Remove this key; remaining keys after the loop are "missing" rows.
        dictExpected.Remove tagKey

NextCmpRow:
    Next r

    ' ------------------------------------------------------------------
    ' POST-LOOP SHEET OPERATIONS
    ' All sheet interaction is deferred to here.  The comparison loop above
    ' is pure in-memory work, which is why it stays responsive.
    ' ------------------------------------------------------------------

    ' MODE_UPDATE: write the modified array back in a single bulk operation.
    ' existArr now contains the corrected values; all other values are unchanged.
    If runMode = MODE_UPDATE And updatedCount > 0 Then
        currentStep = modeLabel & ": Writing updates"
        Application.StatusBar = "Writing " & updatedCount & " updated cell(s)..."
        loPoints.DataBodyRange.Value = existArr
    End If

    ' MODE_COMPARE: apply highlighting in two passes to minimise COM calls.
    '   Pass 1 - yellow: one range write per differing row (covers the whole row).
    '   Pass 2 - red:    one cell write per differing cell (overwrites yellow).
    ' This replaces the original approach which read Interior.Color for every
    ' cell in every differing row just to avoid overwriting red with yellow.
    If runMode = MODE_COMPARE And diffRowCount > 0 Then
        currentStep = modeLabel & ": Applying highlights"
        Application.StatusBar = "Highlighting " & diffRowCount & " row(s)..."

        ' Pass 1: yellow rows
        For r = 1 To existRows
            If rowHasDiff(r) Then
                wsPoints.Range( _
                    wsPoints.Cells(dataStartRow + r - 1, dataStartCol), _
                    wsPoints.Cells(dataStartRow + r - 1, dataStartCol + numPtsCols - 1) _
                ).Interior.Color = CLR_ROW_YELLOW
            End If
        Next r

        ' Pass 2: red cells
        For r = 1 To existRows
            If rowHasDiff(r) Then
                For c = 1 To numPtsCols
                    If cellIsDiff(r, c) Then
                        wsPoints.Cells(dataStartRow + r - 1, dataStartCol + c - 1) _
                            .Interior.Color = CLR_CELL_RED
                    End If
                Next c
            End If
        Next r
    End If

    ' ------------------------------------------------------------------
    ' Log expected rows not found in the Points List (missing devices)
    ' ------------------------------------------------------------------
    currentStep = modeLabel & ": Logging missing rows"
    If dictExpected.Count > 0 Then

        ' Build a set of device names present in the existing Points List.
        ' Devices with no points yet are excluded from "missing" reporting.
        Set dictExistingDevs = CreateObject("Scripting.Dictionary")
        dictExistingDevs.CompareMode = vbTextCompare
        If ptsIedNameCol > 0 Then
            For r = 1 To existRows
                expIedName = SafeStr(existArr, r, ptsIedNameCol)
                If Len(expIedName) > 0 Then
                    If Not dictExistingDevs.Exists(expIedName) Then
                        dictExistingDevs.Add expIedName, True
                    End If
                End If
            Next r
        End If

        ' Build a map: expected TAG NAME -> IED NAME, so we can look up the
        ' device for each remaining (unmatched) expected key.
        Set dictExpTagToDev = CreateObject("Scripting.Dictionary")
        dictExpTagToDev.CompareMode = vbTextCompare
        If ptsIedNameCol > 0 Then
            For r = 1 To expRows
                tagKey = SafeStrVal(expArr(r, tagNameCol))
                expIedName = SafeStrVal(expArr(r, ptsIedNameCol))
                If Len(tagKey) > 0 And Not dictExpTagToDev.Exists(tagKey) Then
                    dictExpTagToDev.Add tagKey, expIedName
                End If
            Next r
        End If

        missingKeys = dictExpected.Keys
        For k = 0 To UBound(missingKeys)
            missingTagName = CStr(missingKeys(k))

            ' Skip if the device has no points in the existing list yet.
            If ptsIedNameCol > 0 And dictExpTagToDev.Exists(missingTagName) Then
                expIedName = CStr(dictExpTagToDev(missingTagName))
                If Len(expIedName) > 0 And Not dictExistingDevs.Exists(expIedName) Then
                    GoTo NextMissingKey
                End If
            End If

            logRows = logRows + 1
            logArr(logRows, 1) = "N/A"
            logArr(logRows, 2) = missingTagName
            logArr(logRows, 3) = "(entire row)"
            logArr(logRows, 4) = "(missing)"
            logArr(logRows, 5) = "(expected but not in Points List)"
            missingCount = missingCount + 1
NextMissingKey:
        Next k
    End If

    ' ------------------------------------------------------------------
    ' Write Compare sheet
    ' ------------------------------------------------------------------
    currentStep = modeLabel & ": Writing Compare sheet"
    Set wsCompare = GetOrCreateSheet(SH_COMPARE)
    wsCompare.Cells.Clear

    wsCompare.Cells(2, 1).Value = "Sheet Row"
    wsCompare.Cells(2, 2).Value = "TAG NAME"
    wsCompare.Cells(2, 3).Value = "Column"
    wsCompare.Cells(2, 4).Value = "Existing Value"
    wsCompare.Cells(2, 5).Value = "Expected Value"
    With wsCompare.Range("A2:E2")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    If logRows > 0 Then
        ReDim writeLog(1 To logRows, 1 To 5)
        For lr = 1 To logRows
            writeLog(lr, 1) = logArr(lr, 1)
            writeLog(lr, 2) = logArr(lr, 2)
            writeLog(lr, 3) = logArr(lr, 3)
            writeLog(lr, 4) = logArr(lr, 4)
            writeLog(lr, 5) = logArr(lr, 5)
        Next lr
        wsCompare.Range("A3").Resize(logRows, 5).Value = writeLog
    End If

    wsCompare.Columns("A:E").AutoFit
    Application.StatusBar = False

    ' ------------------------------------------------------------------
    ' Summary message
    ' ------------------------------------------------------------------
    If runMode = MODE_COMPARE Then
        If logRows = 0 Then
            MsgBox "No differences found. Points List matches expected output.", vbInformation
        Else
            msg = "Comparison complete." & vbNewLine & _
                  "  Rows with differences:  " & diffRowCount & vbNewLine & _
                  "  Cell differences:       " & diffCellCount & vbNewLine & _
                  "  Missing rows:           " & missingCount & vbNewLine & vbNewLine & _
                  "Details written to '" & SH_COMPARE & "' sheet." & vbNewLine & _
                  "Differences highlighted in Points List (yellow row / red cell)."
            MsgBox msg, vbInformation
        End If

    ElseIf runMode = MODE_UPDATE Then
        If updatedCount = 0 And missingCount = 0 Then
            MsgBox "No differences found. Points List is already up to date.", vbInformation
        Else
            msg = "Update complete." & vbNewLine & _
                  "  Cells updated:              " & updatedCount & vbNewLine & _
                  "  Rows affected:              " & diffRowCount & vbNewLine & _
                  "  Missing rows (not updated): " & missingCount & vbNewLine & vbNewLine & _
                  "Change log written to '" & SH_COMPARE & "' sheet."
            MsgBox msg, vbInformation
        End If
    End If

CleanExit:
    Erase expArr
    Erase logArr
    Erase rowHasDiff
    Erase cellIsDiff
    Set ctx = Nothing
    Set dictNone = Nothing
    Set dictExpected = Nothing
    Set loPoints = Nothing
    Set wsPoints = Nothing
    PerformanceOff
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Debug.Print "*** ERROR at: " & currentStep & " | " & Err.Number & ": " & Err.Description
    MsgBox "Error at: " & currentStep & vbNewLine & vbNewLine & _
           "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit

End Sub


' ============================================================================
'  PRIVATE: LoadContext
' ============================================================================
' Reads all four tables into arrays and builds every lookup structure needed
' by the generation and comparison engines.  Returns a Dictionary containing
' all context, or Nothing on failure.
'
' KEY CONTENTS:
'   devData / tmpData / catData   - 2D Variant arrays of table data
'   devRows / tmpRows / catRows   - row counts
'   ptsHeaders / numPtsCols       - Points List header info
'   devKeyCol / devNameCol / etc. - column indices for key fields
'   colSource / colSourceIdx      - column map arrays (see below)
'   dictTemplate                  - I/O TEMPLATE -> Collection of row indices
'   dictCategory                  - POINT CATEGORY -> row index
'   tokenText / tokenDevCol       - replacement token mappings
'
' COLUMN MAP (colSource / colSourceIdx):
'   For each Points List column c, colSource(c) indicates the data source:
'     "D" = Device List,    column index in colSourceIdx(c)
'     "T" = Template Table, column index in colSourceIdx(c)
'     "C" = Category Table, column index in colSourceIdx(c)
'     "X" = Derived field (TAG NAME = IED NAME + separator + TAG SUFFIX)
'     ""  = Unmapped - no matching header found, column left blank
'
'   Resolution order: Device List > Template Table > Category Table > Aliases.
' ============================================================================
Private Function LoadContext() As Object

    Dim ctx          As Object
    Dim loDevices    As ListObject
    Dim loTemplates  As ListObject
    Dim loCategories As ListObject
    Dim loPoints     As ListObject
    Dim devHeaders() As String
    Dim tmpHeaders() As String
    Dim catHeaders() As String
    Dim ptsHeaders() As String
    Dim devData      As Variant
    Dim tmpData      As Variant
    Dim catData      As Variant
    Dim devCols      As Long
    Dim tmpCols      As Long
    Dim catCols      As Long
    Dim numPtsCols   As Long
    Dim devRows      As Long
    Dim tmpRows      As Long
    Dim catRows      As Long
    Dim dictDevHdr   As Object
    Dim dictTmpHdr   As Object
    Dim dictCatHdr   As Object
    Dim aliasDict    As Object
    Dim devKeyCol    As Long
    Dim devNameCol   As Long
    Dim tmpKeyCol    As Long
    Dim tmpCatCol    As Long
    Dim catKeyCol    As Long
    Dim ptsDescCol   As Long
    Dim ptsSuffixCol As Long
    Dim ptsCatCol    As Long
    Dim colSource()  As String
    Dim colSourceIdx() As Long
    Dim tokens()     As String
    Dim tokenText()  As String
    Dim tokenDevCol() As Long
    Dim numTokens    As Long
    Dim dictTemplate As Object
    Dim dictCategory As Object
    Dim hdr          As String
    Dim aliasHdr     As String
    Dim tKey         As String
    Dim catKey       As String
    Dim tok          As String
    Dim c            As Long
    Dim r            As Long
    Dim t            As Long

    On Error GoTo Fail

    Set ctx = CreateObject("Scripting.Dictionary")

    ' --- Read table references ---
    Set loDevices = ThisWorkbook.Worksheets(SH_DEVICES).ListObjects(TBL_DEVICES)
    Set loTemplates = ThisWorkbook.Worksheets(SH_TEMPLATES).ListObjects(TBL_TEMPLATES)
    Set loCategories = ThisWorkbook.Worksheets(SH_CATEGORIES).ListObjects(TBL_CATEGORIES)
    Set loPoints = ThisWorkbook.Worksheets(SH_POINTS).ListObjects(TBL_POINTS)

    ' --- Read headers into typed String arrays ---
    ' Using String arrays avoids type-mismatch issues when using UCase/Trim later.
    ReadHeaders loDevices, devHeaders, devCols
    ReadHeaders loTemplates, tmpHeaders, tmpCols
    ReadHeaders loCategories, catHeaders, catCols
    ReadHeaders loPoints, ptsHeaders, numPtsCols

    ' --- Read data bodies into 2D Variant arrays ---
    ReadData loDevices, devData, devRows
    ReadData loTemplates, tmpData, tmpRows
    ReadData loCategories, catData, catRows

    ' --- Build header dictionaries: UCASE(headerName) -> column index ---
    Set dictDevHdr = BuildHeaderDict(devHeaders, devCols)
    Set dictTmpHdr = BuildHeaderDict(tmpHeaders, tmpCols)
    Set dictCatHdr = BuildHeaderDict(catHeaders, catCols)

    ' --- Validate required join-key columns ---
    If Not dictDevHdr.Exists(UCase(HDR_IO_TEMPLATE)) Then _
        MsgBox "DEVICE_LIST missing '" & HDR_IO_TEMPLATE & "'.", vbCritical: GoTo Fail
    If Not dictDevHdr.Exists(UCase(HDR_IED_NAME)) Then _
        MsgBox "DEVICE_LIST missing '" & HDR_IED_NAME & "'.", vbCritical: GoTo Fail
    If Not dictTmpHdr.Exists(UCase(HDR_IO_TEMPLATE)) Then _
        MsgBox "TEMPLATE_TABLE missing '" & HDR_IO_TEMPLATE & "'.", vbCritical: GoTo Fail

    devKeyCol = CLng(dictDevHdr(UCase(HDR_IO_TEMPLATE)))
    devNameCol = CLng(dictDevHdr(UCase(HDR_IED_NAME)))
    tmpKeyCol = CLng(dictTmpHdr(UCase(HDR_IO_TEMPLATE)))

    tmpCatCol = 0
    If dictTmpHdr.Exists(UCase(HDR_POINT_CAT)) Then
        tmpCatCol = CLng(dictTmpHdr(UCase(HDR_POINT_CAT)))
    End If

    ' --- Build the column map ---
    ReDim colSource(1 To numPtsCols)
    ReDim colSourceIdx(1 To numPtsCols)
    Set aliasDict = BuildAliasDict()

    For c = 1 To numPtsCols
        hdr = UCase(Trim(ptsHeaders(c)))

        If hdr = UCase(HDR_TAG_NAME) Then
            colSource(c) = "X"                          ' Derived field

        ElseIf dictDevHdr.Exists(hdr) Then
            colSource(c) = "D"
            colSourceIdx(c) = CLng(dictDevHdr(hdr))

        ElseIf dictTmpHdr.Exists(hdr) Then
            colSource(c) = "T"
            colSourceIdx(c) = CLng(dictTmpHdr(hdr))

        ElseIf dictCatHdr.Exists(hdr) Then
            colSource(c) = "C"
            colSourceIdx(c) = CLng(dictCatHdr(hdr))

        ElseIf aliasDict.Exists(hdr) Then
            aliasHdr = CStr(aliasDict(hdr))
            If dictDevHdr.Exists(aliasHdr) Then
                colSource(c) = "D"
                colSourceIdx(c) = CLng(dictDevHdr(aliasHdr))
            ElseIf dictTmpHdr.Exists(aliasHdr) Then
                colSource(c) = "T"
                colSourceIdx(c) = CLng(dictTmpHdr(aliasHdr))
            ElseIf dictCatHdr.Exists(aliasHdr) Then
                colSource(c) = "C"
                colSourceIdx(c) = CLng(dictCatHdr(aliasHdr))
            End If
        End If
        ' If none match, colSource(c) stays "" - column left blank
    Next c

    ' Locate special output columns by header name
    ptsDescCol = FindHeader(ptsHeaders, numPtsCols, HDR_DESCRIPTION)
    ptsSuffixCol = FindHeader(ptsHeaders, numPtsCols, HDR_TAG_SUFFIX)
    ptsCatCol = FindHeader(ptsHeaders, numPtsCols, HDR_POINT_CAT)

    ' --- Build replacement token lookup ---
    ' Parse RPL_TOKENS and find each token's column in the Device List.
    ' Tokens whose column doesn't exist in the Device List are silently skipped.
    tokens = Split(RPL_TOKENS, ";")
    numTokens = 0
    ReDim tokenText(0 To UBound(tokens))
    ReDim tokenDevCol(0 To UBound(tokens))

    For t = 0 To UBound(tokens)
        tok = Trim(tokens(t))
        If Len(tok) > 0 Then
            If dictDevHdr.Exists(UCase(tok)) Then
                tokenText(numTokens) = tok
                tokenDevCol(numTokens) = CLng(dictDevHdr(UCase(tok)))
                numTokens = numTokens + 1
            End If
        End If
    Next t

    ' --- Build template dictionary: I/O TEMPLATE -> Collection of row indices ---
    Set dictTemplate = CreateObject("Scripting.Dictionary")
    dictTemplate.CompareMode = vbTextCompare
    For r = 1 To tmpRows
        tKey = SafeStr(tmpData, r, tmpKeyCol)
        If Len(tKey) > 0 Then
            If Not dictTemplate.Exists(tKey) Then Set dictTemplate(tKey) = New Collection
            dictTemplate(tKey).Add r
        End If
    Next r

    ' --- Build category dictionary: POINT CATEGORY -> row index ---
    Set dictCategory = CreateObject("Scripting.Dictionary")
    dictCategory.CompareMode = vbTextCompare
    If catRows > 0 Then
        catKeyCol = 0
        If dictCatHdr.Exists(UCase(HDR_POINT_CAT)) Then
            catKeyCol = CLng(dictCatHdr(UCase(HDR_POINT_CAT)))
        End If
        If catKeyCol > 0 Then
            For r = 1 To catRows
                catKey = SafeStr(catData, r, catKeyCol)
                If Len(catKey) > 0 Then
                    If Not dictCategory.Exists(catKey) Then dictCategory.Add catKey, r
                End If
            Next r
        End If
    End If

    ' --- Pack everything into the context dictionary ---
    ctx.Add "devData", devData:          ctx.Add "devRows", CLng(devRows)
    ctx.Add "devCols", CLng(devCols):    ctx.Add "tmpData", tmpData
    ctx.Add "tmpRows", CLng(tmpRows):    ctx.Add "tmpCols", CLng(tmpCols)
    ctx.Add "catData", catData:          ctx.Add "catRows", CLng(catRows)
    ctx.Add "catCols", CLng(catCols):    ctx.Add "ptsHeaders", ptsHeaders
    ctx.Add "numPtsCols", CLng(numPtsCols)
    ctx.Add "devKeyCol", devKeyCol:      ctx.Add "devNameCol", devNameCol
    ctx.Add "tmpKeyCol", tmpKeyCol:      ctx.Add "tmpCatCol", tmpCatCol
    ctx.Add "colSource", colSource:      ctx.Add "colSourceIdx", colSourceIdx
    ctx.Add "ptsDescCol", ptsDescCol:    ctx.Add "ptsSuffixCol", ptsSuffixCol
    ctx.Add "ptsCatCol", ptsCatCol
    ctx.Add "numTokens", CLng(numTokens)
    ctx.Add "tokenText", tokenText:      ctx.Add "tokenDevCol", tokenDevCol
    ctx.Add "dictTemplate", dictTemplate
    ctx.Add "dictCategory", dictCategory

    Set LoadContext = ctx
    Exit Function

Fail:
    Set LoadContext = Nothing

End Function


' ============================================================================
'  PRIVATE: BuildExpectedArray
' ============================================================================
' Generates the output array by joining Device List x Template Table x
' Category Table.  Core generation engine used by all public macros.
'
' PARAMETERS:
'   ctx             - context Dictionary from LoadContext()
'   dictSkip        - hash set of IED NAMEs to skip (pass empty to include all)
'   outArr()        - (output) the generated 2D Variant array
'   totalRows       - (output) number of rows generated
'   skippedExisting - (output) devices skipped because they are in dictSkip
'   skippedNoTpl    - (output) devices skipped because I/O TEMPLATE is blank
'   devicesAdded    - (optional output) Collection of IED NAMEs included
'
' ALGORITHM - TWO PASSES:
'   PASS 1 counts output rows so we can ReDim to the exact size upfront,
'   avoiding ReDim Preserve (which copies the entire array on every resize).
'   PASS 2 populates the array.
' ============================================================================
Private Sub BuildExpectedArray(ByVal ctx As Object, _
                               ByVal dictSkip As Object, _
                               ByRef outArr() As Variant, _
                               ByRef totalRows As Long, _
                               ByRef skippedExisting As Long, _
                               ByRef skippedNoTpl As Long, _
                               Optional ByRef devicesAdded As Collection = Nothing)

    ' Unpack context into local variables for cleaner code and faster access
    ' (local variables are faster than Dictionary lookups inside tight loops).
    Dim devData      As Variant:     devData = ctx("devData")
    Dim devRows      As Long:        devRows = CLng(ctx("devRows"))
    Dim tmpData      As Variant:     tmpData = ctx("tmpData")
    Dim catData      As Variant:     catData = ctx("catData")
    Dim numPtsCols   As Long:        numPtsCols = CLng(ctx("numPtsCols"))
    Dim devKeyCol    As Long:        devKeyCol = CLng(ctx("devKeyCol"))
    Dim devNameCol   As Long:        devNameCol = CLng(ctx("devNameCol"))
    Dim tmpCatCol    As Long:        tmpCatCol = CLng(ctx("tmpCatCol"))
    Dim ptsDescCol   As Long:        ptsDescCol = CLng(ctx("ptsDescCol"))
    Dim ptsSuffixCol As Long:        ptsSuffixCol = CLng(ctx("ptsSuffixCol"))
    Dim numTokens    As Long:        numTokens = CLng(ctx("numTokens"))
    Dim colSource()  As String:      colSource = ctx("colSource")
    Dim colSourceIdx() As Long:      colSourceIdx = ctx("colSourceIdx")
    Dim tokenText()  As String:      tokenText = ctx("tokenText")
    Dim tokenDevCol() As Long:       tokenDevCol = ctx("tokenDevCol")
    Dim dictTemplate As Object:      Set dictTemplate = ctx("dictTemplate")
    Dim dictCategory As Object:      Set dictCategory = ctx("dictCategory")

    Dim tKey        As String
    Dim iedName     As String
    Dim tagSuffix   As String
    Dim descStr     As String
    Dim ptCatVal    As String
    Dim rplValues() As String
    Dim tmpRowIdx   As Variant  ' Collection item (Long wrapped as Variant via For Each)
    Dim dictSeen    As Object
    Dim outRow      As Long
    Dim catRowIdx   As Long
    Dim tRow        As Long
    Dim col         As Long
    Dim r           As Long
    Dim t           As Long

    ' ---- PASS 1: Count output rows ----
    totalRows = 0
    skippedExisting = 0
    skippedNoTpl = 0

    For r = 1 To devRows
        tKey = SafeStr(devData, r, devKeyCol)
        If Len(tKey) = 0 Then
            skippedNoTpl = skippedNoTpl + 1
            GoTo CntSkip
        End If

        iedName = SafeStr(devData, r, devNameCol)
        If Len(iedName) > 0 Then
            If dictSkip.Exists(iedName) Then
                skippedExisting = skippedExisting + 1
                GoTo CntSkip
            End If
        End If

        If dictTemplate.Exists(tKey) Then
            totalRows = totalRows + dictTemplate(tKey).Count
        End If
CntSkip:
    Next r

    If totalRows = 0 Then Exit Sub

    ' ---- PASS 2: Populate output array ----
    ReDim outArr(1 To totalRows, 1 To numPtsCols)

    ' Pre-allocate replacement cache (only if tokens exist).
    If numTokens > 0 Then ReDim rplValues(0 To numTokens - 1)

    ' Seen-set prevents duplicate names in devicesAdded when a device has
    ' multiple template rows.
    Set dictSeen = CreateObject("Scripting.Dictionary")
    dictSeen.CompareMode = vbTextCompare

    outRow = 0

    For r = 1 To devRows
        ' Same skip logic as Pass 1
        tKey = SafeStr(devData, r, devKeyCol)
        If Len(tKey) = 0 Then GoTo NxtDev
        iedName = SafeStr(devData, r, devNameCol)
        If Len(iedName) > 0 Then
            If dictSkip.Exists(iedName) Then GoTo NxtDev
        End If
        If Not dictTemplate.Exists(tKey) Then GoTo NxtDev

        ' Track this device name if the caller wants a list of devices added.
        If Not devicesAdded Is Nothing Then
            If Len(iedName) > 0 Then
                If Not dictSeen.Exists(iedName) Then
                    dictSeen.Add iedName, True
                    devicesAdded.Add iedName
                End If
            End If
        End If

        ' Cache replacement token values for this device.
        ' Done once per device (not per template row) since tokens are
        ' device-level attributes (RPLIED, RPL1, RPL2, etc.).
        If numTokens > 0 Then
            For t = 0 To numTokens - 1
                rplValues(t) = SafeStr(devData, r, tokenDevCol(t))
            Next t
        End If

        ' For each template row matching this device's I/O TEMPLATE...
        For Each tmpRowIdx In dictTemplate(tKey)
            tRow = CLng(tmpRowIdx)
            outRow = outRow + 1

            ' Resolve the category row for this template point.
            ' catRowIdx=0 means no category match; category columns left blank.
            catRowIdx = 0
            If tmpCatCol > 0 Then
                ptCatVal = SafeStr(tmpData, tRow, tmpCatCol)
                If Len(ptCatVal) > 0 Then
                    If dictCategory.Exists(ptCatVal) Then
                        catRowIdx = CLng(dictCategory(ptCatVal))
                    End If
                End If
            End If

            ' Fill each output column based on the column map
            For col = 1 To numPtsCols
                Select Case colSource(col)

                    Case "D"  ' Device List column
                        If Not IsError(devData(r, colSourceIdx(col))) Then
                            outArr(outRow, col) = devData(r, colSourceIdx(col))
                        End If

                    Case "T"  ' Template Table column
                        If Not IsError(tmpData(tRow, colSourceIdx(col))) Then
                            outArr(outRow, col) = tmpData(tRow, colSourceIdx(col))
                        End If

                    Case "C"  ' Category Table column
                        If catRowIdx > 0 Then
                            If Not IsError(catData(catRowIdx, colSourceIdx(col))) Then
                                outArr(outRow, col) = catData(catRowIdx, colSourceIdx(col))
                            End If
                        End If

                    Case "X"  ' Derived: TAG NAME = IED NAME + "_" + TAG SUFFIX
                        ' Read the suffix from the output array (not tmpData) to ensure
                        ' we use the same value that was actually written to the output.
                        tagSuffix = ""
                        If ptsSuffixCol > 0 Then
                            If Not IsEmpty(outArr(outRow, ptsSuffixCol)) Then
                                tagSuffix = SafeStrVal(outArr(outRow, ptsSuffixCol))
                            End If
                        End If
                        If Len(iedName) > 0 And Len(tagSuffix) > 0 Then
                            outArr(outRow, col) = iedName & TAG_NAME_SEP & tagSuffix
                        ElseIf Len(iedName) > 0 Then
                            outArr(outRow, col) = iedName
                        End If

                End Select
            Next col

            ' PLACEHOLDER REPLACEMENT in DESCRIPTION column.
            ' Replace literal token text (e.g. RPLIED, RPL1) with device-specific values.
            If ptsDescCol > 0 And numTokens > 0 Then
                If Not IsEmpty(outArr(outRow, ptsDescCol)) Then
                    If Not IsError(outArr(outRow, ptsDescCol)) Then
                        descStr = SafeStrVal(outArr(outRow, ptsDescCol))
                        If Len(descStr) > 0 Then
                            For t = 0 To numTokens - 1
                                If InStr(1, descStr, tokenText(t), vbTextCompare) > 0 Then
                                    descStr = Replace(descStr, tokenText(t), _
                                                      rplValues(t), 1, -1, vbTextCompare)
                                End If
                            Next t
                            outArr(outRow, ptsDescCol) = descStr
                        End If
                    End If
                End If
            End If

        Next tmpRowIdx
NxtDev:
    Next r

End Sub


' ============================================================================
'  PRIVATE: BuildExistingDeviceDict
' ============================================================================
' Returns a Dictionary (used as a hash set) of all IED NAME values currently
' in the Points List.  GeneratePointsList uses this to skip already-generated
' devices, preventing duplicates.
' ============================================================================
Private Function BuildExistingDeviceDict() As Object

    Dim dict       As Object
    Dim loPoints   As ListObject
    Dim ptsHeaders() As String
    Dim numCols    As Long
    Dim nameCol    As Long
    Dim ptsRaw     As Variant
    Dim cellVal    As Variant
    Dim nm         As String
    Dim r          As Long

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Set loPoints = ThisWorkbook.Worksheets(SH_POINTS).ListObjects(TBL_POINTS)

    ' Guard 1: table has no data range at all
    If loPoints.DataBodyRange Is Nothing Then
        Set BuildExistingDeviceDict = dict
        Exit Function
    End If

    ' Guard 2: table reports zero rows (placeholder row edge case)
    If loPoints.ListRows.Count = 0 Then
        Set BuildExistingDeviceDict = dict
        Exit Function
    End If

    ReadHeaders loPoints, ptsHeaders, numCols
    nameCol = FindHeader(ptsHeaders, numCols, HDR_IED_NAME)
    If nameCol = 0 Then
        Set BuildExistingDeviceDict = dict
        Exit Function
    End If

    ' Read all data in one bulk operation
    ptsRaw = loPoints.DataBodyRange.Value

    ' Guard 3: single-cell table returns a scalar, not an array
    If Not IsArray(ptsRaw) Then
        Set BuildExistingDeviceDict = dict
        Exit Function
    End If

    For r = 1 To UBound(ptsRaw, 1)
        cellVal = ptsRaw(r, nameCol)
        ' Triple-guard: Error variants can't be CStr'd; Null throws on CStr in some locales.
        If Not IsError(cellVal) Then
            If Not IsEmpty(cellVal) Then
                If Not IsNull(cellVal) Then
                    nm = Trim(CStr(cellVal))
                    If Len(nm) > 0 Then
                        If Not dict.Exists(nm) Then dict.Add nm, True
                    End If
                End If
            End If
        End If
    Next r

    Set BuildExistingDeviceDict = dict

End Function


' ============================================================================
'  PRIVATE UTILITY FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' SafeStr - Extract a cell from a 2D Variant array as a trimmed String.
'
' Cell values from Range.Value can be any subtype: String, Double, Date,
' Boolean, Empty, Null, or Error (#N/A etc.).  Calling CStr() on an Error
' or Null throws a Type Mismatch (Err 13).  SafeStr checks all problematic
' subtypes before converting, with On Error Resume Next as a final safety net.
'
' Returns the trimmed string value, or "" for any non-string-convertible value.
' ----------------------------------------------------------------------------
Private Function SafeStr(ByRef arr As Variant, ByVal r As Long, ByVal c As Long) As String
    Dim v As Variant
    On Error Resume Next
    v = arr(r, c)
    If Err.Number <> 0 Then SafeStr = "": Err.Clear: Exit Function
    On Error GoTo 0
    SafeStr = ""
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    On Error Resume Next
    SafeStr = Trim(CStr(v))
    If Err.Number <> 0 Then SafeStr = "": Err.Clear
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' SafeStrVal - Same as SafeStr but for a standalone Variant value (not
' indexed from a 2D array).  Used when reading from the output array (outArr).
' ----------------------------------------------------------------------------
Private Function SafeStrVal(ByVal v As Variant) As String
    SafeStrVal = ""
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    On Error Resume Next
    SafeStrVal = Trim(CStr(v))
    If Err.Number <> 0 Then SafeStrVal = "": Err.Clear
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' ReadHeaders - Read ListObject headers into a 1-based String() array.
'
' Using String() instead of Variant() avoids type ambiguity when calling
' Trim() and UCase() later.
'
' Edge case: a single-column table returns a scalar from .Value, not a 2D
' array.  Both cases are handled.
' ----------------------------------------------------------------------------
Private Sub ReadHeaders(ByVal lo As ListObject, ByRef outHeaders() As String, ByRef outCols As Long)
    Dim rawHdr As Variant
    Dim c      As Long
    rawHdr = lo.HeaderRowRange.Value
    If Not IsArray(rawHdr) Then
        ' Single-column table - scalar value
        outCols = 1
        ReDim outHeaders(1 To 1)
        outHeaders(1) = SafeStrVal(rawHdr)
    Else
        ' Normal multi-column table - 2D array (1 row x N cols)
        outCols = UBound(rawHdr, 2)
        ReDim outHeaders(1 To outCols)
        For c = 1 To outCols
            outHeaders(c) = SafeStrVal(rawHdr(1, c))
        Next c
    End If
End Sub

' ----------------------------------------------------------------------------
' ReadData - Read ListObject data body into a 2D Variant array.
'
' Returns outRows=0 and outData=Empty if the table is empty.
'
' Edge case: a single-cell data body returns a scalar; we wrap it into a
' 1x1 2D array so all downstream code can safely use arr(r, c) syntax.
' ----------------------------------------------------------------------------
Private Sub ReadData(ByVal lo As ListObject, ByRef outData As Variant, ByRef outRows As Long)
    Dim rawData As Variant
    Dim tmpArr(1 To 1, 1 To 1) As Variant
    outRows = 0
    outData = Empty
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If lo.ListRows.Count = 0 Then Exit Sub
    rawData = lo.DataBodyRange.Value
    If Not IsArray(rawData) Then
        ' Single cell - wrap into 2D array
        tmpArr(1, 1) = rawData
        outRows = 1
        outData = tmpArr
    Else
        outRows = UBound(rawData, 1)
        outData = rawData
    End If
End Sub

' ----------------------------------------------------------------------------
' FindHeader - Search a String() header array for a target header name.
'
' Returns the 1-based column index, or 0 if not found.
' Comparison is case-insensitive via UCase().
' ----------------------------------------------------------------------------
Private Function FindHeader(ByRef headers() As String, ByVal numCols As Long, _
                             ByVal target As String) As Long
    Dim c  As Long
    Dim uT As String
    uT = UCase(target)
    For c = 1 To numCols
        If UCase(Trim(headers(c))) = uT Then
            FindHeader = c
            Exit Function
        End If
    Next c
    FindHeader = 0
End Function

' ----------------------------------------------------------------------------
' BuildHeaderDict - Build a Dictionary mapping UCASE(header) -> Long index.
'
' CompareMode = vbTextCompare makes the dictionary case-insensitive.
' CLng(c) stores values as Long (not Variant) to avoid type issues when
' used as array indices.  First occurrence wins for duplicate headers.
' ----------------------------------------------------------------------------
Private Function BuildHeaderDict(ByRef headers() As String, ByVal numCols As Long) As Object
    Dim dict As Object
    Dim c    As Long
    Dim key  As String
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    For c = 1 To numCols
        key = UCase(Trim(headers(c)))
        If Len(key) > 0 Then
            If Not dict.Exists(key) Then dict.Add key, CLng(c)
        End If
    Next c
    Set BuildHeaderDict = dict
End Function

' ----------------------------------------------------------------------------
' BuildAliasDict - Parse the HEADER_ALIASES constant into a Dictionary.
'
' Maps UCASE(targetHeader) -> UCASE(sourceHeader).
' Handles cases where the Points List uses a different column name than the
' source table for the same data.
'
' Format: "TARGET1|SOURCE1;TARGET2|SOURCE2;..."
' ----------------------------------------------------------------------------
Private Function BuildAliasDict() As Object
    Dim dict    As Object
    Dim pairs() As String
    Dim parts() As String
    Dim tgtKey  As String
    Dim i       As Long

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    If Len(HEADER_ALIASES) = 0 Then
        Set BuildAliasDict = dict
        Exit Function
    End If

    pairs = Split(HEADER_ALIASES, ";")
    For i = LBound(pairs) To UBound(pairs)
        parts = Split(Trim(pairs(i)), "|")
        If UBound(parts) >= 1 Then
            tgtKey = UCase(Trim(parts(0)))
            If Not dict.Exists(tgtKey) Then dict.Add tgtKey, UCase(Trim(parts(1)))
        End If
    Next i

    Set BuildAliasDict = dict
End Function

' ----------------------------------------------------------------------------
' GetOrCreateSheet - Return a worksheet by name, creating it if needed.
' Used for the Compare and Orphaned Points output sheets.
' ----------------------------------------------------------------------------
Private Function GetOrCreateSheet(ByVal shName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = shName
    End If
    Set GetOrCreateSheet = ws
End Function

' ----------------------------------------------------------------------------
' ProgressBar - Build a visual progress bar string for the Excel status bar.
'
' Returns a string like:  [████████████░░░░░░░░] 60%
'
' Uses Unicode block characters so the bar renders in the status bar font:
'   ChrW(9608) = █  full block  (completed portion)
'   ChrW(9617) = ░  light shade (remaining portion)
'
' BAR_WIDTH controls the total number of block characters.  25 fits
' comfortably alongside a row-count label in the status bar.
' ----------------------------------------------------------------------------
Private Function ProgressBar(ByVal current As Long, ByVal total As Long) As String
    Const BAR_WIDTH As Long = 25
    Dim pct    As Double
    Dim filled As Long

    If total <= 0 Then
        ProgressBar = "[" & String(BAR_WIDTH, ChrW(9617)) & "]   0%"
        Exit Function
    End If

    pct = current / total
    filled = CLng(Int(pct * BAR_WIDTH))
    If filled > BAR_WIDTH Then filled = BAR_WIDTH

    ProgressBar = "[" & String(filled, ChrW(9608)) & _
                  String(BAR_WIDTH - filled, ChrW(9617)) & "] " & _
                  Format(pct * 100, "0") & "%"
End Function

' ----------------------------------------------------------------------------
' PerformanceOn / PerformanceOff - Performance envelope.
'
' ScreenUpdating = False prevents Excel from repainting after every cell
' change - typically the single largest performance gain.
'
' Calculation = xlCalculationManual prevents formula recalculation after
' every write.  Switching back to xlCalculationAutomatic in PerformanceOff
' triggers a single recalc when the operation completes.
'
' EnableEvents = False prevents Worksheet_Change and similar event handlers
' from firing during bulk writes.
'
' IMPORTANT: PerformanceOff must be called in EVERY exit path (CleanExit and
' ErrHandler) to restore Excel to normal operating mode.  Using a GoTo label
' named "PerformanceOff:" does NOT call this subroutine - always call it
' explicitly.
' ----------------------------------------------------------------------------
Private Sub PerformanceOn()
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayStatusBar = True
    End With
End Sub

Private Sub PerformanceOff()
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .StatusBar = False
    End With
End Sub
