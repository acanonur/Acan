Attribute VB_Name = "ACE_BOM_Import"
'==============================================================================
' ACE_BOM_Import
'------------------------------------------------------------------------------
' Imports a CATIA BOM extraction workbook (e.g. "flint bom 2.xlsx", produced by
' the CATIA_BOM_Extractor_* macros) into the ACE cost model workbook.
'
' The macro does two things in one run:
'
'   1) It copies the complete extraction - all columns, all formatting and all
'      thumbnails - into a NEW sheet of the ACE workbook ("BOM Extract"), so the
'      raw extraction stays available inside the cost model.
'
'   2) It fills the "Input" sheet of the ACE workbook from that extraction:
'
'          BOM extraction column      ->   Input sheet column
'          ------------------------------------------------------
'          Level                      ->   A  Level
'          Part Number                ->   C  Part Number
'          Description                ->   D  Name
'          First Level                ->   E  Other Reference
'          Thumbnail (picture)        ->   F  Picture
'          Qty                        ->   G  Qty System
'
'      The columns are found BY HEADER NAME, so extra columns of the newer
'      extractions ("First Level", "Type", material data, ...) do no harm and
'      the column order may change. Rows of sub-assemblies and of single
'      bodies are imported like any other row; the indentation the extractor
'      writes in front of a body name is removed.
'
'      B Lookup Key stays empty - the extraction has nothing to fill it with.
'      Everything else (costs, factory, material, process, all formulas) is
'      left untouched.
'
' Entry points (Alt+F8):
'
'      ImportBOM                - does both steps (normal use)
'      ImportBOM_PartsOnly      - same, but the "Assembly" rows of a structure
'                                 sheet are skipped: only the parts that carry
'                                 cost reach the Input sheet
'      ImportBOM_SheetOnly      - only creates the "BOM Extract" copy sheet
'      ImportBOM_InputOnly      - only fills the Input sheet
'      ClearBOMImport           - removes imported values + pictures from Input
'      FitPicturesToCells       - re-fits all thumbnails into their Picture
'                                 cells (after changing row heights e.g.)
'      ResetInputSheet          - empties the whole Input sheet for a new
'                                 estimate; all formulas stay untouched
'      NewInputSheet            - creates a new, empty copy of the Input sheet
'                                 (formulas + formatting kept, values removed)
'      CreateMacroButtons       - puts the two push buttons [ Import BOM ] and
'                                 [ Clean Input ] on the Input sheet
'
' Installation:
'      Excel -> Alt+F11 -> File -> Import File... -> ACE_BOM_Import.bas
'      Save the ACE workbook as .xlsm.
'
'==============================================================================
Option Explicit

'--- ACE "Input" sheet layout -------------------------------------------------
Private Const INPUT_SHEET As String = "Input"
Private Const INPUT_HEADER_ROW As Long = 14
Private Const INPUT_FIRST_DATA_ROW As Long = 15

Private Const COL_LEVEL As Long = 1          ' A - Level
Private Const COL_PARTNUM As Long = 3        ' C - Part Number
Private Const COL_NAME As Long = 4           ' D - Name
Private Const COL_OTHERREF As Long = 5       ' E - Other Reference
Private Const COL_PICTURE As Long = 6        ' F - Picture
Private Const COL_QTY As Long = 7            ' G - Qty System

' The extractors write a "First Level" column (the first level node of the
' assembly a row belongs to). True = it is imported into E "Other Reference",
' so the Input sheet can be filtered per first level as well.
Private Const IMPORT_FIRST_LEVEL As Boolean = True

' The extractors write a "Type" column (Assembly / Part / Body). True = the
' rows of sub-assemblies are written in bold in the Input sheet, so structure
' rows can be told apart from the parts that carry the cost.
Private Const MARK_ASSEMBLY_ROWS As Boolean = True

' Column that carries a formula on every prepared template row. Used to find out
' how many prepared rows the Input sheet has (column I = "Total Cost").
Private Const TEMPLATE_PROBE_COL As String = "I"

' Entry columns of a BOM line. They are emptied before every import and by
' ClearBOMImport, so that nothing of the previous BOM stays behind:
'   A  Level                B  Lookup Key           C  Part Number
'   D  Name                 E  Other Reference      F  Picture
'   G  Qty System           P  Material Reference   Q  Description / Comments
'   AC Process Reference    AD Process Description  AH Process Time
'   AI Direct Operator      AN Qty Operators setup  AO Set Up Time / Lot Size
'   AP Other Setup Cost
' Cells containing a formula are never removed, only entered values.
' Note: emptying B (Lookup Key), P (Material Reference) and Q (Comments) makes
' the VLOOKUPs of the columns R / S / T show #N/A until new entries are made -
' that is the same as deleting those cells by hand.
Private Const CLEAR_COLUMNS As String = "A,B,C,D,E,F,G,P,Q,AC,AD,AH,AI,AN,AO,AP"

'--- Behaviour ----------------------------------------------------------------
Private Const BOM_SHEET_NAME As String = "BOM Extract"
Private Const WRITE_LEVEL As Boolean = True  ' False = keep the template levels
Private Const COPY_PICTURES As Boolean = True
Private Const PIC_TAG As String = "BOMPIC_"  ' name prefix of imported pictures
Private Const PIC_MARGIN As Double = 2#      ' points of free space inside cell
Private Const MIN_PIC_ROW_HEIGHT As Double = 45#

' Columns that keep their template defaults when the Input sheet is emptied.
' They are entered values as well, but the model needs them to stay valid:
'   K  Company (AGG)    L  Country (DEU)   AA Process Type (None)
'   AE Lot Size (1)     BC Valuation Method (Industrial)
' Emptying AE would produce #DIV/0! in the setup cost (it is a divisor),
' emptying K would set all factory rates to 0. Add or remove letters here to
' change what a reset keeps.
Private Const RESET_KEEP_COLUMNS As String = "K,L,AA,AE,BC"

'--- Push buttons on the Input sheet (CreateMacroButtons) ---------------------
Private Const BUTTON_ANCHOR_CELL As String = "F5"   ' top left corner of the first button
Private Const BUTTON_WIDTH As Double = 100#
Private Const BUTTON_HEIGHT As Double = 30#
Private Const BUTTON_GAP As Double = 8#
Private Const BTN_IMPORT_NAME As String = "btnImportBOM"
Private Const BTN_CLEAN_NAME As String = "btnCleanInput"

'--- Shape type constants (avoids a hard reference to the Office library) ------
Private Const SHP_PICTURE As Long = 13
Private Const SHP_LINKED_PICTURE As Long = 11

'--- Custom errors ------------------------------------------------------------
Private Const ERR_NO_PARTNUM As Long = vbObjectError + 1
Private Const ERR_USER_CANCEL As Long = vbObjectError + 2
Private Const ERR_NO_ROWS As Long = vbObjectError + 3

'--- Run mode -----------------------------------------------------------------
' Set by the entry points: True = the "Assembly" rows of a structure sheet are
' skipped, so only the parts that carry cost land in the Input sheet.
Private mSkipAssemblyRows As Boolean
Private mFilteredOut As Long

'--- Saved application state ---------------------------------------------------
Private mScreen As Boolean
Private mEvents As Boolean
Private mCalc As XlCalculation
Private mAlerts As Boolean
Private mStateSaved As Boolean


'==============================================================================
' ENTRY POINT: full import (copy sheet + fill Input sheet)
'==============================================================================
Public Sub ImportBOM()
    mSkipAssemblyRows = False
    RunImport True, True
End Sub

'==============================================================================
' ENTRY POINT: import a sheet of the assembly structure version, but WITHOUT
' its sub-assembly rows - only the parts that carry cost reach the Input sheet
'==============================================================================
Public Sub ImportBOM_PartsOnly()
    mSkipAssemblyRows = True
    RunImport True, True
End Sub

'==============================================================================
' ENTRY POINT: only create the separate copy of the extraction
'==============================================================================
Public Sub ImportBOM_SheetOnly()
    mSkipAssemblyRows = False
    RunImport True, False
End Sub

'==============================================================================
' ENTRY POINT: only fill the Input sheet
'==============================================================================
Public Sub ImportBOM_InputOnly()
    mSkipAssemblyRows = False
    RunImport False, True
End Sub


'==============================================================================
' MAIN WORKER
'
'   makeCopySheet : create the "BOM Extract" sheet holding the raw extraction
'   fillInput     : write Level / Part Number / Name / Picture / Qty to "Input"
'
' The extraction is always copied into this workbook first; the Input sheet is
' then filled from that local copy. That way the source workbook can be closed
' immediately and the thumbnails are copied inside a single workbook, which is
' considerably faster and more reliable than a cross-workbook copy.
'==============================================================================
Private Sub RunImport(ByVal makeCopySheet As Boolean, ByVal fillInput As Boolean)
    Dim srcPath As String
    Dim srcWb As Workbook, srcWs As Worksheet
    Dim bomWs As Worksheet
    Dim inputWs As Worksheet
    Dim srcWasOpen As Boolean
    Dim keepCopySheet As Boolean
    Dim hdrRow As Long, lastRow As Long
    Dim nRows As Long, nPics As Long, nSkipped As Long
    Dim wasProtected As Boolean
    Dim startedAt As Double
    Dim copySheetName As String

    On Error GoTo CleanFail

    If Not SheetExists(INPUT_SHEET) Then
        MsgBox "This workbook has no sheet named '" & INPUT_SHEET & "'." & vbCrLf & _
               "Run the macro from the ACE cost model workbook.", vbExclamation, "BOM Import"
        Exit Sub
    End If

    srcPath = PickBOMFile()
    If Len(srcPath) = 0 Then Exit Sub                 ' user cancelled

    startedAt = Timer
    SaveAppState

    '--- 1. open (or reuse) the extraction workbook ----------------------------
    Set srcWb = GetWorkbook(srcPath, srcWasOpen)
    If srcWb Is Nothing Then
        MsgBox "Could not open:" & vbCrLf & srcPath, vbCritical, "BOM Import"
        GoTo CleanExit
    End If

    Set srcWs = FindBOMSheet(srcWb)
    If srcWs Is Nothing Then
        MsgBox "No BOM data found in:" & vbCrLf & srcWb.Name & vbCrLf & vbCrLf & _
               "Expected a sheet with a 'Part Number' column header.", _
               vbExclamation, "BOM Import"
        GoTo CleanExit
    End If

    '--- 2. copy the whole extraction into this workbook ----------------------
    Application.StatusBar = "BOM import: copying extraction sheet ..."
    keepCopySheet = makeCopySheet
    Set bomWs = CreateCopySheet(srcWs, keepCopySheet)
    If bomWs Is Nothing Then GoTo CleanExit          ' user cancelled
    copySheetName = bomWs.Name

    If Not srcWasOpen Then srcWb.Close SaveChanges:=False
    Set srcWb = Nothing

    '--- 3. locate the data inside the copy -----------------------------------
    hdrRow = FindHeaderRow(bomWs)
    lastRow = LastDataRow(bomWs, hdrRow)
    nRows = lastRow - hdrRow

    If nRows <= 0 Then
        MsgBox "The extraction sheet contains no data rows.", vbExclamation, "BOM Import"
        If Not keepCopySheet Then DeleteSheet bomWs
        GoTo CleanExit
    End If

    '--- 4. fill the Input sheet ----------------------------------------------
    If fillInput Then
        Set inputWs = ThisWorkbook.Worksheets(INPUT_SHEET)
        wasProtected = UnprotectSheet(inputWs)
        If inputWs.ProtectContents Then
            MsgBox "The '" & INPUT_SHEET & "' sheet is protected with a password " & _
                   "and could not be unlocked." & vbCrLf & _
                   "Unprotect it manually and run the macro again.", _
                   vbExclamation, "BOM Import"
            If Not keepCopySheet Then DeleteSheet bomWs
            GoTo CleanExit
        End If

        FillInputSheet bomWs, hdrRow, inputWs, nRows, nPics, nSkipped

        If wasProtected Then ProtectSheet inputWs
    End If

    '--- 5. drop the working copy if it was not requested ----------------------
    If Not keepCopySheet Then DeleteSheet bomWs

    RestoreAppState
    Application.Calculate

    MsgBox BuildReport(makeCopySheet, fillInput, copySheetName, keepCopySheet, _
                       nRows, nPics, nSkipped, Timer - startedAt), _
           vbInformation, "BOM Import"
    Exit Sub

CleanExit:
    If Not srcWb Is Nothing Then
        If Not srcWasOpen Then srcWb.Close SaveChanges:=False
    End If
    RestoreAppState
    Exit Sub

CleanFail:
    Dim errNo As Long, msg As String
    errNo = Err.Number
    msg = "BOM import failed:" & vbCrLf & vbCrLf & _
          "Error " & Err.Number & " - " & Err.Description

    On Error Resume Next
    If Not srcWb Is Nothing Then
        If Not srcWasOpen Then srcWb.Close SaveChanges:=False
    End If
    If Not keepCopySheet Then DeleteSheet bomWs      ' remove the temporary copy
    On Error GoTo 0

    RestoreAppState
    If errNo = ERR_USER_CANCEL Then
        MsgBox "BOM import cancelled - nothing was changed.", vbInformation, "BOM Import"
    ElseIf errNo = ERR_NO_ROWS Then
        MsgBox "Nothing to import: every row of that extraction is a sub-assembly row." & vbCrLf & _
               "Use ImportBOM (instead of ImportBOM_PartsOnly) for a structure sheet.", _
               vbExclamation, "BOM Import"
    Else
        MsgBox msg, vbCritical, "BOM Import"
    End If
End Sub


'==============================================================================
' Writes Level / Part Number / Name / Qty and the thumbnails into "Input".
'==============================================================================
Private Sub FillInputSheet(bomWs As Worksheet, ByVal hdrRow As Long, inputWs As Worksheet, _
                           ByRef nRows As Long, ByRef nPics As Long, ByRef nSkipped As Long)
    Dim cLevel As Long, cPart As Long, cDesc As Long, cQty As Long, cFirst As Long, cType As Long
    Dim preparedLastRow As Long, capacity As Long
    Dim levels As Variant, parts As Variant, descs As Variant, qtys As Variant
    Dim firsts As Variant, types As Variant, rawTypes As Variant
    Dim nSrc As Long, idx() As Long, keepRow As Boolean
    Dim pics As Object, knownNames As Object
    Dim shp As Object
    Dim i As Long, srcRow As Long, tgtRow As Long
    Dim answer As VbMsgBoxResult
    Dim prevActive As Object

    cPart = FindColumn(bomWs, hdrRow, "part number", "partnumber", "part no", "part-number")
    cLevel = FindColumn(bomWs, hdrRow, "level")
    cDesc = FindColumn(bomWs, hdrRow, "description", "name", "designation")
    cQty = FindColumn(bomWs, hdrRow, "qty", "quantity", "qty system", "quantity total")
    cFirst = 0
    If IMPORT_FIRST_LEVEL Then
        cFirst = FindColumn(bomWs, hdrRow, "first level", "firstlevel", "first-level", _
                            "top level", "toplevel", "main assembly")
    End If
    cType = 0
    If MARK_ASSEMBLY_ROWS Then cType = FindColumn(bomWs, hdrRow, "type", "row type", "kind")

    If cPart = 0 Then Err.Raise ERR_NO_PARTNUM, , "No 'Part Number' column found in the extraction."

    '--- which source rows are imported? --------------------------------------
    ' (the assembly rows of a structure sheet can be skipped)
    nSrc = nRows
    If cType > 0 Then rawTypes = ReadColumn(bomWs, cType, hdrRow + 1, hdrRow + nSrc, False)

    ReDim idx(1 To nSrc)
    nRows = 0
    mFilteredOut = 0
    For i = 1 To nSrc
        keepRow = True
        If mSkipAssemblyRows And cType > 0 Then
            If StrComp(Trim$(CStr(rawTypes(i, 1))), "Assembly", vbTextCompare) = 0 Then keepRow = False
        End If
        If keepRow Then
            nRows = nRows + 1
            idx(nRows) = i
        Else
            mFilteredOut = mFilteredOut + 1
        End If
    Next i

    If nRows = 0 Then Err.Raise ERR_NO_ROWS, , "Nothing left to import - every row of the " & _
                                               "extraction is a sub-assembly row."

    '--- how many prepared rows does the template offer? ----------------------
    preparedLastRow = LastTemplateRow(inputWs)
    capacity = preparedLastRow - INPUT_FIRST_DATA_ROW + 1

    If nRows > capacity Then
        answer = MsgBox("The extraction has " & nRows & " rows but the Input sheet only has " & _
                        capacity & " prepared rows (" & INPUT_FIRST_DATA_ROW & ":" & preparedLastRow & ")." & vbCrLf & vbCrLf & _
                        "Yes  = extend the sheet by copying the last prepared row down" & vbCrLf & _
                        "No   = import the first " & capacity & " rows only" & vbCrLf & _
                        "Cancel = abort", vbQuestion + vbYesNoCancel, "BOM Import")
        Select Case answer
            Case vbCancel
                Err.Raise ERR_USER_CANCEL, , "Import cancelled by the user."
            Case vbYes
                If ExtendTemplate(inputWs, preparedLastRow, INPUT_FIRST_DATA_ROW + nRows - 1) Then
                    preparedLastRow = INPUT_FIRST_DATA_ROW + nRows - 1
                    capacity = nRows
                End If
        End Select
        If nRows > capacity Then
            nSkipped = nRows - capacity
            nRows = capacity
        End If
    End If

    '--- read the extraction into arrays --------------------------------------
    ' PickRows keeps only the source rows of idx(), so a truncated or filtered
    ' import stays row-aligned with the pictures further down
    parts = PickRows(ReadColumn(bomWs, cPart, hdrRow + 1, hdrRow + nSrc, False), idx, nRows)
    If cLevel > 0 Then levels = PickRows(ReadColumn(bomWs, cLevel, hdrRow + 1, hdrRow + nSrc, True), idx, nRows)
    If cDesc > 0 Then descs = PickRows(ReadColumn(bomWs, cDesc, hdrRow + 1, hdrRow + nSrc, False), idx, nRows)
    If cQty > 0 Then qtys = PickRows(ReadColumn(bomWs, cQty, hdrRow + 1, hdrRow + nSrc, True), idx, nRows)
    If cFirst > 0 Then firsts = PickRows(ReadColumn(bomWs, cFirst, hdrRow + 1, hdrRow + nSrc, False), idx, nRows)
    If cType > 0 Then types = PickRows(rawTypes, idx, nRows)

    '--- clear the previous import --------------------------------------------
    Application.StatusBar = "BOM import: clearing previous import ..."
    ClearInputRange inputWs, INPUT_FIRST_DATA_ROW, preparedLastRow

    '--- write the values (one shot per column) -------------------------------
    Application.StatusBar = "BOM import: writing " & nRows & " rows ..."
    inputWs.Cells(INPUT_FIRST_DATA_ROW, COL_PARTNUM).Resize(nRows, 1).Value = parts
    If WRITE_LEVEL And cLevel > 0 Then _
        inputWs.Cells(INPUT_FIRST_DATA_ROW, COL_LEVEL).Resize(nRows, 1).Value = levels
    If cDesc > 0 Then _
        inputWs.Cells(INPUT_FIRST_DATA_ROW, COL_NAME).Resize(nRows, 1).Value = descs
    If cQty > 0 Then _
        inputWs.Cells(INPUT_FIRST_DATA_ROW, COL_QTY).Resize(nRows, 1).Value = qtys
    If cFirst > 0 Then _
        inputWs.Cells(INPUT_FIRST_DATA_ROW, COL_OTHERREF).Resize(nRows, 1).Value = firsts

    '--- mark the sub-assembly rows -------------------------------------------
    ' the whole block is reset first, so a re-import never leaves an old row bold
    If cType > 0 Then
        inputWs.Cells(INPUT_FIRST_DATA_ROW, COL_LEVEL).Resize(nRows, COL_QTY).Font.Bold = False
        For i = 1 To nRows
            If StrComp(Trim$(CStr(types(i, 1))), "Assembly", vbTextCompare) = 0 Then
                inputWs.Cells(INPUT_FIRST_DATA_ROW + i - 1, COL_LEVEL).Resize(1, COL_QTY).Font.Bold = True
            End If
        Next i
    End If

    '--- thumbnails ------------------------------------------------------------
    If Not COPY_PICTURES Then Exit Sub

    Set pics = PictureMap(bomWs)
    If pics.Count = 0 Then Exit Sub

    Set prevActive = ActiveSheet
    ThisWorkbook.Activate
    inputWs.Activate

    ' names of the shapes that are already there - used to identify each paste
    Set knownNames = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    For Each shp In inputWs.Shapes
        If Not knownNames.Exists(shp.Name) Then knownNames.Add shp.Name, True
    Next shp
    Err.Clear
    On Error GoTo 0

    For i = 1 To nRows
        srcRow = hdrRow + idx(i)
        tgtRow = INPUT_FIRST_DATA_ROW + i - 1
        If pics.Exists(srcRow) Then
            If i Mod 20 = 0 Then
                Application.StatusBar = "BOM import: copying picture " & i & " of " & nRows & " ..."
                DoEvents
            End If
            If CopyPictureToCell(pics(srcRow), inputWs, tgtRow, COL_PICTURE, knownNames) Then nPics = nPics + 1
        End If
    Next i

    ' safety net: fit everything once more, whatever happened during the pasting
    Application.StatusBar = "BOM import: fitting pictures into their cells ..."
    On Error Resume Next
    For Each shp In inputWs.Shapes
        tgtRow = PictureRowOf(shp)
        If tgtRow >= INPUT_FIRST_DATA_ROW Then FitShapeToCell shp, inputWs, tgtRow, COL_PICTURE
    Next shp

    inputWs.Range("A" & INPUT_FIRST_DATA_ROW).Select
    If Not prevActive Is Nothing Then prevActive.Activate
    Err.Clear
    On Error GoTo 0
End Sub


'==============================================================================
' Copies one thumbnail into the Picture column and fits it into the cell.
'
' knownNames holds the names of all shapes that were already on the sheet, so
' the pasted shape can be identified reliably (the paste itself lands wherever
' Excel wants - the fit afterwards puts it into its cell).
'==============================================================================
Private Function CopyPictureToCell(srcShape As Object, tgtWs As Worksheet, _
                                   ByVal r As Long, ByVal c As Long, _
                                   knownNames As Object) As Boolean
    Dim newShp As Object
    Dim i As Long, before As Long, tries As Long

    On Error Resume Next

    before = tgtWs.Shapes.Count

    ' paste next to the target cell - if anything below fails, the picture is at
    ' least in the right region of the sheet
    tgtWs.Cells(r, c).Select

    For tries = 1 To 3
        Err.Clear
        srcShape.Copy
        DoEvents
        tgtWs.Paste
        DoEvents
        If tgtWs.Shapes.Count > before Then Exit For
    Next tries
    Application.CutCopyMode = False
    If tgtWs.Shapes.Count <= before Then
        Err.Clear
        Exit Function                                   ' nothing was pasted
    End If

    ' the new shape is the one whose name was not there before
    For i = tgtWs.Shapes.Count To 1 Step -1
        If Not knownNames.Exists(tgtWs.Shapes(i).Name) Then
            Set newShp = tgtWs.Shapes(i)
            Exit For
        End If
    Next i
    If newShp Is Nothing Then Set newShp = tgtWs.Shapes(tgtWs.Shapes.Count)

    newShp.Name = PIC_TAG & r
    If Not knownNames.Exists(newShp.Name) Then knownNames.Add newShp.Name, True

    CopyPictureToCell = FitShapeToCell(newShp, tgtWs, r, c)
    Err.Clear
End Function


'==============================================================================
' Scales a picture into its cell (keeping the aspect ratio), centres it there
' and lets it move AND size with the cell afterwards.
'
' Every step has its own error trap, so one failing property can never stop the
' whole fit - that is what left the thumbnails at their paste position before.
'==============================================================================
Private Function FitShapeToCell(shp As Object, ws As Worksheet, _
                                ByVal r As Long, ByVal c As Long) As Boolean
    Dim cel As Range
    Dim w As Double, h As Double
    Dim maxW As Double, maxH As Double, f As Double

    On Error Resume Next

    Set cel = ws.Cells(r, c)
    If cel Is Nothing Then Exit Function

    If ws.Rows(r).RowHeight < MIN_PIC_ROW_HEIGHT Then ws.Rows(r).RowHeight = MIN_PIC_ROW_HEIGHT

    ' size the picture freely while it is being fitted
    shp.Placement = xlFreeFloating
    shp.LockAspectRatio = False

    w = 0: h = 0
    w = shp.Width
    h = shp.Height
    If w <= 0 Or h <= 0 Then
        Err.Clear
        Exit Function
    End If

    maxW = cel.Width - 2 * PIC_MARGIN
    maxH = cel.Height - 2 * PIC_MARGIN
    If maxW < 5 Then maxW = 5
    If maxH < 5 Then maxH = 5

    f = maxW / w
    If maxH / h < f Then f = maxH / h

    shp.Width = w * f
    shp.Height = h * f
    shp.Left = cel.Left + (cel.Width - shp.Width) / 2
    shp.Top = cel.Top + (cel.Height - shp.Height) / 2
    shp.Placement = xlMoveAndSize                  ' move and size with the cell

    FitShapeToCell = (Err.Number = 0)
    Err.Clear
End Function


'==============================================================================
' ENTRY POINT: re-fit all imported thumbnails into their Picture cells.
'
' Run this after changing row heights or column widths, or to repair pictures
' that ended up at the wrong place.
'==============================================================================
Public Sub FitPicturesToCells()
    Dim ws As Worksheet
    Dim shp As Object
    Dim wasProtected As Boolean
    Dim r As Long, n As Long

    If Not SheetExists(INPUT_SHEET) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(INPUT_SHEET)

    SaveAppState
    wasProtected = UnprotectSheet(ws)

    On Error Resume Next
    For Each shp In ws.Shapes
        r = PictureRowOf(shp)
        If r >= INPUT_FIRST_DATA_ROW Then
            If FitShapeToCell(shp, ws, r, COL_PICTURE) Then n = n + 1
        End If
    Next shp
    Err.Clear
    On Error GoTo 0

    If wasProtected Then ProtectSheet ws
    RestoreAppState

    MsgBox n & " picture(s) fitted into the Picture column.", vbInformation, "BOM Import"
End Sub


'==============================================================================
' Row a picture of the Input sheet belongs to, 0 if it is not a data picture.
' Imported pictures carry their row in the name (BOMPIC_<row>), which also works
' when the picture has been dragged somewhere else.
'==============================================================================
Private Function PictureRowOf(shp As Object) As Long
    Dim s As String, r As Long

    On Error Resume Next

    s = shp.Name
    If InStr(1, s, PIC_TAG, vbTextCompare) = 1 Then
        s = Mid$(s, Len(PIC_TAG) + 1)
        If IsNumeric(s) Then PictureRowOf = CLng(s)
        Err.Clear
        Exit Function
    End If

    If shp.Type = SHP_PICTURE Or shp.Type = SHP_LINKED_PICTURE Then
        r = 0
        r = shp.TopLeftCell.Row
        If r >= INPUT_FIRST_DATA_ROW Then PictureRowOf = r
    End If
    Err.Clear
End Function


'==============================================================================
' Builds a dictionary  source row -> picture shape  for the extraction sheet.
' The extractor anchors one thumbnail per row in the first column.
'==============================================================================
Private Function PictureMap(ws As Worksheet) As Object
    Dim d As Object, shp As Object, r As Long

    Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    For Each shp In ws.Shapes
        If shp.Type = SHP_PICTURE Or shp.Type = SHP_LINKED_PICTURE Then
            r = 0
            r = shp.TopLeftCell.Row
            If r > 0 Then
                If Not d.Exists(r) Then d.Add r, shp
            End If
        End If
    Next shp
    Err.Clear
    On Error GoTo 0
    Set PictureMap = d
End Function


'==============================================================================
' ENTRY POINT: remove everything a previous import wrote into "Input".
'==============================================================================
Public Sub ClearBOMImport()
    Dim ws As Worksheet
    Dim wasProtected As Boolean

    If Not SheetExists(INPUT_SHEET) Then Exit Sub
    If MsgBox("Clear the entry columns of the '" & INPUT_SHEET & "' sheet?" & vbCrLf & vbCrLf & _
              "A  Level" & vbTab & vbTab & "B  Lookup Key" & vbCrLf & _
              "C  Part Number" & vbTab & "D  Name" & vbCrLf & _
              "E  Other Reference" & vbTab & "F  Picture" & vbCrLf & _
              "G  Qty System" & vbTab & "P  Material Reference" & vbCrLf & _
              "Q  Description / Comments" & vbTab & "AC Process Reference" & vbCrLf & _
              "AD Process Description" & vbTab & "AH Process Time" & vbCrLf & _
              "AI Direct Operator" & vbTab & "AN Qty Operators for setup" & vbCrLf & _
              "AO Set Up Time per Lot Size" & vbTab & "AP Other Setup Cost" & vbCrLf & vbCrLf & _
              "... and all pictures of the data area." & vbCrLf & _
              "Formulas and all other columns stay untouched.", _
              vbQuestion + vbYesNo, "BOM Import") <> vbYes Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(INPUT_SHEET)
    SaveAppState
    wasProtected = UnprotectSheet(ws)
    ClearInputRange ws, INPUT_FIRST_DATA_ROW, LastTemplateRow(ws)
    If wasProtected Then ProtectSheet ws
    RestoreAppState
End Sub


'==============================================================================
' ENTRY POINT: put the two push buttons on the Input sheet.
'
'   [ Import BOM ]   -> ImportBOM
'   [ Clean Input ]  -> ClearBOMImport
'
' Run once; the buttons are saved with the workbook. Running it again simply
' replaces the two buttons (nothing is duplicated).
'==============================================================================
Public Sub CreateMacroButtons()
    Dim ws As Worksheet
    Dim anchor As Range
    Dim wasProtected As Boolean

    If Not SheetExists(INPUT_SHEET) Then
        MsgBox "This workbook has no sheet named '" & INPUT_SHEET & "'.", vbExclamation, "BOM Import"
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets(INPUT_SHEET)
    SaveAppState
    wasProtected = UnprotectSheet(ws)

    On Error Resume Next
    Set anchor = ws.Range(BUTTON_ANCHOR_CELL)
    If anchor Is Nothing Then Set anchor = ws.Range("F5")
    Err.Clear
    On Error GoTo 0

    DeleteShapeByName ws, BTN_IMPORT_NAME
    DeleteShapeByName ws, BTN_CLEAN_NAME

    AddMacroButton ws, BTN_IMPORT_NAME, "Import BOM", "ImportBOM", _
                   anchor.Left, anchor.Top, BUTTON_WIDTH, BUTTON_HEIGHT
    AddMacroButton ws, BTN_CLEAN_NAME, "Clean Input", "ClearBOMImport", _
                   anchor.Left + BUTTON_WIDTH + BUTTON_GAP, anchor.Top, _
                   BUTTON_WIDTH, BUTTON_HEIGHT

    If wasProtected Then ProtectSheet ws
    RestoreAppState

    MsgBox "The buttons  [ Import BOM ]  and  [ Clean Input ]  are on the '" & _
           INPUT_SHEET & "' sheet (at " & BUTTON_ANCHOR_CELL & ")." & vbCrLf & vbCrLf & _
           "Save the workbook to keep them." & vbCrLf & _
           "To move a button: right-click it and drag it to its place.", _
           vbInformation, "BOM Import"
End Sub


'==============================================================================
' Creates one form control button and links it to a macro.
'==============================================================================
Private Sub AddMacroButton(ws As Worksheet, ByVal shapeName As String, ByVal caption As String, _
                           ByVal macroName As String, ByVal posLeft As Double, ByVal posTop As Double, _
                           ByVal w As Double, ByVal h As Double)
    Dim btn As Object

    On Error Resume Next
    Set btn = ws.Buttons.Add(posLeft, posTop, w, h)
    If btn Is Nothing Then
        ' fall back to the shape based form control
        Set btn = ws.Shapes.AddFormControl(0, posLeft, posTop, w, h)   ' 0 = xlButtonControl
    End If
    If btn Is Nothing Then
        Err.Clear
        Exit Sub
    End If

    btn.Name = shapeName
    btn.OnAction = macroName
    btn.Placement = xlMove                     ' move with the cells, keep the size

    ' caption + font - whichever of the two flavours answers, the other call
    ' simply fails silently
    btn.Characters.Text = caption
    btn.Characters.Font.Bold = True
    btn.Characters.Font.Size = 10
    Err.Clear

    btn.TextFrame.Characters.Text = caption
    btn.TextFrame.Characters.Font.Bold = True
    btn.TextFrame.Characters.Font.Size = 10
    Err.Clear
End Sub


'==============================================================================
' Deletes a shape by name if it exists.
'==============================================================================
Private Sub DeleteShapeByName(ws As Worksheet, ByVal shapeName As String)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    Err.Clear
End Sub


'==============================================================================
' Clears the imported columns and the imported pictures of the given row range.
' Formulas and all other columns are untouched.
'==============================================================================
Private Sub ClearInputRange(ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim cols As Object, key As Variant, c As Long
    Dim colRange As Range, consts As Range

    If lastRow < firstRow Then Exit Sub

    Set cols = ColumnSet(ws, CLEAR_COLUMNS)

    For Each key In cols.Keys
        c = CLng(key)
        ' the level column is only cleared if the import writes it again
        If c <> COL_LEVEL Or WRITE_LEVEL Then
            Set colRange = ws.Range(ws.Cells(firstRow, c), ws.Cells(lastRow, c))
            Set consts = Nothing
            On Error Resume Next
            Set consts = colRange.SpecialCells(xlCellTypeConstants)
            Err.Clear
            On Error GoTo 0
            If Not consts Is Nothing Then consts.ClearContents
        End If
    Next key

    DeleteRowPictures ws, firstRow, lastRow
End Sub


'==============================================================================
' Deletes the pictures of the data area: everything tagged by this macro plus
' any picture sitting in the Picture column between firstRow and lastRow.
' The sheet logo and anything outside that area is kept.
'==============================================================================
Private Sub DeleteRowPictures(ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim shp As Object, i As Long
    Dim doomed As Collection
    Dim r As Long

    Set doomed = New Collection
    On Error Resume Next
    For Each shp In ws.Shapes
        r = 0
        r = shp.TopLeftCell.Row
        If InStr(1, shp.Name, PIC_TAG, vbTextCompare) = 1 Then
            doomed.Add shp                              ' imported, wherever it sits
        ElseIf (shp.Type = SHP_PICTURE Or shp.Type = SHP_LINKED_PICTURE) Then
            ' any picture inside the data rows - also one that was dropped at the
            ' wrong place; the sheet logo sits above the data area and is kept
            If r >= firstRow And r <= lastRow Then doomed.Add shp
        End If
    Next shp
    For i = 1 To doomed.Count
        doomed(i).Delete
    Next i
    Err.Clear
    On Error GoTo 0
End Sub


'==============================================================================
' ENTRY POINT: empty the whole "Input" sheet for a new estimate.
'
' Every entered value of the data area is deleted, but NO formula is touched -
' the cost model stays completely intact and simply calculates on empty inputs.
' Pictures of the data area are removed as well.
'
' The columns listed in RESET_KEEP_COLUMNS keep their template defaults
' (factory, process type, lot size, ...), otherwise the sheet would fill up
' with #DIV/0! and #N/A after the reset.
'==============================================================================
Public Sub ResetInputSheet()
    Dim ws As Worksheet
    Dim wasProtected As Boolean
    Dim lastRow As Long, nCleared As Long

    If Not SheetExists(INPUT_SHEET) Then
        MsgBox "This workbook has no sheet named '" & INPUT_SHEET & "'.", vbExclamation, "Reset Input"
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets(INPUT_SHEET)
    lastRow = LastTemplateRow(ws)

    If MsgBox("Delete ALL entered values of the '" & INPUT_SHEET & "' sheet " & _
              "(rows " & INPUT_FIRST_DATA_ROW & "-" & lastRow & ") and start a new estimate?" & vbCrLf & vbCrLf & _
              "Formulas are NOT touched - only typed-in values and pictures are removed." & vbCrLf & _
              "Defaults of the columns " & RESET_KEEP_COLUMNS & " are kept." & vbCrLf & vbCrLf & _
              "This cannot be undone.", _
              vbExclamation + vbYesNo + vbDefaultButton2, "Reset Input") <> vbYes Then Exit Sub

    SaveAppState
    On Error GoTo Failed

    wasProtected = UnprotectSheet(ws)
    If ws.ProtectContents Then
        RestoreAppState
        MsgBox "The '" & INPUT_SHEET & "' sheet is password protected and could not be unlocked.", _
               vbExclamation, "Reset Input"
        Exit Sub
    End If

    nCleared = ClearEnteredValues(ws, INPUT_FIRST_DATA_ROW, lastRow)
    If wasProtected Then ProtectSheet ws

    RestoreAppState
    Application.Calculate

    MsgBox "The '" & INPUT_SHEET & "' sheet is ready for a new estimate." & vbCrLf & vbCrLf & _
           "Cleared cells: " & nCleared & vbCrLf & _
           "Rows: " & INPUT_FIRST_DATA_ROW & "-" & lastRow & vbCrLf & _
           "All formulas are unchanged.", vbInformation, "Reset Input"
    Exit Sub

Failed:
    Dim msg As String
    msg = "Reset failed:" & vbCrLf & vbCrLf & "Error " & Err.Number & " - " & Err.Description
    RestoreAppState
    MsgBox msg, vbCritical, "Reset Input"
End Sub


'==============================================================================
' ENTRY POINT: create a NEW, empty input worksheet.
'
' The "Input" sheet is duplicated with all its formulas and formatting, the copy
' is emptied the same way as ResetInputSheet does, and the original sheet stays
' exactly as it is.
'
' Note: the evaluation sheets (Overview, System, CBS, ...) keep referring to the
' original "Input" sheet - the new sheet is a self contained working copy.
'==============================================================================
Public Sub NewInputSheet()
    Dim srcWs As Worksheet, newWs As Worksheet
    Dim lastRow As Long, nCleared As Long
    Dim newName As String

    If Not SheetExists(INPUT_SHEET) Then
        MsgBox "This workbook has no sheet named '" & INPUT_SHEET & "'.", vbExclamation, "New Input Sheet"
        Exit Sub
    End If

    Set srcWs = ThisWorkbook.Worksheets(INPUT_SHEET)
    newName = UniqueSheetName(INPUT_SHEET & " (new)")

    If MsgBox("Create the empty sheet '" & newName & "' as a copy of '" & INPUT_SHEET & "'?" & vbCrLf & vbCrLf & _
              "All formulas and the formatting are taken over, all entered values " & _
              "are removed. The existing '" & INPUT_SHEET & "' sheet is not changed." & vbCrLf & vbCrLf & _
              "Note: Overview / System / CBS keep calculating from the original '" & _
              INPUT_SHEET & "' sheet.", vbQuestion + vbYesNo, "New Input Sheet") <> vbYes Then Exit Sub

    SaveAppState
    On Error GoTo Failed

    srcWs.Copy After:=srcWs
    Set newWs = ThisWorkbook.Sheets(srcWs.Index + 1)
    On Error Resume Next
    newWs.Name = newName
    Err.Clear
    On Error GoTo Failed

    UnprotectSheet newWs
    lastRow = LastTemplateRow(newWs)
    nCleared = ClearEnteredValues(newWs, INPUT_FIRST_DATA_ROW, lastRow)

    RestoreAppState
    Application.Calculate
    newWs.Activate
    newWs.Range("A" & INPUT_FIRST_DATA_ROW).Select

    MsgBox "New input sheet created:  " & newWs.Name & vbCrLf & vbCrLf & _
           "Cleared cells: " & nCleared & vbCrLf & _
           "All formulas are unchanged.", vbInformation, "New Input Sheet"
    Exit Sub

Failed:
    Dim msg As String
    msg = "Could not create the new input sheet:" & vbCrLf & vbCrLf & _
          "Error " & Err.Number & " - " & Err.Description
    RestoreAppState
    MsgBox msg, vbCritical, "New Input Sheet"
End Sub


'==============================================================================
' Deletes every entered (non formula) value of the data area and returns how
' many cells were cleared. Cells containing a formula are never touched, the
' columns of RESET_KEEP_COLUMNS are skipped completely.
'==============================================================================
Private Function ClearEnteredValues(ws As Worksheet, ByVal firstRow As Long, _
                                    ByVal lastRow As Long) As Long
    Dim keep As Object
    Dim lastCol As Long, c As Long, n As Long
    Dim colRange As Range, consts As Range

    If lastRow < firstRow Then Exit Function

    Set keep = ColumnSet(ws, RESET_KEEP_COLUMNS)

    lastCol = ws.Cells(INPUT_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < COL_QTY Then lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1

    ' one column at a time - constants are cleared, formulas stay untouched
    For c = 1 To lastCol
        If Not keep.Exists(c) Then
            Set colRange = ws.Range(ws.Cells(firstRow, c), ws.Cells(lastRow, c))
            Set consts = Nothing
            On Error Resume Next
            Set consts = colRange.SpecialCells(xlCellTypeConstants)
            Err.Clear
            On Error GoTo 0
            If Not consts Is Nothing Then
                n = n + consts.Count
                consts.ClearContents
            End If
        End If
    Next c

    DeleteRowPictures ws, firstRow, lastRow
    ClearEnteredValues = n
End Function


'==============================================================================
' Turns a list of column letters ("A,B,AA") into a set of column numbers.
'==============================================================================
Private Function ColumnSet(ws As Worksheet, ByVal letters As String) As Object
    Dim d As Object
    Dim parts As Variant, i As Long, s As String, c As Long

    Set d = CreateObject("Scripting.Dictionary")
    parts = Split(letters, ",")

    For i = LBound(parts) To UBound(parts)
        s = Trim$(CStr(parts(i)))
        If Len(s) > 0 Then
            c = 0
            On Error Resume Next
            c = ws.Range(s & "1").Column
            Err.Clear
            On Error GoTo 0
            If c > 0 Then
                If Not d.Exists(c) Then d.Add c, True
            End If
        End If
    Next i

    Set ColumnSet = d
End Function


'==============================================================================
' Creates the separate sheet holding the complete extraction (same format).
' Returns the new sheet. keepIt=False means the sheet is only a temporary
' working copy and will be deleted by the caller.
'==============================================================================
Private Function CreateCopySheet(srcWs As Worksheet, ByVal keepIt As Boolean) As Worksheet
    Dim newWs As Worksheet
    Dim targetName As String
    Dim answer As VbMsgBoxResult
    Dim after As Worksheet

    If keepIt Then
        targetName = BOM_SHEET_NAME
        If SheetExists(targetName) Then
            answer = MsgBox("A sheet named '" & targetName & "' already exists." & vbCrLf & vbCrLf & _
                            "Yes  = replace it" & vbCrLf & _
                            "No   = keep it and create a new numbered sheet" & vbCrLf & _
                            "Cancel = abort", vbQuestion + vbYesNoCancel, "BOM Import")
            Select Case answer
                Case vbCancel: Exit Function
                Case vbYes: DeleteSheet ThisWorkbook.Worksheets(targetName)
                Case vbNo: targetName = UniqueSheetName(BOM_SHEET_NAME)
            End Select
        End If
    Else
        targetName = UniqueSheetName("~BOM_tmp")
    End If

    If SheetExists(INPUT_SHEET) Then
        Set after = ThisWorkbook.Worksheets(INPUT_SHEET)
    Else
        Set after = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    End If

    srcWs.Copy After:=after
    ' the copy is always placed directly behind the reference sheet
    Set newWs = ThisWorkbook.Sheets(after.Index + 1)

    On Error Resume Next
    newWs.Name = targetName
    Err.Clear
    On Error GoTo 0

    FlattenFormulas newWs
    Set CreateCopySheet = newWs
End Function


'==============================================================================
' The extraction is static data. Any formula in the copy would now point back at
' the source workbook, so every formula is replaced by its result.
'==============================================================================
Private Sub FlattenFormulas(ws As Worksheet)
    Dim fRange As Range, area As Range

    On Error Resume Next
    Set fRange = ws.Cells.SpecialCells(xlCellTypeFormulas)
    Err.Clear
    On Error GoTo 0
    If fRange Is Nothing Then Exit Sub

    On Error Resume Next
    For Each area In fRange.Areas          ' a multi area range cannot be set at once
        area.Value = area.Value
    Next area
    Err.Clear
    On Error GoTo 0
End Sub


'==============================================================================
' Copies the last prepared template row down so that more BOM rows fit.
'==============================================================================
Private Function ExtendTemplate(ws As Worksheet, ByVal preparedLastRow As Long, _
                                ByVal neededLastRow As Long) As Boolean
    On Error GoTo Failed
    If neededLastRow <= preparedLastRow Then
        ExtendTemplate = True
        Exit Function
    End If

    ws.Rows(preparedLastRow).Copy _
        Destination:=ws.Rows(preparedLastRow + 1 & ":" & neededLastRow)
    Application.CutCopyMode = False

    ' the copied rows must not carry the values of the source row
    ClearInputRange ws, preparedLastRow + 1, neededLastRow

    ExtendTemplate = True
    Exit Function

Failed:
    Application.CutCopyMode = False
    ExtendTemplate = False
    Err.Clear
End Function


'==============================================================================
' HELPERS
'==============================================================================

'--- file / workbook ----------------------------------------------------------
Private Function PickBOMFile() As String
    Dim f As Variant
    Dim startDir As String

    startDir = ThisWorkbook.Path
    On Error Resume Next
    If Len(startDir) > 0 Then ChDir startDir
    On Error GoTo 0

    f = Application.GetOpenFilename( _
            FileFilter:="BOM extraction (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", _
            Title:="Select the CATIA BOM extraction file")
    If VarType(f) = vbBoolean Then Exit Function      ' cancelled
    PickBOMFile = CStr(f)
End Function

Private Function GetWorkbook(ByVal fullPath As String, ByRef alreadyOpen As Boolean) As Workbook
    Dim wb As Workbook

    alreadyOpen = False
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            alreadyOpen = True
            Set GetWorkbook = wb
            Exit Function
        End If
    Next wb

    On Error Resume Next
    Set GetWorkbook = Application.Workbooks.Open(Filename:=fullPath, _
                                                 UpdateLinks:=0, ReadOnly:=True)
    Err.Clear
    On Error GoTo 0
End Function

'--- extraction layout --------------------------------------------------------
Private Function FindBOMSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If FindHeaderRow(ws) > 0 Then
            Set FindBOMSheet = ws
            Exit Function
        End If
    Next ws
End Function

' The header row is the first row (within the first 20) that carries a
' "Part Number" column header.
Private Function FindHeaderRow(ws As Worksheet) As Long
    Dim r As Long, c As Long, txt As String
    Dim maxC As Long

    maxC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If maxC < 5 Then maxC = 30

    For r = 1 To 20
        For c = 1 To maxC
            txt = LCase$(Trim$(CStr(ws.Cells(r, c).Value)))
            txt = Replace(txt, "_", " ")
            If txt = "part number" Or txt = "partnumber" Or txt = "part no" Or txt = "part-number" Then
                FindHeaderRow = r
                Exit Function
            End If
        Next c
    Next r
End Function

Private Function FindColumn(ws As Worksheet, ByVal hdrRow As Long, ParamArray names() As Variant) As Long
    Dim c As Long, i As Long, txt As String, maxC As Long

    If hdrRow <= 0 Then Exit Function
    maxC = ws.Cells(hdrRow, ws.Columns.Count).End(xlToLeft).Column
    If maxC < 1 Then maxC = 30

    For c = 1 To maxC
        txt = LCase$(Trim$(CStr(ws.Cells(hdrRow, c).Value)))
        txt = Replace(Replace(txt, vbLf, " "), "_", " ")
        Do While InStr(txt, "  ") > 0
            txt = Replace(txt, "  ", " ")
        Loop
        For i = LBound(names) To UBound(names)
            If txt = LCase$(CStr(names(i))) Then
                FindColumn = c
                Exit Function
            End If
        Next i
    Next c
End Function

Private Function LastDataRow(ws As Worksheet, ByVal hdrRow As Long) As Long
    Dim cPart As Long, r As Long

    cPart = FindColumn(ws, hdrRow, "part number", "partnumber", "part no", "part-number")
    If cPart = 0 Then Exit Function

    r = ws.Cells(ws.Rows.Count, cPart).End(xlUp).Row
    If r < hdrRow Then r = hdrRow
    LastDataRow = r
End Function

' Last row of the Input sheet that is prepared with the model formulas.
Private Function LastTemplateRow(ws As Worksheet) As Long
    Dim r As Long

    r = ws.Cells(ws.Rows.Count, TEMPLATE_PROBE_COL).End(xlUp).Row
    If r < INPUT_FIRST_DATA_ROW Then r = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If r < INPUT_FIRST_DATA_ROW Then r = INPUT_FIRST_DATA_ROW
    LastTemplateRow = r
End Function

' Keeps the rows listed in idx() of a column that was read for ALL source rows.
' That is what keeps values, pictures and the target rows aligned when rows are
' filtered out or the import is truncated.
'==============================================================================
Private Function PickRows(all As Variant, idx() As Long, ByVal n As Long) As Variant
    Dim out() As Variant, j As Long

    ReDim out(1 To n, 1 To 1)
    For j = 1 To n
        out(j, 1) = all(idx(j), 1)
    Next j
    PickRows = out
End Function


'==============================================================================
' Reads one column into a 2D array; text stays text, numeric text becomes a number.
Private Function ReadColumn(ws As Worksheet, ByVal col As Long, ByVal firstRow As Long, _
                            ByVal lastRow As Long, ByVal asNumber As Boolean) As Variant
    Dim raw As Variant, out() As Variant
    Dim n As Long, i As Long, v As Variant

    n = lastRow - firstRow + 1
    ReDim out(1 To n, 1 To 1)

    If n = 1 Then
        ReDim raw(1 To 1, 1 To 1)
        raw(1, 1) = ws.Cells(firstRow, col).Value
    Else
        raw = ws.Range(ws.Cells(firstRow, col), ws.Cells(lastRow, col)).Value
    End If

    For i = 1 To n
        v = raw(i, 1)
        If IsEmpty(v) Then
            out(i, 1) = vbNullString
        ElseIf asNumber Then
            If IsNumeric(v) Then
                out(i, 1) = CDbl(v)
            Else
                out(i, 1) = v
            End If
        ElseIf VarType(v) = vbString Then
            ' body rows of the extractor are written indented ("    PartBody"),
            ' the leading blanks must not end up in the part number
            out(i, 1) = Trim$(v)
        Else
            out(i, 1) = v
        End If
    Next i

    ReadColumn = out
End Function

'--- sheets -------------------------------------------------------------------
Private Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function

Private Function UniqueSheetName(ByVal baseName As String) As String
    Dim i As Long, nm As String

    nm = baseName
    i = 1
    Do While SheetExists(nm)
        i = i + 1
        nm = baseName & " " & i
    Loop
    UniqueSheetName = nm
End Function

Private Sub DeleteSheet(ws As Worksheet)
    Dim prev As Boolean
    If ws Is Nothing Then Exit Sub
    prev = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    ws.Delete
    Err.Clear
    On Error GoTo 0
    Application.DisplayAlerts = prev
End Sub

'--- protection ---------------------------------------------------------------
Private Function UnprotectSheet(ws As Worksheet) As Boolean
    If Not ws.ProtectContents Then Exit Function
    On Error Resume Next
    ws.Unprotect
    Err.Clear
    On Error GoTo 0
    UnprotectSheet = Not ws.ProtectContents
End Function

Private Sub ProtectSheet(ws As Worksheet)
    On Error Resume Next
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, _
               AllowFormattingCells:=True, AllowFormattingRows:=True
    Err.Clear
    On Error GoTo 0
End Sub

'--- application state --------------------------------------------------------
Private Sub SaveAppState()
    If mStateSaved Then Exit Sub
    mScreen = Application.ScreenUpdating
    mEvents = Application.EnableEvents
    mCalc = Application.Calculation
    mAlerts = Application.DisplayAlerts
    mStateSaved = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub RestoreAppState()
    If Not mStateSaved Then Exit Sub
    On Error Resume Next
    Application.CutCopyMode = False
    Application.Calculation = mCalc
    Application.DisplayAlerts = mAlerts
    Application.EnableEvents = mEvents
    Application.ScreenUpdating = mScreen
    Application.StatusBar = False
    Err.Clear
    On Error GoTo 0
    mStateSaved = False
End Sub

'--- reporting ----------------------------------------------------------------
Private Function BuildReport(ByVal madeCopy As Boolean, ByVal filledInput As Boolean, _
                             ByVal copySheetName As String, ByVal keptCopySheet As Boolean, _
                             ByVal nRows As Long, ByVal nPics As Long, _
                             ByVal nSkipped As Long, ByVal secs As Double) As String
    Dim s As String

    s = "BOM import finished." & vbCrLf & vbCrLf
    If madeCopy And keptCopySheet Then
        s = s & "Extraction copied to sheet:  " & copySheetName & vbCrLf
    End If
    If filledInput Then
        s = s & "Rows written to '" & INPUT_SHEET & "':  " & nRows & _
                "  (rows " & INPUT_FIRST_DATA_ROW & "-" & INPUT_FIRST_DATA_ROW + nRows - 1 & ")" & vbCrLf
        s = s & "Pictures copied:  " & nPics & vbCrLf
        If mFilteredOut > 0 Then
            s = s & "Sub-assembly rows skipped:  " & mFilteredOut & vbCrLf
        End If
        If nSkipped > 0 Then
            s = s & vbCrLf & "NOT imported: " & nSkipped & " row(s) - the Input sheet " & _
                    "had no prepared rows left." & vbCrLf
        End If
    End If
    s = s & vbCrLf & "Duration: " & Format$(secs, "0.0") & " s"

    BuildReport = s
End Function
