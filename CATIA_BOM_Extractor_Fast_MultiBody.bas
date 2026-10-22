Option Explicit

' ==============================================================================
' MACRO: MASTER BOM v5 (SPLIT BY LEVEL + NUMBERS FIRST, PICTURES LAST,
'                       NO LONG FREEZES, OWN EXCEL, AUTO-SAVE AND LOG)
'
' HOW A RUN GOES:
'   1. You are asked at which level the BOM is split. The question shows how
'      many components and sub-assemblies every level of the product has.
'        0 = no split: one sheet for the whole product
'        1 = one worksheet per first level node, named as in the CATIA tree
'            (e.g. "_AAA_LASER IGNITER GENV TCA"); the parts directly below the
'            product get one more sheet
'        n = one worksheet per sub-assembly on level n. While one is measured
'            and pictured, the others are hidden, so CATIA only draws that one.
'            Parts that are not inside any of them get one extra sheet.
'      The "Index" sheet links to every sheet and shows its level, quantity,
'      path, number of rows, status and time.
'   2. You are asked whether pictures are wanted. "No" is the fastest run.
'   3. Every sheet is filled in TWO PASSES:
'        pass 1 - all numbers: the display is frozen and nothing is shown or
'                 hidden, CATIA only loads and measures the PARTS. The assembly
'                 rows are added up from their parts afterwards. The sheet is
'                 complete and visible in Excel before the first picture.
'        pass 2 - all pictures: every part (and every sub-assembly that is a
'                 CATProduct of its own) is opened ALONE in a window of its
'                 own, pictured in the iso view and closed again - nothing in
'                 the product has to be hidden for that. A part number is
'                 pictured once for the whole run. Every picture goes into
'                 Excel straight away.
'   4. A component that fails is marked on the Index sheet and the run goes on
'      with the next one.
'
' WHERE THE RESULT IS:
'   - a NEW Excel window of its own (your own Excel stays free while it runs)
'   - saved automatically as  C:\Temp\catia_bom\BOM_<product>_<date>_<time>.xlsx
'     during the run and at the end - also when the run stops with an error
'   - the log of the run next to it (.txt, same name): one line per row and per
'     picture, so if CATIA ever hangs, the LAST line says on what. The "Log" sheet
'     lists every step that took longer than SLOW_STEP_SECONDS.
'
' WHY THIS VERSION DOES NOT FREEZE FOR HOURS:
'   - Product.ApplyWorkMode is RECURSIVE. The old version called it on whole
'     sub-assemblies (the component of a sheet, every assembly row), which loads
'     every CATPart below in ONE call: CATIA "not responding" for as long as that
'     takes. Now only PARTS are promoted, one per row, with the message queue
'     pumped in between.
'   - Mass / volume / area of an assembly row are the SUM of its parts. The old
'     version let CATIA integrate the whole subtree again (Analyze.Volume and
'     WetArea over hundreds of parts) - in ONE call per assembly row.
'   - Nothing is loaded while the levels are counted, and no file alert can wait
'     hidden behind the windows (DisplayFileAlerts is off from the start).
'   - A sub-assembly that is NOT LOADED in CATIA is not loaded by the macro
'     either (that would again be one huge blocking call): it gets a row marked
'     "NOT LOADED" and the final message lists it - load it and run again if
'     its content is needed.
'   - The tree walks pump the message queue every few nodes.
'   - Error 50290 means "Excel is busy" (a cell in edit mode, an open dialog).
'     The macro writes into an Excel of its own, waits while it is busy and
'     never ends the run because of one write that failed.
'
' WHY IT IS FAST AND QUIET:
'   - NO temporary geometry is created in the CATParts (MEASURE_MODE). The
'     dimensions come from the inertia, which only reads.
'   - The specification tree is switched off through the WINDOW LAYOUT, which
'     works in every language and immediately.
'   - Every wait pumps the message queue (Settle) instead of blocking the thread.
'   - HSOSynchronized is off for the run, so Selection.Add does not repaint the
'     tree every time.
'   - The rows of a sheet are collected in memory and go into Excel in a few
'     big blocks instead of one call per cell.
'   - Mass, material, dimensions and body names are read ONCE per part number
'     for the whole run, over all sheets.
'
' IF IT IS STILL TOO SLOW: split at a deeper level first, or answer "No" to
' the pictures. The settings block below has more switches:
'   CAPTURE_ASSEMBLY_SHOTS = False  no pictures on the assembly rows
'   CAPTURE_PART_SHOTS = False      no pictures at all
'   ASSY_DIMS_MAX_PARTS = 0         no length/width/height on assembly rows
'   DETECT_MULTIBODY = False        no extra row per body of a multi-body part
'
' TWO VERSIONS - PICK THE ONE YOU NEED (Tools > Macro > Macros):
'
'   GenerateMasterBOM      PART LIST (the classic BOM, unchanged behaviour)
'                          - one row per part number
'                          - Qty = number of instances in the whole product
'                          - no sub-assembly rows, so the list stays short
'
'   GenerateAssemblyBOM    ASSEMBLY STRUCTURE
'                          - sub-assemblies get their own row, marked
'                            "Assembly" in column O and written in bold,
'                            directly above their content
'                          - every assembly lists its own content, so a part
'                            that sits in three assemblies appears below each
'                            of them and can be located in the tree
'                          - Qty is the number of that part below that path,
'                            the sum over the rows stays the total quantity
'
' BOTH versions write:
'   - Column N "First Level": the first level node (direct child of the root)
'     the row belongs to, named as in the CATIA tree ("<name>-<part number>"),
'     so the BOM can be filtered per first level
'   - Column O "Type": Assembly / Part / Body, or "Not loaded" for a component
'     that could not be loaded (no values, no picture)
'   - Column P "CATIA tree": every row named and indented as in the CATIA
'     tree, with its instance - and the rows are grouped like the tree: the
'     +/- at the left of Excel fold and unfold every assembly
'   - Columns A..M unchanged, so both sheets import into the ACE cost model
'
' MULTI-BODY CATPart:
'   - If a leaf CATPart contains MORE THAN ONE solid Body (PartBody), each
'     body is treated as an extra sub-part of the product tree:
'       * one extra row per body, Level = part level + 1
'       * all bodies are hidden, only the measured body is shown
'       * screenshot, mass, volume, area, bounding box dims and material
'         are reported for that single body
'   - Bodies with no geometry (empty bodies) are ignored
' ==============================================================================

' Only used for the short naps inside Settle(); Private so that the module can
' live next to the other extractors without an "ambiguous name" error.
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' ==============================================================================
' SETTINGS
' ==============================================================================
' --- run mode --------------------------------------------------------------
' Both entry points set these two before the run; they are never changed by
' hand. mGroupBy can be:
'   "GLOBAL" - one row per part number in the whole tree (part list); column N
'              then lists all first level branches the part is used in
'   "PARENT" - one row per part per parent assembly (assembly structure)
'   "FIRST"  - one row per part per first level branch (in between)
Private mIncludeSubassemblies As Boolean
Private mGroupBy As String

' --- NAMES AS IN THE CATIA TREE ----------------------------------------------
' The Teamcenter tree shows a component as
'   "<Definition>_<Revision>_<Description>-<part number> (<instance>)",
' e.g. "MTH1111000X0003-01_AAA_MK2-BL3-TC_INTEGRATED-1DP000722498 (START_SKELETON)"
' or "__-CH4-Schrauben (Product1.4)" when the three fields are empty, and
' "[?]" for a field CATIA cannot read. The column "CATIA tree" and the column
' "First Level" use this format. Fields: {DEFINITION} {REVISION} {DESCRIPTION}
' {NOMENCLATURE} {PN}. The log of the run lists these properties of every first
' level node - if the names do not match the tree, the format is changed here.
Private Const TREE_NAME_FORMAT As String = "{DEFINITION}_{REVISION}_{DESCRIPTION}-{PN}"

' --- HOW LENGTH / WIDTH / HEIGHT ARE MEASURED -------------------------------
'   "SAFE" (default) - the equivalent box from the inertia. NOTHING is created
'                      inside the CATParts: no bounding box feature, no update
'                      engine, no delete command. That is what stops CATIA from
'                      popping an error panel for every single component, keeps
'                      the parts unmodified and makes the run much faster.
'                      The values are an inertia equivalent box, not the exact
'                      axis aligned box - good enough for a cost BOM.
'   "BOX"            - the exact axis aligned bounding box. Creates temporary
'                      geometry inside every CATPart and runs CATIA's update
'                      engine on it. Slow, marks the parts as modified and CATIA
'                      shows its own modal error panel for every part that
'                      cannot be updated or is read only - VBA cannot suppress
'                      those. The whole path switches itself off after the first
'                      failure, so you get at most one dialog per run.
Private Const MEASURE_MODE As String = "SAFE"

' --- SCREENSHOTS ------------------------------------------------------------
Private Const CAPTURE_PART_SHOTS As Boolean = True
' Assembly thumbnails have to show and hide every leaf below the node, so a
' node with more leaves than the limit is listed without a picture instead of
' stalling the run. Set CAPTURE_ASSEMBLY_SHOTS = False for the fastest run.
Private Const CAPTURE_ASSEMBLY_SHOTS As Boolean = True
Private Const MAX_ASSY_LEAVES_FOR_SHOT As Long = 400
' False (default) = every picture is taken in the MAIN window: every part is
' hidden, the one of the row is shown, captured and hidden again. A part
' number is pictured only once for the whole run.
' True = every part (and every sub-assembly that is a CATProduct of its own)
' is first tried in a window of its own. On some installations (seen with
' Teamcenter) such a window is never displayed and CATIA captures the main
' window instead - every picture the same. Those pictures are recognised and
' taken in the main window after all.
Private Const PICTURES_IN_OWN_WINDOW As Boolean = False
' True = after the split level the macro asks whether pictures are wanted at
' all. A run without pictures is the quickest way to all numbers of a big product.
Private Const ASK_PICTURES As Boolean = True
' One fixed isometric viewpoint for all pictures instead of whatever camera the
' session happened to be left in.
Private Const SET_ISO_VIEW As Boolean = True
' The specification tree is switched off through the WINDOW LAYOUT, which is
' language independent. The compass has no automation property at all, so it can
' only be toggled by its command name - and that name is translated. Put the
' name of your installation here, or "" to leave the compass alone:
'   English "Compass"   German "Kompass"   French "Boussole"
Private Const COMPASS_COMMAND As String = "Compass"

' --- SETTLE TIMES (message pumping, never a blocking Sleep) -----------------
Private Const PART_SETTLE_MS As Long = 30
Private Const ASSY_SETTLE_MS As Long = 120

' --- SPEED ------------------------------------------------------------------
' Multi-body detection has to load every leaf CATPart. Switch it off when you do
' not need one row per body.
Private Const DETECT_MULTIBODY As Boolean = True
' Product.ApplyWorkMode is RECURSIVE: one call on the root loads every CATPart
' of the product in a single blocking COM call - hours of "CATIA not responding"
' on a big product. Parts are promoted one at a time in the row loop instead.
' Set this to True only if your parts must all be loaded first.
Private Const FORCE_DESIGN_MODE As Boolean = False
' A sub-assembly that is NOT LOADED (unloaded in the tree) looks like an empty
' node. True = load its structure while the levels are counted. That can load a
' lot from Teamcenter before the first question even appears, so it is off.
Private Const EXPAND_UNLOADED As Boolean = False

' A component in NoShow hides its WHOLE subtree in the 3D view, whatever the
' children say. So a leaf below a sub-assembly that the user had hidden would be
' captured as an empty picture. The run therefore shows every node once and then
' hides only the leaves. Side effect: components that were hidden before the run
' are visible after it. Set PRESERVE_VISIBILITY = True to have the macro record
' the previous state of every node and put it back - correct, but it costs one
' selection round trip per node before the first row.
Private Const PRESERVE_VISIBILITY As Boolean = False

' --- ASSEMBLY ROWS ----------------------------------------------------------
' Mass / volume / area of a sub-assembly row are the SUM of the part rows below
' it (quantities included). Nothing is loaded or integrated for that, so it is
' cheap. Parts without values count as 0.
Private Const ASSEMBLY_METRICS As Boolean = True
' Length / width / height of an assembly row come from CATIA's inertia of the
' whole assembly, which IS computed over everything below it. Only assemblies
' with at most this many part instances get them; bigger ones show "N/A".
' 0 = never.
Private Const ASSY_DIMS_MAX_PARTS As Long = 200

' --- SPLITTING THE BOM BY LEVEL ----------------------------------------------
' At the start the macro asks at which level the BOM is split, and shows how
' many components and sub-assemblies every level has:
'   0 = no split - one sheet for the whole product
'   n = one worksheet per sub-assembly on level n. Each of them is opened in a
'       window of its own, so CATIA only has to draw that sub-assembly while it
'       is measured and pictured. Parts that are not inside any of them get one
'       extra sheet, and an "Index" sheet links to every component sheet.
' A sub-assembly on level n that is used several times gets ONE sheet; its
' quantities are per piece and the Index says how many pieces there are.
' False (default) = no question: the whole product on ONE sheet, in the order
' of the CATIA tree, with the tree in column P and +/- groups like the tree.
Private Const ASK_SPLIT_LEVEL As Boolean = False
' False (default) = every component is done in the main window: the other
' components are hidden meanwhile, so CATIA still only draws the one that is
' worked on. True = each component in a window of its own - but on some
' installations (seen with Teamcenter) the pictures of such a window come out
' as pictures of the main window, so it is off.
' A component that is not the root of its own CATProduct file (e.g. a
' "Product1" inserted inside another assembly) is always done in the main window.
Private Const OPEN_UNITS_IN_NEW_WINDOW As Boolean = False

' --- EXCEL, SAVING AND LOG ---------------------------------------------------
' True = the BOM goes into an Excel of its OWN (a separate Excel window). Your
' own Excel stays free while the macro runs, and nothing you do there can make
' a write fail. False = the Excel that is already open is used, as before.
Private Const OWN_EXCEL_INSTANCE As Boolean = True
' When the macro's Excel is busy anyway (somebody left a cell in edit mode or a
' dialog open there), it waits this long before it skips that one write.
Private Const EXCEL_WAIT_SECONDS As Long = 120
' The workbook is saved here during the run (after a component sheet, at most
' every SAVE_EVERY_SECONDS) and at the end - also when the run stops with an
' error. "" = no automatic saving. The log of the run is written next to it.
Private Const SAVE_FOLDER As String = "C:\Temp\catia_bom"
Private Const SAVE_EVERY_SECONDS As Long = 300
' Steps that take at least this long are listed on the "Log" sheet
Private Const SLOW_STEP_SECONDS As Double = 5

' --- caches (whole run, keyed by part number) --------------------------------
' A part that appears on several sheets is read only once.
Private mDimCache As Object      ' part number -> Array(L, W, H)
Private mMatCache As Object      ' part number -> Array(material, density)
Private mBodyCache As Object     ' part number -> body names, or Empty for one body
Private mPropCache As Object     ' part number -> Array(mass, volume, area)
Private mAssyKids As Object      ' assembly part number -> Collection of Array(child part number, is assembly)
Private mRollCache As Object     ' assembly part number -> Array(mass, volume, area, parts, parts without values)
Private mBBoxDisabled As Boolean ' the temp geometry path failed once - stop it

' --- the sheet being written (reset for every sheet) --------------------------
Private mQty As Object, mRef As Object, mDesc As Object, mProps As Object
Private mLevel As Object, mBodies As Object, mPartNum As Object, mFirst As Object
Private mIsAssy As Object, mNodeLeaves As Object
Private mLeaves As Collection    ' every leaf instance of the sheet
Private mAllNodes As Collection  ' every instance of the sheet: assemblies AND leaves
Private mAssyNodes As Collection ' only the assembly instances of the sheet
Private mJobRow As Collection    ' pictures to take in pass 2: row ...
Private mJobKind As Collection   ' ... "PART" / "ASSY" / "BODY" / "UNIT"
Private mJobKey As Collection    ' ... row key
Private mJobBody As Collection   ' ... body name for "BODY"
Private mBuf() As Variant        ' the rows of the sheet: mBuf(column, sheet row - 1)
Private mBufRows As Long         ' rows used in mBuf
Private mFlushed As Boolean      ' mBuf is already in Excel
Private mLostBlocks As Object    ' first buffer row of every block Excel did not take
Private mBoldRows As Collection  ' sheet rows written in bold (assemblies)
Private mAssyRows As Collection  ' Array(sheet row, part number, product) of every assembly row
Private mSheetName As String     ' name of the sheet being written

' --- the whole run -------------------------------------------------------------
Private mXlApp As Object, mXlBook As Object, mFso As Object
Private mOwnXl As Boolean        ' mXlApp was started by this run
Private mXlLost As Boolean       ' Excel was closed or crashed - nothing can be written any more
Private mXlSkipAt As Double      ' when the user last chose to skip a write to a busy Excel (-1 = never)
Private mXlLastErr As Long       ' error of the last failed XlWriteBlock (0 = skipped / busy)
Private mPromoted As Object      ' part numbers already promoted to design mode in this run
Private mBodyValCache As Object  ' part number|body -> Array(mass, vol, area, L, W, H, material, density)
Private mShotFiles As Object     ' "P|pn" / "A|pn" / "B|pn|body" -> picture file, reused by every row
Private mShotSigs As Object      ' same keys -> fingerprint of that picture file
Private mShotNo As Long          ' running number of the picture files
Private mPartDocs As Object      ' part number -> the CATPart document pass 1 found for it
Private mTempDir As String
Private mSheetNo As Long
Private mSheetNames As Object    ' every sheet name used in the workbook
Private mShotParts As Boolean    ' pictures of the part (and body) rows in this run
Private mShotAssy As Boolean     ' pictures of the assembly rows in this run
Private mLogPath As String, mSavePath As String
Private mSavedOnce As Boolean    ' the workbook has been saved at least once
Private mSaveTried As Boolean    ' a save has been attempted (throttles retries of a failing save)
Private mLastSave As Double      ' Timer of the last save attempt
Private mLastSaveOK As Boolean   ' the last save attempt worked
Private mLastSaveAt As Date      ' time of the last save that worked
Private mIdxRedo As Collection   ' Index writes Excel could not take - repeated at the end
Private mIdxFmtOK As Boolean     ' the text formats of the Index are applied
Private mIdxEarly As Collection  ' Index rows written BEFORE those formats - written again after them
Private mNotLoaded As Object     ' part numbers of sub-assemblies that are not loaded in CATIA
Private mNoPartNumber As Object  ' row names of components whose part number cannot be read
Private mProvisional As Object   ' row names given before the part number could be read
Private mSlow As Collection      ' Array(seconds, step) of every slow step
Private mPump As Long            ' node counter of the tree walks


' ==============================================================================
' ENTRY POINT 1: PART LIST  (the classic BOM)
' One row per part number, Qty = instances, no assembly rows.
' ==============================================================================
Sub GenerateMasterBOM()
    mIncludeSubassemblies = False
    mGroupBy = "GLOBAL"
    Call RunBOM
End Sub

' ==============================================================================
' ENTRY POINT 2: ASSEMBLY STRUCTURE
' Sub-assemblies get their own row and every assembly lists its own content.
' ==============================================================================
Sub GenerateAssemblyBOM()
    mIncludeSubassemblies = True
    mGroupBy = "PARENT"
    Call RunBOM
End Sub


' ==============================================================================
' THE WORKER USED BY BOTH ENTRY POINTS
'
'   1. asks at which level the BOM is split (0 = one sheet, as before) and
'      whether pictures are wanted
'   2. split level 0:  one sheet for the whole product
'      split level n:  one sheet per sub-assembly on level n, each worked in its
'                      own CATIA window; parts above level n on an extra sheet;
'                      an "Index" sheet with a link to every component sheet
'   3. every sheet is filled in two passes: first ALL numbers (display frozen,
'      no show/hide), then all pictures
' ==============================================================================
Private Sub RunBOM()

    Dim catDoc As Document
    Dim rootProd As Product
    Dim rootSel As Selection
    Dim rootWin As Object
    Dim rootViewer As Viewer
    Dim rootLayout As Long, rootLayoutSet As Boolean
    Dim compassToggled As Boolean

    Dim splitLevel As Long, ans As Long
    Dim rootPN As String, modeName As String
    Dim xlSheet As Object, idxSheet As Object, firstSheet As Object
    Dim startTime As Double, unitStart As Double
    Dim nRows As Long, nAssy As Long, nParts As Long
    Dim totRows As Long, totAssy As Long, totParts As Long
    Dim nUnits As Long, nUnitsOK As Long
    Dim statusTxt As String, unitOK As Boolean, failMsg As String

    ' components on the split level, keyed by part number
    Dim dUnitQty As Object, dUnitRef As Object, dUnitFirst As Object
    Dim dUnitDesc As Object, dUnitPath As Object
    Dim colUnitInst As Collection      ' every instance on the split level
    Dim colAbove As Collection         ' every assembly instance above it
    Dim colLooseInst As Collection     ' rows of the extra sheet: parts in no component
                                       ' (+ the assemblies above, in the structure mode)
    Dim colLooseLevel As Collection, colLooseFirst As Collection
    Dim colLooseKey As Collection, colLooseIsAssy As Collection
    Dim colLooseParts As Collection    ' only the PARTS of it - they are what gets hidden

    ' components that cannot get their own window (they are defined inside
    ' another CATProduct file) are done afterwards in the main window
    Dim colFbKey As Collection, colFbSheet As Collection, colFbIdxRow As Collection

    Dim unitKey As Variant, idxRow As Long, i As Long
    Dim unitInst As Product, unitRoot As Product
    Dim unitDoc As Object, unitWin As Object
    Dim unitViewer As Viewer, unitSel As Selection
    Dim unitLayout As Long, unitLayoutSet As Boolean
    Dim ownWindow As Boolean
    Dim rootSceneSet As Boolean
    Dim idxHdrOK As Boolean
    Dim unitName As String, looseName As String, looseText As String
    Dim dLabelCount As Object, lbl As String

    startTime = Timer

    ' nothing of an earlier run may be used by this one - not even by the
    ' error handler (log file, workbook, save path)
    Set mXlApp = Nothing
    Set mXlBook = Nothing
    mOwnXl = False
    mXlLost = False
    mXlSkipAt = -1
    mLogPath = ""
    mSavePath = ""
    mSavedOnce = False
    mSaveTried = False
    mLastSaveOK = False
    Set mIdxRedo = New Collection
    mIdxFmtOK = False
    Set mIdxEarly = New Collection

    On Error GoTo Fail

    ' --- 1. DOCUMENT ------------------------------------------------------------
    On Error Resume Next
    Set catDoc = CATIA.ActiveDocument
    If Err.Number <> 0 Or catDoc Is Nothing Then
        On Error GoTo 0
        MsgBox "No document open.", vbCritical
        Exit Sub
    End If
    On Error GoTo Fail

    If InStr(1, catDoc.Name, ".CATProduct", vbTextCompare) = 0 Then
        MsgBox "Please open an assembly (CATProduct) first.", vbExclamation
        Exit Sub
    End If

    Set rootProd = catDoc.Product
    Set rootSel = catDoc.Selection
    Set rootWin = CATIA.ActiveWindow

    ' a file alert (missing / locked document) is a MODAL panel that can sit
    ' hidden behind the windows and stop everything - none from here on
    rootPN = ""
    On Error Resume Next
    CATIA.DisplayFileAlerts = False
    rootPN = rootProd.PartNumber
    Err.Clear
    On Error GoTo Fail
    If Len(rootPN) = 0 Then rootPN = "product"
    If mIncludeSubassemblies Then modeName = "Assembly structure" Else modeName = "Part list"

    ' --- 2. LOG FILE, SAVE PATH, CACHES ------------------------------------------
    Set mFso = CreateObject("Scripting.FileSystemObject")
    Call StartRunFiles(rootPN)
    Call LogLine("BOM run on " & catDoc.Name & "  (" & modeName & ")")

    Set mDimCache = CreateObject("Scripting.Dictionary")
    Set mMatCache = CreateObject("Scripting.Dictionary")
    Set mBodyCache = CreateObject("Scripting.Dictionary")
    Set mPropCache = CreateObject("Scripting.Dictionary")
    Set mAssyKids = CreateObject("Scripting.Dictionary")
    Set mRollCache = CreateObject("Scripting.Dictionary")
    Set mPromoted = CreateObject("Scripting.Dictionary")
    Set mNotLoaded = CreateObject("Scripting.Dictionary")
    Set mNoPartNumber = CreateObject("Scripting.Dictionary")
    Set mProvisional = CreateObject("Scripting.Dictionary")
    Set mBodyValCache = CreateObject("Scripting.Dictionary")
    Set mShotFiles = CreateObject("Scripting.Dictionary")
    Set mShotSigs = CreateObject("Scripting.Dictionary")
    Set mPartDocs = CreateObject("Scripting.Dictionary")
    mShotNo = 0
    Set mSheetNames = CreateObject("Scripting.Dictionary")
    mSheetNames.CompareMode = vbTextCompare
    mBBoxDisabled = False
    mSheetNo = 0
    mPump = 0

    ' --- 3. THE QUESTIONS ----------------------------------------------------------
    splitLevel = AskSplitLevel(rootProd)
    If splitLevel < 0 Then
        Call LogLine("cancelled at the split level question")
        On Error Resume Next
        CATIA.DisplayFileAlerts = True
        CATIA.StatusBar = ""
        Err.Clear
        Exit Sub
    End If

    mShotParts = CAPTURE_PART_SHOTS
    mShotAssy = CAPTURE_ASSEMBLY_SHOTS
    If ASK_PICTURES And (mShotParts Or mShotAssy) Then
        ans = MsgBox("Take the pictures too?" & vbCrLf & vbCrLf & _
                     "Yes = all numbers first, then one picture per row" & vbCrLf & _
                     "No  = numbers only - the fastest run", _
                     vbYesNoCancel + vbQuestion, "BOM - pictures")
        If ans = vbCancel Then
            Call LogLine("cancelled at the pictures question")
            On Error Resume Next
            CATIA.DisplayFileAlerts = True
            Err.Clear
            Exit Sub
        End If
        If ans = vbNo Then
            mShotParts = False
            mShotAssy = False
        End If
    End If
    Call LogLine("split level " & splitLevel & ", pictures: " & IIf(mShotParts Or mShotAssy, "yes", "no"))
    Call LogTreeNames(rootProd)

    ' --- 4. SESSION --------------------------------------------------------------
    On Error Resume Next
    CATIA.HSOSynchronized = False    ' otherwise every Selection.Add repaints the tree
    If FORCE_DESIGN_MODE Then rootProd.ApplyWorkMode 2
    Err.Clear
    On Error GoTo Fail

    mTempDir = "C:\Temp\catia_bom_pictures"
    Call EnsureFolder(mTempDir)
    ' leftovers of an aborted run must never end up in a new sheet
    On Error Resume Next
    mFso.DeleteFile mTempDir & "\*.*", True
    Err.Clear
    On Error GoTo Fail

    ' --- 5. EXCEL ------------------------------------------------------------------
    ' An Excel of its own: what the user does in his own Excel during the run
    ' (a cell in edit mode, an open dialog) can then never make a write fail.
    If Not StartExcel(firstSheet) Then Err.Raise vbObjectError + 513, , "Excel could not be started"
    Call KeepOneSheet(mXlBook)

    ' --- 6. MAIN WINDOW: geometry only, fixed camera ----------------------------
    Call PrepareWindow(rootWin, rootViewer, rootLayout, rootLayoutSet)
    compassToggled = ToggleCompass()

    If splitLevel = 0 Then

        ' ===== ONE SHEET FOR THE WHOLE PRODUCT ==================================
        Set xlSheet = firstSheet
        Call NameSheet(xlSheet, "BOM")
        unitStart = Timer
        unitOK = ProcessTree(rootProd, rootSel, rootViewer, xlSheet, 0, "", "", 0, _
                             nRows, nAssy, nParts, statusTxt)
        totRows = nRows: totAssy = nAssy: totParts = nParts
        Call LogLine("sheet BOM: " & statusTxt & ", " & nRows & " rows, " & _
                     Format(Elapsed(unitStart), "0") & " s")
        ' (a failure is reported in the final message - AFTER CATIA is restored
        '  and the workbook is saved, never before)

    Else

        ' ===== ONE SHEET PER SUB-ASSEMBLY ON THE SPLIT LEVEL =====================
        Set dUnitQty = CreateObject("Scripting.Dictionary")
        Set dUnitRef = CreateObject("Scripting.Dictionary")
        Set dUnitFirst = CreateObject("Scripting.Dictionary")
        Set dUnitDesc = CreateObject("Scripting.Dictionary")
        Set dUnitPath = CreateObject("Scripting.Dictionary")
        Set colUnitInst = New Collection
        Set colAbove = New Collection
        Set colLooseInst = New Collection
        Set colLooseLevel = New Collection
        Set colLooseFirst = New Collection
        Set colLooseKey = New Collection
        Set colLooseIsAssy = New Collection
        Set colLooseParts = New Collection
        Set colFbKey = New Collection
        Set colFbSheet = New Collection
        Set colFbIdxRow = New Collection

        Call SetStatus("BOM: collecting the components on level " & splitLevel & " ...")
        Call CollectUnits(rootProd, 1, splitLevel, "", "", "", dUnitQty, dUnitRef, dUnitFirst, _
                          dUnitDesc, dUnitPath, colUnitInst, colAbove, _
                          colLooseInst, colLooseLevel, colLooseFirst, colLooseKey, _
                          colLooseIsAssy, colLooseParts)
        Call LogLine(dUnitQty.Count & " components on level " & splitLevel & ", " & _
                     colLooseParts.Count & " parts above it")

        Set idxSheet = firstSheet
        Call NameSheet(idxSheet, "Index")
        Call WriteIndexHeader(idxSheet, splitLevel, idxHdrOK)
        idxRow = 3

        ' how many components would get the same (shortened) sheet name - those
        ' keep their part number in it, so they can be told apart
        Set dLabelCount = CreateObject("Scripting.Dictionary")
        dLabelCount.CompareMode = vbTextCompare
        For Each unitKey In dUnitQty.Keys
            lbl = SheetLabel(TreeName(dUnitRef(unitKey), CStr(unitKey)), CStr(unitKey))
            If dLabelCount.Exists(lbl) Then
                dLabelCount(lbl) = dLabelCount(lbl) + 1
            Else
                dLabelCount.Add lbl, 1
            End If
        Next unitKey

        ' --- the components that are CATProducts of their own: own window ------
        For Each unitKey In dUnitQty.Keys
            nUnits = nUnits + 1
            unitStart = Timer
            nRows = 0: nAssy = 0: nParts = 0
            statusTxt = "not started"
            unitName = CStr(unitKey)
            Set xlSheet = Nothing
            mXlSkipAt = -1                   ' a "skip" given for an earlier component ends here
            Call CheckExcel                  ' Excel gone: stop the whole run here
            ' the Index header (and its text formats) before any more Index rows
            If Not (idxHdrOK And mIdxFmtOK) Then Call WriteIndexHeader(idxSheet, splitLevel, idxHdrOK)
            ' an error in here ends this component only, never the whole run
            On Error GoTo UnitFail

            Set unitInst = dUnitRef(unitKey)
            ' the name of the component as the CATIA tree shows it
            unitName = TreeName(unitInst, CStr(unitKey))
            Call LogLine("=== component " & nUnits & " of " & dUnitQty.Count & ": " & unitName)
            lbl = SheetLabel(unitName, CStr(unitKey))
            If dLabelCount.Exists(lbl) Then
                If dLabelCount(lbl) > 1 And lbl <> unitName And Len(CStr(unitKey)) < 25 Then
                    lbl = Left$(lbl, 31 - Len(CStr(unitKey)) - 1) & "-" & CStr(unitKey)
                End If
            End If
            Set xlSheet = AddSheet(mXlBook, lbl)
            If xlSheet Is Nothing Then Err.Raise vbObjectError + 514, , "Excel could not add a sheet"

            Call NoteIndexRow(idxSheet, idxRow, nUnits, unitName, CStr(dUnitDesc(unitKey)), _
                               splitLevel, CLng(dUnitQty(unitKey)), CStr(dUnitFirst(unitKey)), _
                               CStr(dUnitPath(unitKey)), mSheetName)

            ' NO ApplyWorkMode on the component here: it is recursive and would
            ' load every CATPart of the component in ONE blocking call. Its parts
            ' are promoted one per row while the numbers are read.
            Call SetStatus("BOM: component " & nUnits & " of " & dUnitQty.Count & _
                           " - opening " & CStr(unitKey) & " ...")
            ownWindow = False
            Set unitDoc = Nothing
            Set unitWin = Nothing
            If OPEN_UNITS_IN_NEW_WINDOW Then Set unitDoc = OwnDocumentOf(unitInst, CStr(unitKey), catDoc)

            If Not unitDoc Is Nothing Then
                On Error Resume Next
                Set unitWin = unitDoc.NewWindow
                If Not unitWin Is Nothing Then unitWin.Activate
                Err.Clear
                On Error GoTo UnitFail
                ownWindow = Not (unitWin Is Nothing)
            End If

            If ownWindow Then
                Set unitViewer = Nothing
                Call PrepareWindow(unitWin, unitViewer, unitLayout, unitLayoutSet)
                ' every new window starts with its own compass - hide it there too
                If compassToggled Then compassToggled = ToggleCompass()
                Set unitSel = unitDoc.Selection
                Set unitRoot = unitDoc.Product

                unitOK = ProcessTree(unitRoot, unitSel, unitViewer, xlSheet, splitLevel, _
                                     CStr(dUnitFirst(unitKey)), CStr(unitKey), CLng(dUnitQty(unitKey)), _
                                     nRows, nAssy, nParts, statusTxt)

                ' close the component window and go back to the product
                On Error Resume Next
                unitWin.Close
                rootWin.Activate
                Err.Clear
                On Error GoTo UnitFail
                Set unitWin = Nothing

                Call NoteIndexResult(idxSheet, idxRow, nRows, statusTxt, Elapsed(unitStart), "own window")
                Call LogLine("component " & CStr(unitKey) & ": " & statusTxt & ", " & nRows & " rows, " & _
                             Format(Elapsed(unitStart), "0") & " s")
                totRows = totRows + nRows: totAssy = totAssy + nAssy: totParts = totParts + nParts
                If unitOK Then nUnitsOK = nUnitsOK + 1
            Else
                ' done later in the main window, once the rest of the scene is hidden
                colFbKey.Add unitKey
                colFbSheet.Add xlSheet
                colFbIdxRow.Add idxRow
                Call LogLine("component " & CStr(unitKey) & ": no file of its own - done in the main window")
            End If

            Call SaveWorkbook(False)
            GoTo UnitNext

UnitCleanup:
            ' (reached from UnitFail below, after the error is cleared)
            On Error Resume Next
            Call LogLine("component " & CStr(unitKey) & " " & statusTxt)
            If Not unitWin Is Nothing Then unitWin.Close
            Set unitWin = Nothing
            rootWin.Activate
            If xlSheet Is Nothing Then
                Call NoteIndexRow(idxSheet, idxRow, nUnits, unitName, CStr(dUnitDesc(unitKey)), _
                                   splitLevel, CLng(dUnitQty(unitKey)), CStr(dUnitFirst(unitKey)), _
                                   CStr(dUnitPath(unitKey)), "")
            End If
            Call NoteIndexResult(idxSheet, idxRow, nRows, statusTxt, Elapsed(unitStart), "-")
            Err.Clear

UnitNext:
            On Error GoTo Fail
            idxRow = idxRow + 1
            DoEvents
        Next unitKey

        ' --- the rest is done in the main window ---------------------------------
        If colFbKey.Count > 0 Or colLooseInst.Count > 0 Then
            ' scene: every assembly above the split level visible, every component
            ' and every loose part hidden - then one thing at a time is shown
            Call SetShowCollection(rootSel, colAbove, 0)
            Call SetShowCollection(rootSel, colUnitInst, 1)
            Call SetShowCollection(rootSel, colLooseParts, 1)
            rootSceneSet = True

            For i = 1 To colFbKey.Count
                mXlSkipAt = -1
                Call CheckExcel
                unitKey = colFbKey(i)
                unitStart = Timer
                Set unitInst = dUnitRef(unitKey)

                Set xlSheet = colFbSheet(i)
                mSheetName = SheetNameOf(xlSheet)
                Call LogLine("=== component " & CStr(unitKey) & " (main window)")
                Call SetShowSingle(rootSel, unitInst, 0)         ' only this branch
                unitOK = ProcessTree(unitInst, rootSel, rootViewer, xlSheet, splitLevel, _
                                     CStr(dUnitFirst(unitKey)), CStr(unitKey), CLng(dUnitQty(unitKey)), _
                                     nRows, nAssy, nParts, statusTxt)
                Call SetShowSingle(rootSel, unitInst, 1)

                Call NoteIndexResult(idxSheet, CLng(colFbIdxRow(i)), nRows, statusTxt, _
                                      Elapsed(unitStart), "main window")
                Call LogLine("component " & CStr(unitKey) & ": " & statusTxt & ", " & nRows & " rows, " & _
                             Format(Elapsed(unitStart), "0") & " s")
                totRows = totRows + nRows: totAssy = totAssy + nAssy: totParts = totParts + nParts
                If unitOK Then nUnitsOK = nUnitsOK + 1
                Call SaveWorkbook(False)
                DoEvents
            Next i

            If colLooseInst.Count > 0 Then
                mXlSkipAt = -1
                Call CheckExcel
                unitStart = Timer
                If splitLevel = 1 Then
                    looseName = "Parts on level 1"
                    looseText = "the parts directly below the product (first level parts)"
                Else
                    looseName = "Above level " & splitLevel
                    looseText = "parts that are not inside any component of level " & splitLevel & _
                                IIf(mIncludeSubassemblies, ", and the assemblies above it", "")
                End If
                Call LogLine("=== " & looseName)
                Set xlSheet = AddSheet(mXlBook, looseName)
                If Not xlSheet Is Nothing Then
                    Call NoteIndexRow(idxSheet, idxRow, 0, "(" & looseName & ")", looseText, _
                                       0, colLooseParts.Count, "", "", mSheetName)
                    unitOK = ProcessLoose(colLooseInst, colLooseLevel, colLooseFirst, colLooseKey, _
                                          colLooseIsAssy, rootSel, rootViewer, xlSheet, _
                                          nRows, nAssy, nParts, statusTxt)
                    Call NoteIndexResult(idxSheet, idxRow, nRows, statusTxt, Elapsed(unitStart), "main window")
                    Call LogLine(looseName & ": " & statusTxt & ", " & nRows & " rows, " & _
                                 Format(Elapsed(unitStart), "0") & " s")
                    totRows = totRows + nRows: totAssy = totAssy + nAssy: totParts = totParts + nParts
                Else
                    Call LogLine(looseName & ": no sheet - Excel could not add one")
                    Call NoteIndexRow(idxSheet, idxRow, 0, "(" & looseName & ")", looseText, _
                                0, colLooseParts.Count, "", "", "")
                    Call NoteIndexResult(idxSheet, idxRow, 0, "stopped: Excel could not add the sheet", _
                                   Elapsed(unitStart), "-")
                End If
            End If

            ' everything visible again
            Call SetShowCollection(rootSel, colUnitInst, 0)
            Call SetShowCollection(rootSel, colLooseParts, 0)
            rootSceneSet = False
        End If

        ' Index writes Excel could not take during the run get a second chance
        mXlSkipAt = -1
        If Not (idxHdrOK And mIdxFmtOK) Then Call WriteIndexHeader(idxSheet, splitLevel, idxHdrOK)
        Call RedoIndexWrites(idxSheet)
        Call FinishIndexLayout(idxSheet)
    End If

    ' --- 7. RESTORE THE SESSION --------------------------------------------------
    On Error Resume Next
    rootWin.Activate
    If rootLayoutSet Then rootWin.Layout = rootLayout
    If compassToggled Then CATIA.StartCommand COMPASS_COMMAND
    If Not rootViewer Is Nothing Then rootViewer.Reframe
    CATIA.RefreshDisplay = True
    CATIA.HSOSynchronized = True
    CATIA.DisplayFileAlerts = True
    CATIA.StatusBar = ""
    rootSel.Clear
    Err.Clear
    On Error GoTo Fail

    Call LogLine("finished after " & Format(Elapsed(startTime), "0") & " s, " & totRows & " rows")
    Call WriteLogSheet(modeName & ", split level " & splitLevel & ", " & totRows & " rows, " & _
                       Format(Elapsed(startTime) / 60#, "0.0") & " min")
    Call ShowSheet(firstSheet)
    Call SaveWorkbook(True)
    Call XlRestore

    If splitLevel = 0 Then
        If Not unitOK Then
            If Left$(statusTxt, 10) = "INCOMPLETE" Then
                failMsg = "BOM exported, but the sheet is INCOMPLETE"
            Else
                failMsg = "The run stopped early"
            End If
            MsgBox failMsg & "  (" & modeName & ")" & vbCrLf & statusTxt & vbCrLf & vbCrLf & _
                   "Rows written:     " & totRows & vbCrLf & _
                   "Time: " & Format(Elapsed(startTime) / 60#, "0.0") & " min" & vbCrLf & vbCrLf & _
                   NotLoadedText() & SavedText(), vbExclamation
            Exit Sub
        End If
        MsgBox "BOM exported  (" & modeName & ")" & vbCrLf & vbCrLf & _
               "Rows written:     " & totRows & vbCrLf & _
               "  sub-assemblies: " & totAssy & vbCrLf & _
               "  parts:          " & totParts & vbCrLf & _
               "Time: " & Format(Elapsed(startTime) / 60#, "0.0") & " min" & vbCrLf & vbCrLf & _
               NotLoadedText() & SavedText(), vbInformation
    Else
        MsgBox "BOM exported  (" & modeName & ", split at level " & splitLevel & ")" & vbCrLf & vbCrLf & _
               "Component sheets: " & nUnits & "  (" & nUnitsOK & " complete)" & vbCrLf & _
               "Rows written:     " & totRows & vbCrLf & _
               "Time: " & Format(Elapsed(startTime) / 60#, "0.0") & " min" & vbCrLf & vbCrLf & _
               "The Index sheet links to every component sheet." & vbCrLf & vbCrLf & _
               NotLoadedText() & SavedText(), vbInformation
    End If
    Exit Sub

' ------------------------------------------------------------------------------
' An error inside one component: noted, then the run goes on with the next one
' ------------------------------------------------------------------------------
UnitFail:
    statusTxt = "stopped: " & Err.Number & " - " & Err.Description
    Resume UnitCleanup

' ------------------------------------------------------------------------------
' Something went wrong outside a component: put CATIA and Excel back into a
' usable state and save what is there before reporting.
' ------------------------------------------------------------------------------
Fail:
    failMsg = "The BOM run stopped with error " & Err.Number & " - " & Err.Description
    Resume FailCleanup

FailCleanup:
    On Error Resume Next
    Call LogLine(failMsg)
    If Not unitWin Is Nothing Then unitWin.Close
    If Not rootWin Is Nothing Then rootWin.Activate
    If rootSceneSet Then
        Call SetShowCollection(rootSel, colUnitInst, 0)
        Call SetShowCollection(rootSel, colLooseParts, 0)
    End If
    If rootLayoutSet Then rootWin.Layout = rootLayout
    If compassToggled Then CATIA.StartCommand COMPASS_COMMAND
    If Not rootViewer Is Nothing Then rootViewer.Reframe
    CATIA.HSOSynchronized = True
    CATIA.RefreshDisplay = True
    CATIA.DisplayFileAlerts = True
    CATIA.StatusBar = ""
    If Not rootSel Is Nothing Then rootSel.Clear
    Call SaveWorkbook(True)
    Call XlRestore
    Err.Clear

    If mXlLost Then
        MsgBox failMsg & vbCrLf & vbCrLf & _
               "Visibility, spec tree and compass were restored." & vbCrLf & _
               "Excel was closed during the run." & vbCrLf & vbCrLf & _
               SavedText(), vbCritical
    Else
        MsgBox failMsg & vbCrLf & vbCrLf & _
               "Visibility, spec tree and compass were restored." & vbCrLf & _
               "The rows written so far are in the Excel workbook." & vbCrLf & vbCrLf & _
               SavedText(), vbCritical
    End If

End Sub


' ==============================================================================
' ONE SHEET FROM A (SUB)TREE
'   treeRoot     the product whose content goes on the sheet
'   levelOffset  0 for the whole product; n when the sheet is a component on
'                level n - the component itself becomes the first row and its
'                content is written with the levels it has in the product
' Returns False if the sheet could not be finished; statusTxt says why.
' An error here never stops the whole run - it only ends this sheet.
' ==============================================================================
Private Function ProcessTree(treeRoot As Product, sel As Selection, oViewer As Viewer, _
                             xlSheet As Object, ByVal levelOffset As Long, _
                             ByVal firstLevel As String, ByVal unitPartNum As String, _
                             ByVal unitQty As Long, ByRef nRows As Long, ByRef nAssy As Long, _
                             ByRef nParts As Long, ByRef statusTxt As String) As Boolean
    Dim r As Long, t0 As Double
    Dim hdrOK As Boolean, nLost As Long

    ProcessTree = False
    nRows = 0: nAssy = 0: nParts = 0
    r = 2
    statusTxt = ""
    t0 = Timer
    On Error GoTo SheetFail

    Call ResetSheetData
    hdrOK = WriteHeader(xlSheet)

    ' --- the tree of this sheet (numbers are NOT read here) ---
    Call SetStatus("BOM: reading the tree of " & mSheetName & " ...")
    Call LogLine(mSheetName & ": reading the tree")
    Call TraverseTree(treeRoot, mQty, mRef, mDesc, mProps, mLevel, mBodies, mPartNum, mFirst, _
                      mIsAssy, mNodeLeaves, mLeaves, levelOffset + 1, firstLevel, unitPartNum)
    Call LogLine(mSheetName & ": " & mQty.Count & " rows, " & mLeaves.Count & " part instances")

    ' --- the component itself is the first row of its sheet ---
    ' (quantity 1: the whole sheet describes ONE piece - the Index has the count.
    '  The part list promises no assembly rows, so it has no component row.)
    If levelOffset > 0 And mIncludeSubassemblies Then
        Call WriteUnitRow(treeRoot, r, levelOffset, unitPartNum, 1, firstLevel)
        nAssy = nAssy + 1
        r = r + 1
    End If

    ' --- PASS 1: all numbers, display frozen, nothing shown or hidden ---
    Call WriteDataRows(r, nAssy, nParts)
    nRows = r - 2
    Call FillAssemblyRows
    nLost = FlushSheet(xlSheet, False)
    Call RetryLostWrites(xlSheet, hdrOK, nLost)

    ' the numbers are complete - make them visible before the pictures start
    Call FinishSheetLayout(xlSheet)
    Call LogLine(mSheetName & ": numbers done after " & Format(Elapsed(t0), "0") & " s")
    Call SaveWorkbook(False)

    ' --- PASS 2: the pictures ---
    Call CheckExcel
    Call TakePictures(sel, oViewer, xlSheet)

    ' rows Excel never took: the sheet is NOT reported as complete
    If Not hdrOK Or nLost > 0 Then
        nRows = nRows - nLost
        statusTxt = SheetLostText(hdrOK, nLost)
        Call LogLine(mSheetName & " " & statusTxt)
        Call MarkSheetIncomplete(xlSheet, statusTxt)
        Exit Function
    End If

    ProcessTree = True
    statusTxt = "OK"
    Exit Function

SheetFail:
    statusTxt = "stopped: " & Err.Number & " - " & Err.Description
    nRows = r - 2
    Resume SheetCleanup

SheetCleanup:
    On Error Resume Next
    Call LogLine(mSheetName & " " & statusTxt)
    Call SetShowCollection(sel, mLeaves, 0)
    CATIA.RefreshDisplay = True
    ' whatever was read so far still goes into the sheet
    If Not mFlushed Then Call FlushSheet(xlSheet, False)
    Call MarkSheetIncomplete(xlSheet, statusTxt)
    Err.Clear
End Function


' ==============================================================================
' THE SHEET OF THE PARTS ABOVE THE SPLIT LEVEL
' (parts that do not sit inside any component of the split level)
' ==============================================================================
Private Function ProcessLoose(colInst As Collection, colLevel As Collection, _
                              colFirst As Collection, colKey As Collection, _
                              colIsAssy As Collection, sel As Selection, oViewer As Viewer, _
                              xlSheet As Object, ByRef nRows As Long, ByRef nAssy As Long, _
                              ByRef nParts As Long, ByRef statusTxt As String) As Boolean
    Dim i As Long, r As Long
    Dim oProd As Product, pn As String, sKey As String
    Dim isAssembly As Boolean
    Dim hdrOK As Boolean, nLost As Long

    ProcessLoose = False
    nRows = 0: nAssy = 0: nParts = 0
    r = 2
    statusTxt = ""
    On Error GoTo LooseFail

    Call ResetSheetData
    hdrOK = WriteHeader(xlSheet)

    ' the rows are keyed like everywhere else (MakeRowKey), so the part list
    ' groups by part number and the assembly structure by parent path
    For i = 1 To colInst.Count
        Set oProd = colInst(i)
        isAssembly = colIsAssy(i)
        pn = ""
        On Error Resume Next
        pn = oProd.PartNumber
        Err.Clear
        On Error GoTo LooseFail
        If Len(pn) = 0 And Not isAssembly Then pn = ProvisionalName(oProd)
        If Len(pn) = 0 Then pn = oProd.Name

        If Not isAssembly Then
            mLeaves.Add oProd
            mAllNodes.Add oProd
        End If

        sKey = CStr(colKey(i))
        If mQty.Exists(sKey) Then
            mQty(sKey) = mQty(sKey) + 1
            Call AddFirstLevel(mFirst, sKey, CStr(colFirst(i)))
        Else
            Call AddBomRow(mQty, mRef, mDesc, mProps, mLevel, mPartNum, mFirst, mIsAssy, _
                           sKey, oProd, pn, CStr(colFirst(i)), CInt(colLevel(i)), isAssembly)
        End If
        Call Pump
    Next i

    Call WriteDataRows(r, nAssy, nParts)
    nRows = r - 2
    Call FillAssemblyRows
    nLost = FlushSheet(xlSheet, False)
    Call RetryLostWrites(xlSheet, hdrOK, nLost)
    Call FinishSheetLayout(xlSheet)
    Call SaveWorkbook(False)
    Call CheckExcel
    ' RunBOM has hidden these parts for the main-window scene - visible again,
    ' so the first picture (and the reference of the show / hide check) has them
    Call SetShowCollection(sel, mLeaves, 0)
    Call TakePictures(sel, oViewer, xlSheet)

    If Not hdrOK Or nLost > 0 Then
        nRows = nRows - nLost
        statusTxt = SheetLostText(hdrOK, nLost)
        Call LogLine(mSheetName & " " & statusTxt)
        Call MarkSheetIncomplete(xlSheet, statusTxt)
        Exit Function
    End If

    ProcessLoose = True
    statusTxt = "OK"
    Exit Function

LooseFail:
    statusTxt = "stopped: " & Err.Number & " - " & Err.Description
    nRows = r - 2
    Resume LooseCleanup

LooseCleanup:
    On Error Resume Next
    Call LogLine(mSheetName & " " & statusTxt)
    CATIA.RefreshDisplay = True
    If Not mFlushed Then Call FlushSheet(xlSheet, False)
    Call MarkSheetIncomplete(xlSheet, statusTxt)
    Err.Clear
End Function



' ==============================================================================
' PASS 1 - ALL NUMBERS OF ONE SHEET (into the row buffer, not into Excel yet)
' Nothing is shown, hidden or captured here and the display is frozen, so CATIA
' only loads and measures. Only PARTS are loaded: an assembly row is added up
' from its parts afterwards (FillAssemblyRows). Every picture that will be
' needed is noted as a job.
' Errors are not handled here on purpose: they go up to the sheet's handler.
' ==============================================================================
Private Sub WriteDataRows(ByRef r As Long, ByRef nAssy As Long, ByRef nParts As Long)
    Dim uniqueKey As Variant, strPartNum As String
    Dim oPartProd As Product, isAssy As Boolean
    Dim oPart As Part, oPartDoc As PartDocument, hasPart As Boolean
    Dim propsArray As Variant, cacheArr As Variant
    Dim dMass As Double, dVol As Double, dArea As Double
    Dim dims(2) As Double
    Dim strMatName As String, dDensity As Double
    Dim solidNames() As String, nSolid As Integer
    Dim bodyNames As Variant, bi As Integer, oBody As Body
    Dim bMass As Double, bVol As Double, bArea As Double
    Dim bDens As Double, bDensInertia As Double, bMat As String
    Dim nDone As Long, nTotal As Long
    Dim tRow As Double, tLoad As Double, tRowAll As Double
    Dim notLoaded As Boolean, bodyKey As String, bv As Variant
    Dim rowStart As Long, dispPn As String, realPn As String, realDesc As String, shotOK As Boolean

    On Error Resume Next
    CATIA.RefreshDisplay = False
    Err.Clear
    On Error GoTo 0

    nTotal = mQty.Count
    For Each uniqueKey In mQty.Keys

        nDone = nDone + 1
        strPartNum = mPartNum(uniqueKey)
        Set oPartProd = mRef(uniqueKey)
        isAssy = mIsAssy(uniqueKey)

        Call BufSet(r, 2, mLevel(uniqueKey))
        Call BufSet(r, 3, strPartNum)
        Call BufSet(r, 4, mDesc(uniqueKey))
        Call BufSet(r, 5, mQty(uniqueKey))
        Call BufSet(r, 14, mFirst(uniqueKey))
        rowStart = r
        dispPn = strPartNum

        If isAssy Then
            ' ---------- SUB-ASSEMBLY: nothing is loaded or measured here ----------
            ' (ApplyWorkMode / Analyze on an assembly work on its WHOLE subtree in
            '  one blocking call - the values are added up from its parts instead)
            nAssy = nAssy + 1
            Call BufSet(r, 15, "Assembly")
            mBoldRows.Add r
            mAssyRows.Add Array(r, strPartNum, oPartProd)
            ' (the extra sheet above the split level has no leaf lists - no picture there)
            If mShotAssy And mNodeLeaves.Exists(CStr(uniqueKey)) Then
                Call AddJob(r, "ASSY", CStr(uniqueKey), "")
            Else
                Call BufSet(r, 1, "-")
            End If

        Else
            ' ---------- PART (or a sub-assembly that is not loaded) ----------
            Call SetStatus("BOM " & mSheetName & ": numbers, row " & nDone & " of " & _
                           nTotal & " - " & strPartNum)
            ' written BEFORE the part is loaded: if CATIA hangs, the last line of
            ' the log names the part
            Call LogLine(mSheetName & "  row " & nDone & "/" & nTotal & "  " & strPartNum)
            tRow = Timer

            ' A sub-assembly that is NOT LOADED has no children and looks like a
            ' part. ApplyWorkMode 2 on it would load its whole content in ONE
            ' blocking call - so it is neither loaded nor measured: it gets an
            ' "Assembly" row marked NOT LOADED, and the final message lists it.
            notLoaded = mNotLoaded.Exists(strPartNum)
            ' a row named before its part number could be read is loaded as a part
            ' only when something says it IS a part - a sub-assembly would be
            ' loaded completely in design mode, in one blocking call
            If Not notLoaded And mProvisional.Exists(strPartNum) Then
                If Not HasPartEvidence(oPartProd) Then
                    notLoaded = True
                    If Not mNotLoaded.Exists(strPartNum) Then mNotLoaded.Add strPartNum, True
                    If Not mNoPartNumber.Exists(strPartNum) Then mNoPartNumber.Add strPartNum, True
                    Call LogLine(mSheetName & "  NOT LOADED (nothing says it is a part): " & strPartNum)
                End If
            End If
            If Not notLoaded And Not mPromoted.Exists(strPartNum) Then
                notLoaded = IsUnloadedAssembly(oPartProd)
                If notLoaded Then
                    mNotLoaded.Add strPartNum, True
                    Call LogLine(mSheetName & "  NOT LOADED sub-assembly: " & strPartNum & _
                                 " - its content is missing in the BOM")
                End If
            End If
        End If

        If notLoaded And Not isAssy And mNoPartNumber.Exists(strPartNum) Then
            ' a component that could not be loaded: whether it is a part or an
            ' assembly is unknown (not counted as either)
            Call BufSet(r, 15, "Not loaded")
            Call BufSet(r, 1, "-")
            For bi = 6 To 11
                Call BufSet(r, bi, "N/A")
            Next bi
            Call BufSet(r, 12, "NOT LOADED - could not be loaded")
            Call BufSet(r, 13, "N/A")

        ElseIf notLoaded And Not isAssy Then
            nAssy = nAssy + 1
            Call BufSet(r, 15, "Assembly")
            mBoldRows.Add r
            Call BufSet(r, 1, "-")
            For bi = 6 To 11
                Call BufSet(r, bi, "N/A")
            Next bi
            Call BufSet(r, 12, "NOT LOADED (or empty) - content missing")
            Call BufSet(r, 13, "N/A")

        ElseIf Not isAssy Then
            nParts = nParts + 1
            Call BufSet(r, 15, "Part")

            ' promote THIS part to design mode first - mass, volume, area and the
            ' inertia box are all 0 on a component in visualization mode. The
            ' CATPart document is shared by all instances of the part, so once
            ' per part number is enough.
            If Not mPromoted.Exists(strPartNum) Or mProvisional.Exists(strPartNum) Then
                On Error Resume Next
                oPartProd.ApplyWorkMode 2
                Err.Clear
                On Error GoTo 0
                If Not mPromoted.Exists(strPartNum) Then mPromoted.Add strPartNum, True
            End If
            tLoad = Elapsed(tRow)

            hasPart = False
            Set oPart = Nothing
            Set oPartDoc = Nothing
            hasPart = LoadPartDocument(oPartProd, oPart, oPartDoc)
            ' the picture pass needs the same CATPart - also when it was only
            ' found through the shape representation
            If hasPart Then
                If Not mPartDocs.Exists(strPartNum) Then mPartDocs.Add strPartNum, oPartDoc
            End If

            ' a row named before the part number could be read: now that the
            ' part is loaded, its real part number and description
            If mProvisional.Exists(strPartNum) Then
                realPn = ""
                realDesc = ""
                On Error Resume Next
                realPn = oPartProd.PartNumber
                realDesc = oPartProd.DescriptionRef
                If Len(realPn) = 0 And hasPart Then realPn = oPartDoc.Product.PartNumber
                If Len(realDesc) = 0 And hasPart Then realDesc = oPartDoc.Product.DescriptionRef
                Err.Clear
                On Error GoTo 0
                If Len(realPn) > 0 Then
                    dispPn = realPn
                    Call BufSet(r, 3, realPn)
                    ' a first level node: its "First Level" name with the real part number
                    If mLevel(uniqueKey) = 1 Then
                        mFirst(uniqueKey) = FirstLevelLabel(oPartProd, realPn)
                        Call BufSet(r, 14, mFirst(uniqueKey))
                    End If
                End If
                If Len(realDesc) > 0 Then Call BufSet(r, 4, realDesc)
            End If

            ' multi-body parts: the body names are cached per part number, so a
            ' part that appears on several sheets gets its body rows on every one
            If hasPart And DETECT_MULTIBODY Then
                If Not mBodyCache.Exists(strPartNum) Then
                    nSolid = GetSolidBodyNames(oPart, solidNames)
                    If nSolid > 1 Then
                        mBodyCache.Add strPartNum, solidNames
                    Else
                        mBodyCache.Add strPartNum, Empty
                    End If
                End If
                If IsArray(mBodyCache(strPartNum)) Then
                    If Not mBodies.Exists(strPartNum) Then mBodies.Add strPartNum, mBodyCache(strPartNum)
                End If
            End If

            ' --- MASS / VOLUME / AREA (once per part number; a zero is not cached) ---
            If mPropCache.Exists(strPartNum) Then
                propsArray = mPropCache(strPartNum)
            Else
                dMass = 0: dVol = 0: dArea = 0
                Call ReadProductProps(oPartProd, dMass, dVol, dArea)
                ' nothing through the instance: ask the CATPart itself
                If dMass <= 0 And dVol <= 0 And dArea <= 0 And hasPart Then
                    Call ReadProductProps(oPartDoc.Product, dMass, dVol, dArea)
                End If
                propsArray = Array(dMass, dVol, dArea)
                If dMass > 0 Or dVol > 0 Or dArea > 0 Then mPropCache.Add strPartNum, propsArray
            End If
            Call PutNumber(r, 6, propsArray(0))
            Call PutNumber(r, 7, propsArray(1))
            Call PutNumber(r, 8, propsArray(2))

            ' --- MATERIAL & DENSITY (once per part number) ---
            strMatName = "N/A"
            dDensity = 0
            If hasPart Then
                If mMatCache.Exists(strPartNum) Then
                    cacheArr = mMatCache(strPartNum)
                    strMatName = CStr(cacheArr(0))
                    dDensity = CDbl(cacheArr(1))
                Else
                    Call GetMaterialAndDensity(oPart, strMatName, dDensity)
                    If dDensity > 0 Or strMatName <> "N/A" Then _
                        mMatCache.Add strPartNum, Array(strMatName, dDensity)
                End If
            End If
            Call BufSet(r, 12, strMatName)
            Call PutNumber(r, 13, dDensity)

            ' --- DIMENSIONS (once per part number) ---
            dims(0) = 0: dims(1) = 0: dims(2) = 0
            If mDimCache.Exists(strPartNum) Then
                cacheArr = mDimCache(strPartNum)
                dims(0) = CDbl(cacheArr(0)): dims(1) = CDbl(cacheArr(1)): dims(2) = CDbl(cacheArr(2))
            Else
                ' only in "BOX" mode: this creates temporary geometry in the CATPart
                If hasPart And UCase$(MEASURE_MODE) = "BOX" Then Call GetBoundingBoxDims(oPart, dims)
                If dims(0) < 0.1 Then
                    On Error Resume Next
                    Call GetInertiaDims(oPartProd.ReferenceProduct, dims)
                    If dims(0) < 0.1 And hasPart Then Call GetInertiaDims(oPartDoc.Product, dims)
                    Err.Clear
                    On Error GoTo 0
                End If
                If dims(0) > 0 Then mDimCache.Add strPartNum, Array(dims(0), dims(1), dims(2))
            End If
            Call PutNumber(r, 9, dims(0))
            Call PutNumber(r, 10, dims(1))
            Call PutNumber(r, 11, dims(2))

            ' a part without any value: the log says what CATIA has of it
            shotOK = True
            If propsArray(0) <= 0 And propsArray(1) <= 0 And propsArray(2) <= 0 And dims(0) <= 0 Then
                Call LogLine(mSheetName & "  NO VALUES " & strPartNum & ": " & PartStateText(oPartProd, hasPart))
                ' a component that could not even be loaded: NOT LOADED, no picture
                If Not hasPart And mProvisional.Exists(strPartNum) And Len(realPn) = 0 Then
                    nParts = nParts - 1
                    Call BufSet(r, 15, "Not loaded")
                    Call BufSet(r, 12, "NOT LOADED - could not be loaded")
                    If Not mNotLoaded.Exists(strPartNum) Then mNotLoaded.Add strPartNum, True
                    If Not mNoPartNumber.Exists(strPartNum) Then mNoPartNumber.Add strPartNum, True
                    shotOK = False
                End If
            End If

            ' --- the picture of this row is taken in pass 2 ---
            If mShotParts And shotOK Then
                Call AddJob(r, "PART", CStr(uniqueKey), "")
            Else
                Call BufSet(r, 1, "-")
            End If

            ' --- MULTI-BODY SUB-ROWS ---
            If hasPart And mBodies.Exists(strPartNum) Then
                bodyNames = mBodies(strPartNum)
                For bi = LBound(bodyNames) To UBound(bodyNames)
                    Set oBody = GetBodyByName(oPart, CStr(bodyNames(bi)))
                    If Not oBody Is Nothing Then
                        r = r + 1

                        Call BufSet(r, 2, mLevel(uniqueKey) + 1)
                        Call BufSet(r, 3, "    " & CStr(bodyNames(bi)))
                        Call BufSet(r, 4, "Body of " & dispPn)
                        Call BufSet(r, 5, mQty(uniqueKey))
                        Call BufSet(r, 14, mFirst(uniqueKey))
                        Call BufSet(r, 15, "Body")
                        Call BufSet(r, 16, CStr(bodyNames(bi)))

                        ' body values: once per part number and body for the whole run
                        bodyKey = strPartNum & "|" & CStr(bodyNames(bi))
                        If mBodyValCache.Exists(bodyKey) Then
                            bv = mBodyValCache(bodyKey)
                        Else
                            bMass = 0: bVol = 0: bArea = 0
                            bDens = 0: bDensInertia = 0
                            bMat = "N/A"

                            Call GetBodyMaterial(oPart, oBody, bMat, bDens)
                            Call GetBodyMassAndDensity(oPartDoc, oBody, bMass, bDensInertia)
                            If bDens <= 0 Then bDens = bDensInertia
                            Call GetBodyVolumeArea(oPartDoc, oPart, oBody, bVol, bArea)
                            If bMass <= 0 And bDens > 0 And bVol > 0 Then bMass = bVol * bDens

                            dims(0) = 0: dims(1) = 0: dims(2) = 0
                            If UCase$(MEASURE_MODE) = "BOX" Then
                                If Not MeasureTarget(oPart, oBody, dims) Then
                                    Call MeasureAABBExtremum(oPart, oPartDoc, oBody, dims)
                                End If
                            End If
                            If dims(0) < 0.1 Then Call GetBodyInertiaDims(oPartDoc, oBody, dims)

                            bv = Array(bMass, bVol, bArea, dims(0), dims(1), dims(2), bMat, bDens)
                            If bMass > 0 Or bVol > 0 Or bArea > 0 Then mBodyValCache.Add bodyKey, bv
                            DoEvents                     ' a part with many bodies takes a while
                        End If

                        Call PutNumber(r, 6, bv(0))
                        Call PutNumber(r, 7, bv(1))
                        Call PutNumber(r, 8, bv(2))
                        Call PutNumber(r, 9, bv(3))
                        Call PutNumber(r, 10, bv(4))
                        Call PutNumber(r, 11, bv(5))
                        Call BufSet(r, 12, CStr(bv(6)))
                        Call PutNumber(r, 13, bv(7))

                        If mShotParts Then
                            Call AddJob(r, "BODY", CStr(uniqueKey), CStr(bodyNames(bi)))
                        Else
                            Call BufSet(r, 1, "-")
                        End If
                    End If
                    Set oBody = Nothing
                Next bi
            End If

            tRowAll = Elapsed(tRow)
            If tRowAll >= SLOW_STEP_SECONDS Then
                Call NoteSlow(tRowAll, mSheetName & "  row " & nDone & "  " & strPartNum & _
                              "  (loading " & Format(tLoad, "0.0") & " s)")
            End If
        End If

        ' the row as the CATIA tree shows it (indented by FlushSheet)
        Call BufSet(rowStart, 16, TreeRowText(oPartProd, dispPn, CLng(mQty(uniqueKey))))

        r = r + 1
        DoEvents
    Next uniqueKey

    On Error Resume Next
    CATIA.RefreshDisplay = True
    Err.Clear
    On Error GoTo 0
End Sub


' ==============================================================================
' THE FIRST ROW OF A COMPONENT SHEET: THE COMPONENT ITSELF
' (its numbers are added up with the other assembly rows - FillAssemblyRows)
' ==============================================================================
Private Sub WriteUnitRow(treeRoot As Product, ByVal r As Long, ByVal lvl As Long, _
                         ByVal pn As String, ByVal qty As Long, ByVal firstLevel As String)
    Dim sDesc As String

    sDesc = ""
    On Error Resume Next
    sDesc = treeRoot.DescriptionRef
    Err.Clear
    On Error GoTo 0

    Call BufSet(r, 2, lvl)
    Call BufSet(r, 3, pn)
    Call BufSet(r, 4, sDesc)
    Call BufSet(r, 5, qty)
    Call BufSet(r, 14, firstLevel)
    Call BufSet(r, 15, "Assembly")
    Call BufSet(r, 16, TreeRowText(treeRoot, pn, 1))
    mBoldRows.Add r
    mAssyRows.Add Array(r, pn, treeRoot)

    ' the whole component, pictured in pass 2
    If mShotAssy Then
        Call AddJob(r, "UNIT", "", "")
    Else
        Call BufSet(r, 1, "-")
    End If
End Sub


' ==============================================================================
' PASS 1b - THE ASSEMBLY ROWS OF THE SHEET
' Mass / volume / area of an assembly are the sum over its content, so they are
' added up from the part values that pass 1 has just read: CATIA does not have
' to evaluate the whole subtree a second time. Length / width / height need
' CATIA's inertia of the assembly - only for assemblies up to
' ASSY_DIMS_MAX_PARTS part instances.
' ==============================================================================
Private Sub FillAssemblyRows()
    Dim it As Variant, r As Long, pn As String
    Dim v As Variant, cacheArr As Variant
    Dim dims(2) As Double
    Dim oProd As Product, t As Double

    For Each it In mAssyRows
        r = it(0)
        pn = CStr(it(1))
        Set oProd = it(2)

        v = RollUp(pn, 0)
        If ASSEMBLY_METRICS Then
            Call PutNumber(r, 6, v(0))
            Call PutNumber(r, 7, v(1))
            Call PutNumber(r, 8, v(2))
        Else
            Call BufSet(r, 6, "N/A")
            Call BufSet(r, 7, "N/A")
            Call BufSet(r, 8, "N/A")
        End If

        dims(0) = 0: dims(1) = 0: dims(2) = 0
        If mDimCache.Exists(pn) Then
            cacheArr = mDimCache(pn)
            dims(0) = CDbl(cacheArr(0)): dims(1) = CDbl(cacheArr(1)): dims(2) = CDbl(cacheArr(2))
        ElseIf ASSY_DIMS_MAX_PARTS > 0 And v(3) > 0 And v(3) <= ASSY_DIMS_MAX_PARTS And v(5) = 0 Then
            ' (only when the whole content is known - otherwise the inertia
            '  would run over parts nobody has counted)
            Call SetStatus("BOM " & mSheetName & ": size of the assembly " & pn)
            Call LogLine(mSheetName & "  assembly size " & pn & "  (" & v(3) & " parts)")
            t = Timer
            On Error Resume Next
            Call GetInertiaDims(oProd.ReferenceProduct, dims)
            Err.Clear
            On Error GoTo 0
            If Elapsed(t) >= SLOW_STEP_SECONDS Then
                Call NoteSlow(Elapsed(t), mSheetName & "  assembly size " & pn & "  (" & v(3) & " parts)")
            End If
            If dims(0) > 0 Then mDimCache.Add pn, Array(dims(0), dims(1), dims(2))
        End If
        Call PutNumber(r, 9, dims(0))
        Call PutNumber(r, 10, dims(1))
        Call PutNumber(r, 11, dims(2))

        If v(5) > 0 Then
            Call BufSet(r, 12, "sum incomplete - content missing")
        Else
            Call BufSet(r, 12, "N/A")
        End If
        Call BufSet(r, 13, "N/A")
        Set oProd = Nothing
        DoEvents
    Next it
End Sub


' ==============================================================================
' HELPER: MASS / VOLUME / AREA OF ONE PIECE OF AN ASSEMBLY, ADDED UP
' Returns Array(mass, volume, area, part instances, part instances without
' values, pieces of unknown content) over every part instance below it. Unknown
' content = a sub-assembly whose sheet failed, or one CATIA has not loaded. Pure bookkeeping - nothing is asked from
' CATIA. The content of every assembly part number is noted while the trees are
' walked (mAssyKids); the part values are the ones pass 1 has read (mPropCache).
' ==============================================================================
Private Function RollUp(ByVal pn As String, ByVal depth As Long) As Variant
    Dim kids As Collection, k As Variant, v As Variant, res As Variant
    Dim m As Double, vo As Double, a As Double
    Dim np As Long, nm As Long, nu As Long

    If mRollCache.Exists(pn) Then
        RollUp = mRollCache(pn)
        Exit Function
    End If
    If depth > 60 Or Not mAssyKids.Exists(pn) Then
        RollUp = Array(0#, 0#, 0#, 0&, 0&, 1&)       ' content unknown
        Exit Function
    End If

    Set kids = mAssyKids(pn)
    For Each k In kids
        If k(1) Then
            v = RollUp(CStr(k(0)), depth + 1)
            m = m + v(0): vo = vo + v(1): a = a + v(2)
            np = np + v(3): nm = nm + v(4): nu = nu + v(5)
        ElseIf mNotLoaded.Exists(CStr(k(0))) Then
            nu = nu + 1                                ' a sub-assembly CATIA has not loaded
        Else
            np = np + 1
            If mPropCache.Exists(CStr(k(0))) Then
                v = mPropCache(CStr(k(0)))
                m = m + v(0): vo = vo + v(1): a = a + v(2)
            Else
                nm = nm + 1
            End If
        End If
    Next k

    res = Array(m, vo, a, np, nm, nu)
    ' a complete sum never changes again; an incomplete one may, later in the run
    If nm = 0 And nu = 0 Then mRollCache.Add pn, res
    RollUp = res
End Function

' the row name of a component whose part number CATIA cannot give yet (not
' loaded): the name of its file ("1DP000524104.CATPart" -> "1DP000524104"),
' else its instance name ("Product1.4") - generic instance names repeat, so
' the number stays, and different components are never merged. Pass 1 loads
' the part and writes the real part number and description.
Private Function ProvisionalName(oProd As Product) As String
    Dim nm As String, pth As String, k As Long

    nm = ""
    On Error Resume Next
    pth = ""
    pth = CStr(oProd.GetMasterShapeRepresentationPathName)
    Err.Clear
    If Len(pth) > 0 Then
        nm = Mid$(pth, InStrRev(Replace(pth, "/", "\"), "\") + 1)
        k = InStrRev(nm, ".")
        If k > 1 Then nm = Left$(nm, k - 1)
    End If
    If Len(nm) = 0 Then
        nm = oProd.Name
        Err.Clear
    End If
    Err.Clear
    On Error GoTo 0
    If Len(nm) = 0 Then nm = "?"
    If Not mProvisional Is Nothing Then
        If Not mProvisional.Exists(nm) Then mProvisional.Add nm, True
    End If
    ProvisionalName = nm
End Function

' something says this component IS a part (and not a sub-assembly): its CATPart
' is in the session, it has a shape, or its file is a CATPart / cgr. Nothing
' is loaded to find out.
Private Function HasPartEvidence(oProd As Product) As Boolean
    Dim tn As String, pth As String, hasShape As Boolean

    HasPartEvidence = False
    On Error Resume Next
    tn = ""
    tn = TypeName(oProd.ReferenceProduct.Parent)
    Err.Clear
    If tn = "PartDocument" Then
        HasPartEvidence = True
        Exit Function
    End If
    hasShape = False
    hasShape = oProd.HasAMasterShapeRepresentation
    If Err.Number = 0 And hasShape Then
        HasPartEvidence = True
        Exit Function
    End If
    Err.Clear
    pth = ""
    pth = LCase$(CStr(oProd.GetMasterShapeRepresentationPathName))
    Err.Clear
    If Right$(pth, 8) = ".catpart" Or Right$(pth, 4) = ".cgr" Then HasPartEvidence = True
    Err.Clear
End Function

' loads the STRUCTURE of a component CATIA has not loaded yet (Teamcenter loads
' on demand) - visualization mode, no geometry in design mode - so its content
' appears like in the CATIA tree. pn and nk are read again afterwards.
Private Sub LoadStructure(oProd As Product, ByRef pn As String, ByRef nk As Long)
    Dim t As Double, nm As String

    nm = ""
    On Error Resume Next
    nm = oProd.Name
    Err.Clear
    Call SetStatus("BOM: loading " & nm & " ...")
    Call LogLine("loading the structure of " & nm)
    t = Timer
    Err.Clear
    oProd.ApplyWorkMode 1
    If Err.Number <> 0 Then Call LogLine("  CATIA refused to load it: error " & Err.Number & " - " & Err.Description)
    Err.Clear
    nk = 0
    nk = oProd.Products.Count
    Err.Clear
    If Len(pn) = 0 Then
        pn = oProd.PartNumber
        Err.Clear
    End If
    On Error GoTo 0
    If Elapsed(t) >= SLOW_STEP_SECONDS Then Call NoteSlow(Elapsed(t), "loading the structure of " & nm)
    DoEvents
End Sub


' ==============================================================================
' HELPER: IS A CHILDLESS NODE A SUB-ASSEMBLY THAT IS NOT LOADED (OR EMPTY)?
' Decided WITHOUT loading anything - every ApplyWorkMode on such a node would
' load its whole content in one blocking call:
'   its CATPart is in the session                         -> a part
'   it has a shape representation                         -> a part
'   its shape file is a CATProduct                        -> NOT a part
'   it has a shape file of another kind (CATPart, cgr)    -> a part
'   its reference is the root of a CATProduct of its own  -> NOT a part
'   no evidence either way                                -> a part (as before)
' ==============================================================================
Private Function IsUnloadedAssembly(oProd As Product) As Boolean
    Dim oDoc As Object, tn As String, pth As String
    Dim docPn As String, myPn As String, hasShape As Boolean

    IsUnloadedAssembly = False
    On Error Resume Next

    Set oDoc = Nothing
    Set oDoc = oProd.ReferenceProduct.Parent
    Err.Clear
    tn = ""
    If Not oDoc Is Nothing Then tn = TypeName(oDoc)
    If tn = "PartDocument" Then Exit Function

    ' a node with a shape of its own (also a CATProduct whose geometry is a
    ' CATPart "shape representation") is a part
    hasShape = False
    hasShape = oProd.HasAMasterShapeRepresentation
    If Err.Number = 0 And hasShape Then Exit Function
    Err.Clear

    pth = ""
    pth = LCase$(CStr(oProd.GetMasterShapeRepresentationPathName))
    Err.Clear
    If Right$(pth, 11) = ".catproduct" Then
        IsUnloadedAssembly = True
        Exit Function
    End If
    If Len(pth) > 0 Then Exit Function

    ' its reference is a CATProduct document whose root IS this node (and not
    ' the assembly the node sits in) - a sub-assembly without children
    If tn = "ProductDocument" Then
        docPn = ""
        docPn = oDoc.Product.PartNumber
        myPn = ""
        myPn = oProd.PartNumber
        Err.Clear
        If Len(myPn) > 0 And StrComp(docPn, myPn, vbTextCompare) = 0 Then IsUnloadedAssembly = True
    End If
    Err.Clear
End Function

' notes the content of an assembly part number once (all its instances share it)
Private Sub NoteAssemblyContent(ByVal pn As String, kids As Collection)
    If Len(pn) = 0 Then Exit Sub
    If mAssyKids Is Nothing Then Exit Sub
    If Not mAssyKids.Exists(pn) Then mAssyKids.Add pn, kids
End Sub



' ==============================================================================
' PASS 2 - ALL PICTURES OF ONE SHEET (after all its numbers are written)
'   1. the component as it is - the picture of its own row
'   2. every part, and every sub-assembly that is a CATProduct of its own, in
'      a WINDOW OF ITS OWN: it is alone there, nothing has to be hidden. A part
'      number is pictured once for the whole run, every further row reuses it.
'   3. the rest (a part CATIA has only as cgr, an assembly inside another file)
'      the old way: hide every part, show one, capture. If hiding does not
'      change the view in this window, those rows get "No Preview" instead of
'      a picture of the whole component.
' Every picture goes into Excel straight away.
' ==============================================================================
Private Sub TakePictures(sel As Selection, oViewer As Viewer, xlSheet As Object)
    Dim j As Long, i As Long, rowNo As Long
    Dim kind As String, key As String, bodyName As String, pn As String
    Dim picPath As String, what As String, shotKey As String
    Dim fullPath As String, basePath As String
    Dim fullSig As String, baseSig As String, sig As String
    Dim oProd As Product, colNode As Collection
    Dim oPart As Part, oPartDoc As PartDocument
    Dim oDoc As Object, homeWin As Object, homeDoc As Document
    Dim bodyNames As Variant
    Dim colWasShown As Collection, colWasHidden As Collection
    Dim colFb As Collection
    Dim t As Double, sceneOK As Boolean
    Dim partSel As Selection, oBody As Body

    If mJobRow.Count = 0 Then Exit Sub

    If oViewer Is Nothing Then
        For j = 1 To mJobRow.Count
            Call XlPutCell(xlSheet, mJobRow(j), 1, "No Viewer")
        Next j
        Exit Sub
    End If

    Call LogLine(mSheetName & ": " & mJobRow.Count & " pictures")
    On Error Resume Next
    Set homeWin = Nothing
    Set homeWin = CATIA.ActiveWindow
    Set homeDoc = Nothing
    Set homeDoc = CATIA.ActiveDocument
    Err.Clear
    On Error GoTo 0
    Set colFb = New Collection

    ' --- 1. the component as it is ---
    fullPath = mTempDir & "\s" & mSheetNo & "_full.jpg"
    fullSig = ""
    If CaptureViewToFile(oViewer, fullPath) Then fullSig = FileSignature(fullPath)

    ' --- 2. windows of their own ---
    For j = 1 To mJobRow.Count
        Call CheckExcel                       ' no hour of pictures into a closed Excel
        rowNo = mJobRow(j)
        kind = mJobKind(j)
        key = mJobKey(j)
        bodyName = mJobBody(j)
        pn = ""
        If Len(key) > 0 Then
            If mPartNum.Exists(key) Then pn = mPartNum(key)
        End If
        what = kind
        If Len(pn) > 0 Then what = kind & " " & pn
        If Len(bodyName) > 0 Then what = what & " / " & bodyName
        Call SetStatus("BOM " & mSheetName & ": picture " & j & " of " & mJobRow.Count & " - " & what)
        Call LogLine(mSheetName & "  picture " & j & "/" & mJobRow.Count & "  " & what)
        t = Timer
        picPath = ""

        Select Case kind
            Case "UNIT"
                If Len(fullSig) > 0 Then picPath = fullPath

            Case "PART"
                shotKey = "P|" & pn
                picPath = CachedShot(shotKey)
                If Len(picPath) = 0 And PICTURES_IN_OWN_WINDOW Then
                    Set oProd = mRef(key)
                    If PartDocOf(oProd, pn, oPart, oPartDoc) Then
                        picPath = ShotInOwnWindow(oPartDoc, homeWin, shotKey, Nothing, Empty, "", "", fullSig)
                    End If
                End If

            Case "BODY"
                shotKey = "B|" & pn & "|" & bodyName
                picPath = CachedShot(shotKey)
                If Len(picPath) = 0 And mBodies.Exists(pn) And PICTURES_IN_OWN_WINDOW Then
                    Set oProd = mRef(key)
                    If PartDocOf(oProd, pn, oPart, oPartDoc) Then
                        bodyNames = mBodies(pn)
                        picPath = ShotInOwnWindow(oPartDoc, homeWin, shotKey, oPart, bodyNames, _
                                                  bodyName, "P|" & pn, fullSig)
                    End If
                End If

            Case "ASSY"
                shotKey = "A|" & pn
                picPath = CachedShot(shotKey)
                If Len(picPath) = 0 And PICTURES_IN_OWN_WINDOW And Not homeDoc Is Nothing Then
                    Set oProd = mRef(key)
                    Set oDoc = OwnDocumentOf(oProd, pn, homeDoc)
                    If Not oDoc Is Nothing Then
                        picPath = ShotInOwnWindow(oDoc, homeWin, shotKey, Nothing, Empty, "", "", fullSig)
                    End If
                End If
        End Select

        If Len(picPath) > 0 Then
            If Not XlInsertPicture(xlSheet, rowNo, picPath) Then Call XlPutCell(xlSheet, rowNo, 1, "No Preview")
        ElseIf kind = "PART" Or kind = "ASSY" Or kind = "BODY" Then
            colFb.Add j                       ' below, by show / hide in this window
        Else
            Call XlPutCell(xlSheet, rowNo, 1, "No Preview")
        End If

        If Elapsed(t) >= SLOW_STEP_SECONDS Then
            Call NoteSlow(Elapsed(t), mSheetName & "  picture of row " & rowNo & "  " & what)
        End If
        Call SaveWorkbook(False)                  ' at most every SAVE_EVERY_SECONDS
        DoEvents
    Next j

    ' --- 3. the rest by show / hide ---
    If colFb.Count > 0 Then
        Call LogLine(mSheetName & ": " & colFb.Count & " pictures by show / hide")
        If PRESERVE_VISIBILITY Then
            Set colWasShown = New Collection
            Set colWasHidden = New Collection
            Call SplitByShowState(sel, mAllNodes, colWasShown, colWasHidden)
        End If
        ' every assembly visible (a hidden one hides its whole subtree), every
        ' part hidden
        Call SetShowCollection(sel, mAssyNodes, 0)
        Call SetShowCollection(sel, mLeaves, 1)
        Call Settle(100)

        ' with every part hidden the view must differ from the whole component -
        ' if it does not, hiding has no effect in this window and every picture
        ' would show the whole component
        sceneOK = True
        baseSig = ""
        basePath = mTempDir & "\s" & mSheetNo & "_base.jpg"
        If CaptureViewToFile(oViewer, basePath) Then baseSig = FileSignature(basePath)
        If Len(fullSig) > 0 And baseSig = fullSig Then
            sceneOK = False
            Call LogLine(mSheetName & ": hiding the parts does not change the view - " & _
                         colFb.Count & " pictures left out (No Preview)")
        End If

        For i = 1 To colFb.Count
            Call CheckExcel
            j = colFb(i)
            rowNo = mJobRow(j)
            kind = mJobKind(j)
            key = mJobKey(j)
            bodyName = mJobBody(j)
            pn = ""
            If mPartNum.Exists(key) Then pn = mPartNum(key)
            If kind = "ASSY" Then
                shotKey = "A|" & pn
            ElseIf kind = "BODY" Then
                shotKey = "B|" & pn & "|" & bodyName
            Else
                shotKey = "P|" & pn
            End If
            picPath = CachedShot(shotKey)       ' an earlier row may have it by now
            t = Timer

            If sceneOK And Len(picPath) = 0 Then
                Call SetStatus("BOM " & mSheetName & ": picture (show / hide) " & i & " of " & colFb.Count)
                Call LogLine(mSheetName & "  picture (show / hide) " & i & "/" & colFb.Count & "  " & kind & " " & pn)
                mShotNo = mShotNo + 1
                picPath = mTempDir & "\shot" & mShotNo & ".jpg"
                If kind = "ASSY" Then
                    Set colNode = Nothing
                    If mNodeLeaves.Exists(key) Then Set colNode = mNodeLeaves(key)
                    If Not colNode Is Nothing Then
                        If colNode.Count > 0 And colNode.Count <= MAX_ASSY_LEAVES_FOR_SHOT Then
                            Call SetShowCollection(sel, colNode, 0)
                            Call Settle(ASSY_SETTLE_MS)
                            Call CaptureViewToFile(oViewer, picPath)
                            Call SetShowCollection(sel, colNode, 1)
                        End If
                    End If
                ElseIf kind = "BODY" Then
                    ' the part shown, all its bodies hidden, then this one body
                    Set oProd = mRef(key)
                    If mBodies.Exists(pn) Then
                        If PartDocOf(oProd, pn, oPart, oPartDoc) Then
                            bodyNames = mBodies(pn)
                            Set partSel = oPartDoc.Selection
                            Call SetShowSingle(sel, oProd, 0)
                            Call SetShowBodyList(oPart, partSel, bodyNames, 1)
                            Set oBody = GetBodyByName(oPart, bodyName)
                            On Error Resume Next
                            partSel.Clear
                            partSel.Add oBody
                            partSel.VisProperties.SetShow 0
                            partSel.Clear
                            Err.Clear
                            On Error GoTo 0
                            Call Settle(PART_SETTLE_MS)
                            Call CaptureViewToFile(oViewer, picPath)
                            Call SetShowBodyList(oPart, partSel, bodyNames, 0)
                            Call SetShowSingle(sel, oProd, 1)
                        End If
                    End If
                Else
                    Set oProd = mRef(key)
                    Call SetShowSingle(sel, oProd, 0)
                    Call Settle(PART_SETTLE_MS)
                    Call CaptureViewToFile(oViewer, picPath)
                    Call SetShowSingle(sel, oProd, 1)
                End If

                If mFso.FileExists(picPath) Then
                    ' nothing visible: no picture of THIS row. (Equal to the whole
                    ' component is fine here - hiding works, so the rest of the
                    ' component simply draws nothing, e.g. parts not loaded.)
                    sig = FileSignature(picPath)
                    ' a body equal to its whole part: hiding the other bodies
                    ' had no effect
                    If kind = "BODY" And mShotSigs.Exists("P|" & pn) Then
                        If mShotSigs("P|" & pn) = sig Then sig = baseSig
                    End If
                    If sig = baseSig Then
                        On Error Resume Next
                        mFso.DeleteFile picPath, True
                        Err.Clear
                        On Error GoTo 0
                        picPath = ""
                    Else
                        If Not mShotFiles.Exists(shotKey) Then
                            mShotFiles.Add shotKey, picPath
                            mShotSigs.Add shotKey, sig
                        End If
                    End If
                Else
                    picPath = ""
                End If
            End If

            If Len(picPath) > 0 Then
                If Not XlInsertPicture(xlSheet, rowNo, picPath) Then Call XlPutCell(xlSheet, rowNo, 1, "No Preview")
            Else
                Call XlPutCell(xlSheet, rowNo, 1, "No Preview")
            End If
            If Elapsed(t) >= SLOW_STEP_SECONDS Then
                Call NoteSlow(Elapsed(t), mSheetName & "  picture of row " & rowNo & "  " & kind & " " & pn)
            End If
            Call SaveWorkbook(False)
            DoEvents
        Next i

        ' --- back to normal ---
        Call RestoreVisibility(sel, mLeaves, colWasShown, colWasHidden)
    End If

    On Error Resume Next
    If mFso.FileExists(fullPath) Then mFso.DeleteFile fullPath, True
    If Len(basePath) > 0 Then
        If mFso.FileExists(basePath) Then mFso.DeleteFile basePath, True
    End If
    CATIA.StatusBar = ""
    Err.Clear
    On Error GoTo 0
End Sub


' ==============================================================================
' HELPER: ONE PICTURE IN A WINDOW OF ITS OWN
' The document (a CATPart, or a sub-assembly's CATProduct) is opened in a new
' window - it is alone there, so nothing in the product has to be hidden.
' For a body picture the other bodies of the part are hidden in that window.
' refKey: the picture a body picture must differ from (the whole part).
' Returns the picture file, or "" when there is none.
' ==============================================================================
Private Function ShotInOwnWindow(oDoc As Object, homeWin As Object, ByVal shotKey As String, _
                                 oPart As Part, bodyNames As Variant, ByVal bodyName As String, _
                                 ByVal refKey As String, ByVal homeSig As String) As String
    Dim w As Object, v As Viewer, partSel As Selection, oBody As Body
    Dim picPath As String, sig As String, ok As Boolean, isolated As Boolean

    ShotInOwnWindow = ""
    mShotNo = mShotNo + 1
    picPath = mTempDir & "\shot" & mShotNo & ".jpg"

    On Error Resume Next
    Set w = Nothing
    Set w = oDoc.NewWindow
    If w Is Nothing Then
        Err.Clear
        Exit Function
    End If
    w.Activate
    Set v = Nothing
    Set v = w.ActiveViewer
    w.Layout = catWindowGeomOnly             ' no specification tree in the picture
    ' (no compass command here: when the window is not really displayed, the
    '  command lands in the main window and changes every later picture)
    Err.Clear

    If Len(bodyName) > 0 And Not oPart Is Nothing Then
        Set partSel = oDoc.Selection
        Call SetShowBodyList(oPart, partSel, bodyNames, 1)
        Set oBody = GetBodyByName(oPart, bodyName)
        partSel.Clear
        partSel.Add oBody
        partSel.VisProperties.SetShow 0
        partSel.Clear
        isolated = True
        Err.Clear
    End If

    If Not v Is Nothing Then
        If SET_ISO_VIEW Then Call SetIsoViewpoint(v)
        Call Settle(ASSY_SETTLE_MS)
        ok = CaptureViewToFile(v, picPath)
    End If

    If isolated Then Call SetShowBodyList(oPart, partSel, bodyNames, 0)
    w.Close
    Err.Clear
    If Not homeWin Is Nothing Then homeWin.Activate
    Err.Clear
    On Error GoTo 0
    If Not ok Then Exit Function

    sig = FileSignature(picPath)
    ' the same as the main window: the new window was not displayed and CATIA
    ' captured the main window - this row is pictured there instead
    If Len(homeSig) > 0 And sig = homeSig Then
        Call LogLine(mSheetName & "  the window of " & shotKey & " was not displayed - main window instead")
        On Error Resume Next
        mFso.DeleteFile picPath, True
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    ' a body picture equal to the picture of the whole part: hiding the other
    ' bodies had no effect - no picture rather than a wrong one
    If Len(refKey) > 0 Then
        If mShotSigs.Exists(refKey) Then
            If mShotSigs(refKey) = sig Then
                On Error Resume Next
                mFso.DeleteFile picPath, True
                Err.Clear
                On Error GoTo 0
                Exit Function
            End If
        End If
    End If

    If Not mShotFiles.Exists(shotKey) Then
        mShotFiles.Add shotKey, picPath
        mShotSigs.Add shotKey, sig
    End If
    ShotInOwnWindow = picPath
End Function

' the picture file of a part / body / assembly that was already taken in this run
Private Function CachedShot(ByVal shotKey As String) As String
    CachedShot = ""
    If mShotFiles Is Nothing Then Exit Function
    If Not mShotFiles.Exists(shotKey) Then Exit Function
    If mFso.FileExists(CStr(mShotFiles(shotKey))) Then CachedShot = CStr(mShotFiles(shotKey))
End Function

' the CATPart of a picture row: the one pass 1 found, else found again
Private Function PartDocOf(oProd As Product, ByVal pn As String, ByRef oPart As Part, _
                           ByRef oPartDoc As PartDocument) As Boolean
    Dim ok As Boolean

    ok = False
    Set oPart = Nothing
    Set oPartDoc = Nothing
    If Not mPartDocs Is Nothing Then
        If mPartDocs.Exists(pn) Then
            On Error Resume Next
            Set oPartDoc = mPartDocs(pn)
            Set oPart = oPartDoc.Part
            ok = (Err.Number = 0) And Not (oPart Is Nothing)
            Err.Clear
            On Error GoTo 0
        End If
    End If
    If Not ok Then ok = LoadPartDocument(oProd, oPart, oPartDoc)
    PartDocOf = ok
End Function

' a fingerprint of a picture file: two captures of an unchanged view are the
' same file, byte for byte
Private Function FileSignature(ByVal path As String) As String
    Dim f As Integer, n As Long, i As Long
    Dim b() As Byte, h1 As Double, h2 As Double

    FileSignature = ""
    On Error Resume Next
    f = FreeFile
    Err.Clear
    Open path For Binary Access Read As #f
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    n = LOF(f)
    If n > 0 Then
        ReDim b(0 To n - 1)
        Get #f, 1, b
        For i = 0 To n - 1
            h1 = h1 * 31# + b(i)
            h1 = h1 - Int(h1 / 2147483647#) * 2147483647#
            h2 = h2 * 37# + b(i)
            h2 = h2 - Int(h2 / 2147483629#) * 2147483629#
        Next i
    End If
    Close #f
    Err.Clear
    FileSignature = n & ":" & h1 & ":" & h2
End Function


' ==============================================================================
' HELPER: REFRAME + CAPTURE THE CURRENT VIEW INTO A FILE
' ==============================================================================
Private Function CaptureViewToFile(oViewer As Viewer, ByVal picPath As String) As Boolean
    On Error Resume Next
    CaptureViewToFile = False
    If oViewer Is Nothing Then Exit Function

    oViewer.Reframe
    oViewer.Update
    Call Settle(PART_SETTLE_MS)

    If mFso.FileExists(picPath) Then mFso.DeleteFile picPath
    ' 5 = JPEG (small file, fast to insert); 4 = BMP as fallback for releases
    ' that refuse JPEG - a BMP of a full viewer is several MB
    oViewer.CaptureToFile 5, picPath
    If Not mFso.FileExists(picPath) Then
        Err.Clear
        oViewer.CaptureToFile 4, picPath
    End If
    CaptureViewToFile = mFso.FileExists(picPath)
    Err.Clear
End Function


' ==============================================================================
' HELPER: NOTE A PICTURE THAT PASS 2 HAS TO TAKE
' ==============================================================================
Private Sub AddJob(ByVal rowNo As Long, ByVal kind As String, ByVal key As String, _
                   ByVal bodyName As String)
    mJobRow.Add rowNo
    mJobKind.Add kind
    mJobKey.Add key
    mJobBody.Add bodyName
End Sub


' ==============================================================================
' HELPER: FRESH ROW DATA FOR A NEW SHEET
' ==============================================================================
Private Sub ResetSheetData()
    Set mQty = CreateObject("Scripting.Dictionary")
    Set mRef = CreateObject("Scripting.Dictionary")
    Set mDesc = CreateObject("Scripting.Dictionary")
    Set mProps = CreateObject("Scripting.Dictionary")
    Set mLevel = CreateObject("Scripting.Dictionary")
    Set mBodies = CreateObject("Scripting.Dictionary")
    Set mPartNum = CreateObject("Scripting.Dictionary")
    Set mFirst = CreateObject("Scripting.Dictionary")
    Set mIsAssy = CreateObject("Scripting.Dictionary")
    Set mNodeLeaves = CreateObject("Scripting.Dictionary")
    Set mLeaves = New Collection
    Set mAllNodes = New Collection
    Set mAssyNodes = New Collection
    Set mJobRow = New Collection
    Set mJobKind = New Collection
    Set mJobKey = New Collection
    Set mJobBody = New Collection
    Set mBoldRows = New Collection
    Set mAssyRows = New Collection
    ReDim mBuf(1 To 16, 1 To 256)
    mBufRows = 0
    mFlushed = False
    Set mLostBlocks = CreateObject("Scripting.Dictionary")
    mSheetNo = mSheetNo + 1
End Sub


' ==============================================================================
' HELPER: THE ROW BUFFER OF THE SHEET BEING WRITTEN
' The rows are collected in memory and go into Excel in a few big blocks
' (FlushSheet) instead of one cross-process call per cell.
' ==============================================================================
Private Sub BufSet(ByVal r As Long, ByVal c As Long, ByVal v As Variant)
    Dim i As Long, n As Long
    i = r - 1                                   ' sheet row 2 = buffer row 1
    If i < 1 Or c < 1 Or c > 16 Then Exit Sub
    n = UBound(mBuf, 2)
    If i > n Then
        Do While n < i
            n = n * 2
        Loop
        ReDim Preserve mBuf(1 To 16, 1 To n)
    End If
    mBuf(c, i) = v
    If i > mBufRows Then mBufRows = i
End Sub

' a value > 0, otherwise "N/A"
Private Sub PutNumber(ByVal r As Long, ByVal c As Long, ByVal v As Variant)
    If IsNumeric(v) Then
        If CDbl(v) > 0 Then
            Call BufSet(r, c, CDbl(v))
            Exit Sub
        End If
    End If
    Call BufSet(r, c, "N/A")
End Sub

' writes the buffer to the sheet (rows 2 ...) in blocks of 500 rows
' onlyLost = True writes only the blocks an earlier flush could not write.
' Returns the number of rows Excel did not take.
Private Function FlushSheet(xlSheet As Object, ByVal onlyLost As Boolean) As Long
    Const BLOCK_ROWS As Long = 500
    Dim arr() As Variant, rowArr() As Variant
    Dim i0 As Long, n As Long, i As Long, c As Variant, k As Long
    Dim rr As Variant, nLost As Long, nBlockLost As Long, v As Variant

    mFlushed = True
    FlushSheet = 0
    If mBufRows = 0 Then Exit Function

    Call SetStatus("BOM " & mSheetName & ": writing " & mBufRows & " rows to Excel ...")
    For i0 = 1 To mBufRows Step BLOCK_ROWS
        If onlyLost And Not mLostBlocks.Exists(i0) Then GoTo NextBlock
        n = mBufRows - i0 + 1
        If n > BLOCK_ROWS Then n = BLOCK_ROWS
        ReDim arr(1 To n, 1 To 16)
        For i = 1 To n
            For k = 1 To 16
                arr(i, k) = mBuf(k, i0 + i - 1)
            Next k
        Next i
        nBlockLost = 0
        If onlyLost And Not mXlLost Then
            ' its long-text cells may have been set to General after the first
            ' try - text again, so nothing is read as a formula
            Call XlReady
            On Error Resume Next
            For Each c In Array(4, 14, 16)
                xlSheet.Range(xlSheet.Cells(i0 + 1, c), xlSheet.Cells(i0 + n, c)).NumberFormat = "@"
            Next c
            Err.Clear
            On Error GoTo 0
        End If
        If Not XlWriteBlock(xlSheet, i0 + 1, 1, arr) Then
            If mXlLost Or mXlLastErr = 0 Then
                nBlockLost = n                       ' Excel gone, or the user skipped
            Else
                ' Excel refused a VALUE: row by row, so it costs one row, not 500
                Call LogLine(mSheetName & ": block from row " & (i0 + 1) & " refused - writing it row by row")
                ReDim rowArr(1 To 1, 1 To 16)
                For i = 1 To n
                    For k = 1 To 16
                        rowArr(1, k) = mBuf(k, i0 + i - 1)
                    Next k
                    If Not XlWriteBlock(xlSheet, i0 + i, 1, rowArr) Then nBlockLost = nBlockLost + 1
                Next i
            End If
        End If
        If nBlockLost > 0 Then
            If Not mLostBlocks.Exists(i0) Then mLostBlocks.Add i0, True
        ElseIf mLostBlocks.Exists(i0) Then
            mLostBlocks.Remove i0
        End If
        nLost = nLost + nBlockLost
NextBlock:
    Next i0

    If Not mXlLost Then Call ShowAsTree(xlSheet)

    If Not mXlLost Then
        Call XlReady
        On Error Resume Next
        ' assembly rows in bold - only looks, so a failure here is ignored
        For Each rr In mBoldRows
            xlSheet.Range(xlSheet.Cells(rr, 2), xlSheet.Cells(rr, 16)).Font.Bold = True
        Next rr
        ' a text cell shows more than 255 characters only as #####: back to
        ' General (the value stays the text it is)
        For i = 1 To mBufRows
            For Each c In Array(4, 14, 16)
                v = mBuf(c, i)
                If VarType(v) = vbString Then
                    If Len(v) > 255 Then xlSheet.Cells(i + 1, c).NumberFormat = "General"
                End If
            Next c
        Next i
        Err.Clear
        On Error GoTo 0
    End If
    FlushSheet = nLost
End Function

' the sheet as a tree: every row grouped below its parent row with Excel's
' outline (the +/- at the left, like the CATIA tree; Excel has up to 8 group
' levels) and the column "CATIA tree" indented by depth - in the assembly
' structure, where the parent rows exist. Only looks - a failure here is
' ignored.
Private Sub ShowAsTree(xlSheet As Object)
    Dim i As Long, j As Long, lv As Long, lv2 As Long, minLv As Long
    Dim ol As Long, ind As Long

    If mBufRows = 0 Then Exit Sub

    ' the part list has no assembly rows - no tree to show; only the bodies
    ' fold below their part
    If Not (mIncludeSubassemblies And UCase$(mGroupBy) = "PARENT") Then
        Call XlReady
        On Error Resume Next
        xlSheet.Outline.SummaryRow = 0
        For i = 1 To mBufRows
            If CStr(mBuf(15, i)) = "Body" Then
                xlSheet.Rows(i + 1).OutlineLevel = 2
                xlSheet.Cells(i + 1, 16).IndentLevel = 1
            End If
        Next i
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    minLv = 0
    For i = 1 To mBufRows
        If IsNumeric(mBuf(2, i)) And Not IsEmpty(mBuf(2, i)) Then
            lv = CLng(mBuf(2, i))
            If minLv = 0 Or lv < minLv Then minLv = lv
        End If
    Next i
    If minLv = 0 Then minLv = 1

    Call XlReady
    On Error Resume Next
    xlSheet.Outline.SummaryRow = 0          ' xlSummaryAbove: the parent row above its content
    i = 1
    Do While i <= mBufRows
        lv = RowDepth(i, minLv)
        j = i
        Do While j < mBufRows
            lv2 = RowDepth(j + 1, minLv)
            If lv2 <> lv Then Exit Do
            j = j + 1
        Loop
        ol = lv
        If ol > 8 Then ol = 8
        If ol > 1 Then xlSheet.Rows((i + 1) & ":" & (j + 1)).OutlineLevel = ol
        ind = lv - 1
        If ind > 15 Then ind = 15
        If ind > 0 Then xlSheet.Range(xlSheet.Cells(i + 1, 16), xlSheet.Cells(j + 1, 16)).IndentLevel = ind
        i = j + 1
        If i Mod 200 = 0 Then DoEvents
    Loop
    Err.Clear
    On Error GoTo 0
End Sub

' depth of buffer row i in the tree of the sheet (1 = the top rows)
Private Function RowDepth(ByVal i As Long, ByVal minLv As Long) As Long
    Dim d As Long
    d = 1
    If IsNumeric(mBuf(2, i)) And Not IsEmpty(mBuf(2, i)) Then d = CLng(mBuf(2, i)) - minLv + 1
    If d < 1 Then d = 1
    RowDepth = d
End Function

' a header or rows Excel did not take: one more try before the pictures (the
' user may have freed Excel meanwhile). The whole buffer is written again -
' the same values, so nothing is doubled.
Private Sub RetryLostWrites(xlSheet As Object, ByRef hdrOK As Boolean, ByRef nLost As Long)
    If hdrOK And nLost = 0 Then Exit Sub
    If mXlLost Then Exit Sub
    Call LogLine(mSheetName & ": writing again what Excel did not take")
    mXlSkipAt = -1
    If Not hdrOK Then
        hdrOK = WriteHeader(xlSheet)
        If hdrOK Then
            ' the text formats came late: every row again, so part numbers
            ' that were written as numbers become text (same values)
            nLost = FlushSheet(xlSheet, False)
        ElseIf nLost > 0 Then
            ' header / formats still failing: at least the lost rows
            nLost = FlushSheet(xlSheet, True)
        End If
    ElseIf nLost > 0 Then
        nLost = FlushSheet(xlSheet, True)
    End If
End Sub

Private Function SheetLostText(ByVal hdrOK As Boolean, ByVal nLost As Long) As String
    Dim s As String
    s = "INCOMPLETE: "
    If Not hdrOK Then s = s & "the header row"
    If Not hdrOK And nLost > 0 Then s = s & " and "
    If nLost > 0 Then s = s & nLost & " rows"
    SheetLostText = s & " could not be written to Excel"
End Function


' ==============================================================================
' HELPER: THE COLUMN HEADERS OF A BOM SHEET
' (identical on every sheet, so every sheet imports into the ACE cost model)
' ==============================================================================
Private Function WriteHeader(xlSheet As Object) As Boolean
    Dim h As Variant, a(1 To 1, 1 To 16) As Variant
    Dim i As Long, fmtOK As Boolean

    h = Array("Thumbnail", "Level", "Part Number", "Description", "Qty", _
              "Mass (kg)", "Volume (m3)", "Area (m2)", "Length (mm)", "Width (mm)", _
              "Height (mm)", "Material", "Density (kg/m3)", "First Level", "Type", "CATIA tree")
    For i = 0 To 15
        a(1, i + 1) = h(i)
    Next i

    ' the formats first: every text column stays TEXT - part numbers keep
    ' leading zeros, "1E5" stays "1E5", and a description like "+X PANEL" or
    ' "-Y WALL" is not taken for a formula (#NAME? or a refused block)
    Call XlReady
    On Error Resume Next
    Err.Clear
    With xlSheet
        .Columns("A:A").NumberFormat = "@"
        .Columns("C:D").NumberFormat = "@"
        .Columns("L:L").NumberFormat = "@"
        .Columns("N:N").NumberFormat = "@"
        .Columns("P:P").NumberFormat = "@"
    End With
    fmtOK = (Err.Number = 0)              ' (an error under Resume Next stays in Err)
    Err.Clear
    ' looks only - a failure here does not count
    With xlSheet
        .Columns("A:A").ColumnWidth = 15
        .Columns("B:B").ColumnWidth = 8
        .Columns("C:D").ColumnWidth = 20
        .Columns("F:H").NumberFormat = "0.000000000"
        .Columns("I:K").NumberFormat = "0.0"
        .Columns("M:M").NumberFormat = "0.000"
        .Columns("N:N").ColumnWidth = 22
        .Columns("O:O").ColumnWidth = 10
        .Range("A1:P1").Font.Bold = True
        .Range("A1:P1").Interior.Color = RGB(220, 220, 220)
    End With
    Err.Clear
    On Error GoTo 0

    ' success = the formats AND the header row: rows written before the text
    ' formats would lose leading zeros, so a late format means a new flush
    WriteHeader = XlWriteBlock(xlSheet, 1, 1, a) And fmtOK
End Function


' ==============================================================================
' HELPER: FINISH A BOM SHEET (column widths + filter) AND SHOW IT
' ==============================================================================
Private Sub FinishSheetLayout(xlSheet As Object)
    Call XlReady
    On Error Resume Next
    xlSheet.Columns("B:P").AutoFit
    xlSheet.Columns("A:A").ColumnWidth = 15
    xlSheet.Columns("N:N").ColumnWidth = 22
    If xlSheet.Columns("P:P").ColumnWidth > 90 Then xlSheet.Columns("P:P").ColumnWidth = 90
    If Not xlSheet.AutoFilterMode Then xlSheet.Range("B1:P1").AutoFilter
    xlSheet.Activate
    ' the header stays visible while scrolling
    mXlApp.ActiveWindow.FreezePanes = False
    mXlApp.ActiveWindow.SplitColumn = 0
    mXlApp.ActiveWindow.SplitRow = 1
    mXlApp.ActiveWindow.FreezePanes = True
    mXlApp.ScreenUpdating = True
    DoEvents
    Err.Clear
End Sub

' the note of a sheet that could not be finished - in column Q, outside the
' columns the ACE import reads
Private Sub MarkSheetIncomplete(xlSheet As Object, ByVal statusTxt As String)
    Call XlPutCell(xlSheet, 1, 17, "!! INCOMPLETE - " & statusTxt)
    Call XlReady
    On Error Resume Next
    xlSheet.Cells(1, 17).Font.Color = RGB(200, 0, 0)
    xlSheet.Tab.Color = RGB(255, 150, 120)
    mXlApp.ScreenUpdating = True
    Err.Clear
End Sub


' ==============================================================================
' HELPER: THE INDEX SHEET
' ==============================================================================
Private Sub WriteIndexHeader(idxSheet As Object, ByVal splitLevel As Long, ByRef blockOK As Boolean)
    Dim a(1 To 2, 1 To 12) As Variant
    Dim it As Variant

    ' 1. the text formats - once. Rows written before them are written again,
    '    so part numbers keep their leading zeros; columns are never set to
    '    text a second time (that would undo the General of long-text cells)
    If Not mIdxFmtOK Then
        Call XlReady
        On Error Resume Next
        Err.Clear
        idxSheet.Columns("B:C").NumberFormat = "@"
        idxSheet.Columns("F:G").NumberFormat = "@"
        mIdxFmtOK = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
        If mIdxFmtOK And mIdxEarly.Count > 0 Then
            Call LogLine("Index formats came late - writing " & mIdxEarly.Count & " rows again")
            For Each it In mIdxEarly
                Call WriteIndexRow(idxSheet, CLng(it(0)), CLng(it(1)), CStr(it(2)), CStr(it(3)), _
                                   CLng(it(4)), CLng(it(5)), CStr(it(6)), CStr(it(7)), CStr(it(8)), False)
            Next it
            Set mIdxEarly = New Collection
        End If
    End If

    ' 2. the title and the column headers
    If Not blockOK Then
        a(1, 1) = "BOM split at level " & splitLevel & _
                  "  -  quantities on a component sheet are per ONE piece of that component; " & _
                  "column E says how many pieces the product contains"
        a(2, 1) = "#"
        a(2, 2) = "Component (as in the CATIA tree)"
        a(2, 3) = "Description"
        a(2, 4) = "Level"
        a(2, 5) = "Qty in product"
        a(2, 6) = "First Level"
        a(2, 7) = "Path"
        a(2, 8) = "Sheet"
        a(2, 9) = "Rows"
        a(2, 10) = "Status"
        a(2, 11) = "Seconds"
        a(2, 12) = "Worked in"
        blockOK = XlWriteBlock(idxSheet, 1, 1, a)
        If blockOK Then
            Call XlReady
            On Error Resume Next
            With idxSheet
                .Cells(1, 1).Font.Italic = True
                .Range("A2:L2").Font.Bold = True
                .Range("A2:L2").Interior.Color = RGB(220, 220, 220)
            End With
            Err.Clear
            On Error GoTo 0
        End If
    End If
End Sub

' withStatus = False (a replay) writes only A:G, so a result that is already
' in I:L is never overwritten by "waiting ..."
Private Function WriteIndexRow(idxSheet As Object, ByVal idxRow As Long, ByVal no As Long, _
                               ByVal pn As String, ByVal desc As String, ByVal lvl As Long, _
                               ByVal qty As Long, ByVal firstLevel As String, ByVal pathTxt As String, _
                               ByVal sheetName As String, ByVal withStatus As Boolean) As Boolean
    Dim a() As Variant
    Dim ok As Boolean, c As Variant

    If withStatus Then
        ReDim a(1 To 1, 1 To 10)
        a(1, 10) = "waiting ..."
    Else
        ReDim a(1 To 1, 1 To 7)
    End If
    If no > 0 Then a(1, 1) = no
    a(1, 2) = pn
    a(1, 3) = desc
    If lvl > 0 Then a(1, 4) = lvl
    a(1, 5) = qty
    a(1, 6) = firstLevel
    a(1, 7) = pathTxt
    ok = XlWriteBlock(idxSheet, idxRow, 1, a)
    WriteIndexRow = ok
    If Not ok Then Exit Function

    ' a text cell shows more than 255 characters only as #####: those cells go
    ' back to General (the value stays the text it is)
    If Len(desc) > 255 Or Len(firstLevel) > 255 Or Len(pathTxt) > 255 Then
        Call XlReady
        On Error Resume Next
        For Each c In Array(3, 6, 7)
            If Len(CStr(a(1, c))) > 255 Then idxSheet.Cells(idxRow, c).NumberFormat = "General"
        Next c
        Err.Clear
        On Error GoTo 0
    End If

    If Len(sheetName) = 0 Then Exit Function
    Call XlReady
    On Error Resume Next
    idxSheet.Hyperlinks.Add idxSheet.Cells(idxRow, 8), "", _
                            "'" & Replace(sheetName, "'", "''") & "'!A1", "", sheetName
    Err.Clear
End Function

Private Function WriteIndexResult(idxSheet As Object, ByVal idxRow As Long, ByVal nRows As Long, _
                                  ByVal statusTxt As String, ByVal secs As Double, _
                                  ByVal whereTxt As String) As Boolean
    Dim a(1 To 1, 1 To 4) As Variant
    Dim ok As Boolean

    a(1, 1) = nRows
    a(1, 2) = statusTxt
    a(1, 3) = Round(secs, 1)
    a(1, 4) = whereTxt
    ok = XlWriteBlock(idxSheet, idxRow, 9, a)
    WriteIndexResult = ok
    If Not ok Then Exit Function

    If statusTxt = "OK" Then Exit Function
    Call XlReady
    On Error Resume Next
    idxSheet.Range(idxSheet.Cells(idxRow, 1), idxSheet.Cells(idxRow, 12)).Interior.Color = RGB(255, 220, 200)
    Err.Clear
End Function

' the Index writes of the run: what Excel could not take is kept and written
' again at the end (RedoIndexWrites), so no component is missing on the Index
Private Sub NoteIndexRow(idxSheet As Object, ByVal idxRow As Long, ByVal no As Long, _
                   ByVal pn As String, ByVal desc As String, ByVal lvl As Long, _
                   ByVal qty As Long, ByVal firstLevel As String, ByVal pathTxt As String, _
                   ByVal sheetName As String)
    If Not WriteIndexRow(idxSheet, idxRow, no, pn, desc, lvl, qty, firstLevel, pathTxt, sheetName, True) Then
        mIdxRedo.Add Array("ROW", idxRow, no, pn, desc, lvl, qty, firstLevel, pathTxt, sheetName)
    ElseIf Not mIdxFmtOK Then
        ' written before the text formats: once more after them
        mIdxEarly.Add Array(idxRow, no, pn, desc, lvl, qty, firstLevel, pathTxt, sheetName)
    End If
End Sub

Private Sub NoteIndexResult(idxSheet As Object, ByVal idxRow As Long, ByVal nRows As Long, _
                      ByVal statusTxt As String, ByVal secs As Double, ByVal whereTxt As String)
    If Not WriteIndexResult(idxSheet, idxRow, nRows, statusTxt, secs, whereTxt) Then
        mIdxRedo.Add Array("RES", idxRow, nRows, statusTxt, secs, whereTxt)
    End If
End Sub

Private Sub RedoIndexWrites(idxSheet As Object)
    Dim it As Variant, ok As Boolean
    If mIdxRedo Is Nothing Then Exit Sub
    If mIdxRedo.Count = 0 Then Exit Sub
    Call LogLine("writing " & mIdxRedo.Count & " Index entries again")
    mXlSkipAt = -1
    For Each it In mIdxRedo
        If it(0) = "ROW" Then
            ok = WriteIndexRow(idxSheet, CLng(it(1)), CLng(it(2)), CStr(it(3)), CStr(it(4)), _
                               CLng(it(5)), CLng(it(6)), CStr(it(7)), CStr(it(8)), CStr(it(9)), False)
        Else
            ok = WriteIndexResult(idxSheet, CLng(it(1)), CLng(it(2)), CStr(it(3)), CDbl(it(4)), CStr(it(5)))
        End If
        If Not ok Then Call LogLine("Index row " & it(1) & " could not be written")
    Next it
    Set mIdxRedo = New Collection
End Sub

Private Sub FinishIndexLayout(idxSheet As Object)
    Call XlReady
    On Error Resume Next
    idxSheet.Columns("A:L").AutoFit
    idxSheet.Columns("A:A").ColumnWidth = 5
    idxSheet.Columns("G:G").ColumnWidth = 60
    idxSheet.Range("A2:L2").AutoFilter
    Err.Clear
End Sub


' ==============================================================================
' HELPER: ASK AT WHICH LEVEL THE BOM IS SPLIT
' Shows how many components (and how many of them are sub-assemblies) every
' level has. Returns -1 when the user cancels, 0 for "no split".
' ==============================================================================
Private Function AskSplitLevel(rootProd As Product) As Long
    Dim cntAll(1 To 30) As Long, cntAssy(1 To 30) As Long
    Dim maxDepth As Long, lvl As Long, n As Long
    Dim msg As String, ans As String, deflt As String
    Dim t As Double

    AskSplitLevel = 0
    If Not ASK_SPLIT_LEVEL Then Exit Function

    Call SetStatus("BOM: reading the product structure ...")
    t = Timer
    Call ScanLevels(rootProd, 1, cntAll, cntAssy, maxDepth)
    Call SetStatus("")
    Call LogLine("structure read in " & Format(Elapsed(t), "0.0") & " s, " & maxDepth & " levels")

    msg = "At which level should the BOM be split?" & vbCrLf & vbCrLf & _
          "0 = no split: one sheet for the whole product" & vbCrLf & _
          "1 = one sheet per first level node, as in the CATIA tree" & vbCrLf & _
          "     (the parts directly below the product on one more sheet)" & vbCrLf & _
          "n = one sheet per sub-assembly on level n" & vbCrLf & vbCrLf & _
          "Level:  components  (sub-assemblies)" & vbCrLf
    For lvl = 1 To maxDepth
        If lvl > 15 Then
            msg = msg & "  ... deeper levels down to " & maxDepth & vbCrLf
            Exit For
        End If
        msg = msg & "  " & lvl & ":   " & cntAll(lvl) & "  (" & cntAssy(lvl) & ")" & vbCrLf
    Next lvl
    ' one sheet per first level node is offered first, when there is any
    If maxDepth >= 1 Then
        If cntAssy(1) > 0 Then deflt = "1" Else deflt = "0"
    Else
        deflt = "0"
    End If

    Do
        ans = Trim$(InputBox(msg, "BOM - split level", deflt))
        If Len(ans) = 0 Then
            AskSplitLevel = -1                       ' Cancel or empty
            Exit Function
        End If

        ' plain digits only - "1e10" or "2.5" would pass IsNumeric and then overflow
        If Len(ans) <= 3 And Not (ans Like "*[!0-9]*") Then
            n = CLng(ans)
            If n = 0 Then
                AskSplitLevel = 0
                Exit Function
            ElseIf n >= 1 And n <= maxDepth And n <= 30 Then
                If cntAssy(n) > 0 Then
                    AskSplitLevel = n
                    Exit Function
                End If
                MsgBox "Level " & n & " has no sub-assemblies, so there is nothing to split." & vbCrLf & _
                       "Please choose another level.", vbExclamation
            Else
                MsgBox "Please enter a number between 0 and " & maxDepth & ".", vbExclamation
            End If
        Else
            MsgBox "Please enter a number between 0 and " & maxDepth & ".", vbExclamation
        End If
    Loop
End Function

' counts the components per level - structure only, nothing is loaded
' (unless EXPAND_UNLOADED is set)
Private Sub ScanLevels(oProd As Product, ByVal lvl As Long, cntAll() As Long, _
                       cntAssy() As Long, ByRef maxDepth As Long)
    Dim i As Long, n As Long, nk As Long
    Dim child As Product

    If lvl > 30 Then Exit Sub
    On Error Resume Next
    n = 0
    n = oProd.Products.Count
    For i = 1 To n
        Set child = Nothing
        Set child = oProd.Products.Item(i)
        If Not child Is Nothing Then
            cntAll(lvl) = cntAll(lvl) + 1
            If lvl > maxDepth Then maxDepth = lvl
            nk = LoadedChildCount(child)
            If nk > 0 Then
                cntAssy(lvl) = cntAssy(lvl) + 1
                Call ScanLevels(child, lvl + 1, cntAll, cntAssy, maxDepth)
            End If
        End If
        Call Pump
    Next i
    Err.Clear
End Sub


' Number of children of a node.
' A sub-assembly that is NOT LOADED reports 0 and passes for a part. With
' EXPAND_UNLOADED its structure is loaded (visualization mode) - but that means
' a call per childless node, which on a big product loads a lot from Teamcenter
' before the first question appears, so by default the tree is taken as it is.
Private Function LoadedChildCount(child As Product) As Long
    Dim nk As Long
    Dim oDoc As Object
    Dim tn As String

    On Error Resume Next
    nk = 0
    nk = child.Products.Count
    If nk = 0 And EXPAND_UNLOADED Then
        Set oDoc = Nothing
        Set oDoc = child.ReferenceProduct.Parent
        tn = ""
        If Not oDoc Is Nothing Then tn = TypeName(oDoc)
        If tn <> "PartDocument" Then
            Err.Clear
            child.ApplyWorkMode 1
            nk = 0
            nk = child.Products.Count
        End If
    End If
    Err.Clear
    LoadedChildCount = nk
End Function


' ==============================================================================
' HELPER: THE COMPONENTS ON THE SPLIT LEVEL
'   dUnit*        one entry per part number (first occurrence + count)
'   colUnitInst   EVERY instance on the split level (to hide them all)
'   colAbove      every assembly instance above it
'   colLoose*     parts above (or on) the split level that sit in no component
' Structure only - nothing is loaded here. The content of every assembly above
' the split level is noted for the sums of the assembly rows.
' ==============================================================================
Private Sub CollectUnits(oProd As Product, ByVal lvl As Long, ByVal target As Long, _
                         ByVal firstName As String, ByVal pathTxt As String, _
                         ByVal parentKey As String, _
                         dUnitQty As Object, dUnitRef As Object, dUnitFirst As Object, _
                         dUnitDesc As Object, dUnitPath As Object, _
                         colUnitInst As Collection, colAbove As Collection, _
                         colLooseInst As Collection, colLooseLevel As Collection, _
                         colLooseFirst As Collection, colLooseKey As Collection, _
                         colLooseIsAssy As Collection, colLooseParts As Collection)
    Dim i As Long, n As Long, nk As Long
    Dim child As Product
    Dim pn As String, rawPn As String, childFirst As String, sDesc As String
    Dim childPath As String, rowKey As String, ownPn As String
    Dim kids As Collection

    On Error Resume Next
    n = 0
    n = oProd.Products.Count
    ownPn = ""
    ownPn = oProd.PartNumber
    Set kids = New Collection

    For i = 1 To n
        Set child = Nothing
        Set child = oProd.Products.Item(i)
        If Not child Is Nothing Then
            nk = 0
            nk = child.Products.Count
            rawPn = ""
            rawPn = child.PartNumber
            ' (a part whose part number cannot be read is not loaded - see TraverseTree)
            If Len(rawPn) = 0 And nk = 0 Then rawPn = ProvisionalName(child)
            pn = rawPn
            If Len(pn) = 0 Then pn = child.Name

            If lvl = 1 Then
                childFirst = FirstLevelLabel(child, pn)
            Else
                childFirst = firstName
            End If
            rowKey = MakeRowKey(parentKey, childFirst, pn)

            If Len(rawPn) > 0 Then kids.Add Array(rawPn, (nk > 0))

            If nk > 0 And lvl = target Then
                ' ---- a component: one sheet per part number ----
                colUnitInst.Add child
                If dUnitQty.Exists(pn) Then
                    ' used again, maybe in another branch: keep every branch and path
                    dUnitQty(pn) = dUnitQty(pn) + 1
                    Call AppendUnique(dUnitFirst, pn, childFirst, ", ")
                    Call AppendUnique(dUnitPath, pn, pathTxt, "; ")
                Else
                    dUnitQty.Add pn, 1
                    dUnitRef.Add pn, child
                    dUnitFirst.Add pn, childFirst
                    dUnitPath.Add pn, pathTxt
                    sDesc = ""
                    sDesc = child.DescriptionRef
                    dUnitDesc.Add pn, sDesc
                End If

            ElseIf nk > 0 Then
                ' ---- an assembly above the split level ----
                colAbove.Add child
                ' in the assembly structure it gets a row on the extra sheet, so
                ' the tree above the components stays visible
                If mIncludeSubassemblies Then
                    colLooseInst.Add child
                    colLooseLevel.Add lvl
                    colLooseFirst.Add childFirst
                    colLooseKey.Add rowKey
                    colLooseIsAssy.Add True
                End If
                ' the path in the names of the CATIA tree
                If Len(pathTxt) > 0 Then
                    childPath = pathTxt & " > " & TreeName(child, pn)
                Else
                    childPath = TreeName(child, pn)
                End If
                Call CollectUnits(child, lvl + 1, target, childFirst, childPath, rowKey, _
                                  dUnitQty, dUnitRef, dUnitFirst, dUnitDesc, dUnitPath, _
                                  colUnitInst, colAbove, colLooseInst, colLooseLevel, _
                                  colLooseFirst, colLooseKey, colLooseIsAssy, colLooseParts)

            Else
                ' ---- a part that is not inside any component ----
                colLooseInst.Add child
                colLooseLevel.Add lvl
                colLooseFirst.Add childFirst
                colLooseKey.Add rowKey
                colLooseIsAssy.Add False
                colLooseParts.Add child
            End If
        End If
        Call Pump
    Next i

    Call NoteAssemblyContent(ownPn, kids)
    Err.Clear
End Sub

' appends a value to a dictionary entry unless it is already in the list
Private Sub AppendUnique(d As Object, ByVal k As String, ByVal v As String, ByVal sep As String)
    Dim cur As String

    If Len(v) = 0 Then Exit Sub
    If Not d.Exists(k) Then
        d.Add k, v
        Exit Sub
    End If
    cur = CStr(d(k))
    If Len(cur) = 0 Then
        d(k) = v
    ElseIf InStr(1, sep & cur & sep, sep & v & sep, vbTextCompare) = 0 Then
        d(k) = cur & sep & v
    End If
End Sub


' ==============================================================================
' HELPER: THE DOCUMENT OF A COMPONENT, IF IT HAS ONE OF ITS OWN
' A component gets its own window only when it is the ROOT of its own
' CATProduct. A component that is defined inside another file (for example a
' "Product1" inserted into an assembly) would open that whole other file, so it
' returns Nothing and the component is done in the main window instead.
' ==============================================================================
Private Function OwnDocumentOf(oInst As Product, ByVal pn As String, rootDoc As Document) As Object
    Dim oRef As Product
    Dim oDoc As Object
    Dim docRootPN As String

    Set OwnDocumentOf = Nothing
    On Error Resume Next

    Set oRef = oInst.ReferenceProduct
    If oRef Is Nothing Then Err.Clear: Exit Function
    Set oDoc = oRef.Parent
    If oDoc Is Nothing Then Err.Clear: Exit Function
    If oDoc Is rootDoc Then Err.Clear: Exit Function
    If StrComp(oDoc.FullName, rootDoc.FullName, vbTextCompare) = 0 Then Err.Clear: Exit Function
    If InStr(1, oDoc.Name, ".CATProduct", vbTextCompare) = 0 Then Err.Clear: Exit Function

    docRootPN = ""
    docRootPN = oDoc.Product.PartNumber
    If StrComp(docRootPN, pn, vbTextCompare) <> 0 Then Err.Clear: Exit Function

    Set OwnDocumentOf = oDoc
    Err.Clear
End Function


' ==============================================================================
' HELPER: A WINDOW READY FOR PICTURES - geometry only, fixed iso camera
' ==============================================================================
Private Sub PrepareWindow(oWin As Object, ByRef oViewer As Viewer, ByRef savedLayout As Long, _
                          ByRef layoutSet As Boolean)
    On Error Resume Next
    Set oViewer = Nothing
    layoutSet = False
    savedLayout = -1

    Set oViewer = oWin.ActiveViewer
    savedLayout = oWin.Layout
    oWin.Layout = catWindowGeomOnly            ' no spec tree in the pictures
    layoutSet = (Err.Number = 0 And savedLayout >= 0)
    Err.Clear

    If SET_ISO_VIEW Then Call SetIsoViewpoint(oViewer)
    Call Settle(100)
    Err.Clear
End Sub


' ==============================================================================
' HELPER: THE COMPASS (only a translated command name exists for it)
' ==============================================================================
Private Function ToggleCompass() As Boolean
    ToggleCompass = False
    If Len(COMPASS_COMMAND) = 0 Then Exit Function
    On Error Resume Next
    CATIA.StartCommand COMPASS_COMMAND
    ToggleCompass = (Err.Number = 0)
    Err.Clear
End Function


' ==============================================================================
' HELPER: EXCEL SHEETS
' ==============================================================================
Private Sub KeepOneSheet(xlBook As Object)
    Dim i As Long
    Call XlReady
    On Error Resume Next
    xlBook.Application.DisplayAlerts = False
    For i = xlBook.Sheets.Count To 2 Step -1
        xlBook.Sheets(i).Delete
    Next i
    xlBook.Application.DisplayAlerts = True
    Err.Clear
End Sub

' a new sheet at the end of the workbook, or Nothing when Excel refuses
Private Function AddSheet(xlBook As Object, ByVal baseName As String) As Object
    Dim ws As Object, e As Long
    Dim t0 As Double, told As Boolean

    Set AddSheet = Nothing
    ' creating a sheet is not one cell that can be skipped: it always waits and
    ' asks again, whatever was answered for an earlier write
    mXlSkipAt = -1
    t0 = Timer
    Do
        If Not XlReady() Then
            Call LogLine("skipped: the sheet " & baseName)
            Exit Function
        End If
        On Error Resume Next
        Err.Clear
        Set ws = Nothing
        Set ws = xlBook.Sheets.Add(, xlBook.Sheets(xlBook.Sheets.Count))
        e = Err.Number
        Err.Clear
        On Error GoTo 0
        If Not ws Is Nothing Then Exit Do
        Call NoteXlError(e)
        If Not IsExcelBusy(e) Then
            Call LogLine("Excel could not add the sheet " & baseName & ": error " & e)
            Exit Function
        End If
        If Not XlKeepTrying(t0, told) Then
            Call LogLine("skipped: the sheet " & baseName)
            Exit Function
        End If
    Loop

    Call NameSheet(ws, baseName)
    Set AddSheet = ws
End Function

' gives a sheet a valid, unused name - a name Excel refuses never ends the run
Private Sub NameSheet(ws As Object, ByVal baseName As String)
    Dim nm As String

    nm = SafeSheetName(baseName)
    Call XlReady
    On Error Resume Next
    Err.Clear
    ws.Name = nm
    If Err.Number <> 0 Then
        Err.Clear
        nm = SafeSheetName("Sheet")
        ws.Name = nm
    End If
    Err.Clear
    nm = ws.Name                         ' whatever Excel really has
    Err.Clear
    On Error GoTo 0
    If Not mSheetNames.Exists(nm) Then mSheetNames.Add nm, True
    mSheetName = nm
End Sub

' Excel sheet names: max 31 characters, none of  : \ / ? * [ ] '  and unique
Private Function SafeSheetName(ByVal baseName As String) As String
    Dim s As String, cand As String, bad As Variant
    Dim k As Long, n As Long

    s = baseName
    For Each bad In Array(":", "\", "/", "?", "*", "[", "]", "'")
        s = Replace(s, CStr(bad), "_")
    Next bad
    s = Trim$(s)
    If Len(s) = 0 Then s = "Sheet"
    If StrComp(s, "History", vbTextCompare) = 0 Then s = "History_"
    If Len(s) > 31 Then s = Left$(s, 31)

    cand = s
    n = 1
    Do While mSheetNames.Exists(cand)
        n = n + 1
        k = Len(CStr(n)) + 1
        cand = Left$(s, 31 - k) & "~" & n
    Loop
    SafeSheetName = cand
End Function

Private Function SheetNameOf(ws As Object) As String
    SheetNameOf = mSheetName
    Call XlReady
    On Error Resume Next
    SheetNameOf = ws.Name
    Err.Clear
End Function

Private Sub ShowSheet(ws As Object)
    If ws Is Nothing Then Exit Sub
    Call XlReady
    On Error Resume Next
    ws.Activate
    Err.Clear
End Sub


' ==============================================================================
' HELPER: TALKING TO A BUSY EXCEL
' Excel refuses calls from another program while it is busy - a cell in edit
' mode, an open dialog or menu, a drag in progress. VBA then gets error 50290
' (or "call was rejected"). Every write therefore keeps trying for
' EXCEL_WAIT_SECONDS, then asks the user (Retry / Cancel). A write that cannot
' be done is logged and skipped - it never ends the run.
' ==============================================================================
Private Function IsExcelBusy(ByVal e As Long) As Boolean
    Select Case e
        Case 50290, -2146777998            ' 0x800AC472: Excel is busy
            IsExcelBusy = True
        Case -2147418111                   ' 0x80010001: call was rejected by callee
            IsExcelBusy = True
        Case -2147417846                   ' 0x8001010A: server busy, retry later
            IsExcelBusy = True
        Case Else
            IsExcelBusy = False
    End Select
End Function

' Excel was closed or has crashed: the COM server is gone for good
Private Function IsExcelGone(ByVal e As Long) As Boolean
    Select Case e
        Case 462                           ' remote server machine does not exist or is unavailable
            IsExcelGone = True
        Case -2147023174, -2147023170      ' 0x800706BA / 0x800706BE: RPC server unavailable / call failed
            IsExcelGone = True
        Case -2147417848                   ' 0x80010108: object disconnected from its clients
            IsExcelGone = True
        Case Else
            IsExcelGone = False
    End Select
End Function

' ends the current component (or the run, outside a component) when Excel is gone
Private Sub CheckExcel()
    Dim e As Long, nm As String

    ' the workbook itself can be closed while Excel stays open
    If Not mXlLost And Not mXlBook Is Nothing Then
        On Error Resume Next
        Err.Clear
        nm = mXlBook.Name
        e = Err.Number
        Err.Clear
        On Error GoTo 0
        If e <> 0 And Not IsExcelBusy(e) Then
            mXlLost = True
            Call LogLine("the BOM workbook is gone (error " & e & ") - nothing can be written any more")
        End If
    End If
    If mXlLost Then Err.Raise vbObjectError + 515, , _
        "Excel or the BOM workbook was closed - the run is stopped"
End Sub

' notes the outcome of an Excel call
Private Sub NoteXlError(ByVal e As Long)
    If IsExcelGone(e) And Not mXlLost Then
        mXlLost = True
        Call LogLine("Excel is gone (error " & e & ") - nothing can be written any more")
    End If
End Sub

' After a "busy" answer from Excel: wait a little and say whether to try again.
' Excel is given EXCEL_WAIT_SECONDS, then the user is asked - no data is ever
' dropped silently. After a "skip", writes are skipped without asking again for
' 30 seconds, so one stuck cell edit does not bring up a question per row.
Private Function XlKeepTrying(ByRef t0 As Double, ByRef told As Boolean) As Boolean
    Dim ans As Long

    XlKeepTrying = False
    If mXlLost Then Exit Function
    If mXlSkipAt >= 0 Then
        If Elapsed(mXlSkipAt) < 30 Then Exit Function
        mXlSkipAt = -1
    End If

    If Not told Then
        told = True
        Call LogLine("Excel is busy - waiting (a cell in edit mode or a dialog open in the BOM Excel?)")
        Call SetStatus("BOM: Excel is busy - finish the cell you are editing or close the dialog in Excel ...")
    End If
    If Elapsed(t0) > EXCEL_WAIT_SECONDS Then
        Call LogLine("Excel still busy after " & EXCEL_WAIT_SECONDS & " s - asking the user")
        ans = MsgBox("The Excel window of the BOM does not accept data right now." & vbCrLf & _
                     "Most likely a cell is being edited there, or a dialog is open." & vbCrLf & vbCrLf & _
                     "Finish it in Excel (Enter or Esc), then click Retry." & vbCrLf & _
                     "Cancel = skip what Excel cannot take now and go on.", _
                     vbRetryCancel + vbExclamation, "BOM - Excel is busy")
        If ans = vbCancel Then
            Call LogLine("the user chose to skip writes while Excel is busy")
            mXlSkipAt = Timer
            Exit Function
        End If
        t0 = Timer
    End If
    Call Settle(500)
    XlKeepTrying = True
End Function

' waits (CATIA stays alive) until Excel takes calls again. Application.Ready is
' only a hint - it can say True while a cell is being edited - so every write
' still retries on a "busy" error itself.
' False = Excel is gone, or the user chose to skip.
Private Function XlReady() As Boolean
    Dim t0 As Double, ok As Boolean, e As Long, told As Boolean

    XlReady = False
    If mXlApp Is Nothing Or mXlLost Then Exit Function
    t0 = Timer
    Do
        ok = False
        On Error Resume Next
        Err.Clear
        ok = mXlApp.Ready
        e = Err.Number
        Err.Clear
        On Error GoTo 0
        If e <> 0 Then
            Call NoteXlError(e)
            If mXlLost Then Exit Function
            ' an error that does not mean "busy": cannot tell, just try the call
            If Not IsExcelBusy(e) Then ok = True
        End If
        If ok Then
            XlReady = True
            Exit Function
        End If
        If Not XlKeepTrying(t0, told) Then Exit Function
    Loop
End Function

' writes a 2-D array (1-based) to the sheet with its top left corner at (r, c)
Private Function XlWriteBlock(ws As Object, ByVal r As Long, ByVal c As Long, arr As Variant) As Boolean
    Dim e As Long, nr As Long, nc As Long
    Dim t0 As Double, told As Boolean

    XlWriteBlock = False
    mXlLastErr = 0
    If ws Is Nothing Then Exit Function
    nr = UBound(arr, 1) - LBound(arr, 1) + 1
    nc = UBound(arr, 2) - LBound(arr, 2) + 1
    t0 = Timer
    Do
        If Not XlReady() Then
            Call LogLine("skipped: Excel write at row " & r & ", column " & c)
            Exit Function
        End If
        On Error Resume Next
        Err.Clear
        ws.Range(ws.Cells(r, c), ws.Cells(r + nr - 1, c + nc - 1)).Value = arr
        e = Err.Number
        Err.Clear
        On Error GoTo 0
        If e = 0 Then
            XlWriteBlock = True
            Exit Function
        End If
        Call NoteXlError(e)
        If Not IsExcelBusy(e) Then
            mXlLastErr = e
            Call LogLine("Excel write failed at row " & r & ", column " & c & ": error " & e)
            Exit Function
        End If
        ' "busy" is answered before Excel does anything, so trying again is safe
        If Not XlKeepTrying(t0, told) Then
            Call LogLine("skipped: Excel write at row " & r & ", column " & c)
            Exit Function
        End If
    Loop
End Function

' one cell
Private Function XlPutCell(ws As Object, ByVal r As Long, ByVal c As Long, ByVal v As Variant) As Boolean
    Dim a(1 To 1, 1 To 1) As Variant
    a(1, 1) = v
    XlPutCell = XlWriteBlock(ws, r, c, a)
End Function

' one picture into column A of a row, sized to the row
Private Function XlInsertPicture(ws As Object, ByVal rowNo As Long, ByVal picPath As String) As Boolean
    Dim shp As Object, e As Long
    Dim x As Double, y As Double
    Dim t0 As Double, told As Boolean

    XlInsertPicture = False
    t0 = Timer
    Do
        If Not XlReady() Then Exit Function
        On Error Resume Next
        Err.Clear
        Set shp = Nothing
        x = ws.Cells(rowNo, 1).Left + 2
        y = ws.Cells(rowNo, 1).Top + 2
        If Err.Number = 0 Then Set shp = ws.Shapes.AddPicture(picPath, False, True, x, y, -1, -1)
        e = Err.Number
        Err.Clear
        On Error GoTo 0
        If Not shp Is Nothing Then Exit Do
        Call NoteXlError(e)
        If Not IsExcelBusy(e) Then
            Call LogLine("Excel could not insert the picture of row " & rowNo & ": error " & e)
            Exit Function
        End If
        If Not XlKeepTrying(t0, told) Then Exit Function
    Loop

    On Error Resume Next
    shp.Height = 50
    If shp.Width > 90 Then shp.Width = 90
    shp.Placement = 1                        ' move and size with cells: a folded row hides its picture
    ws.Rows(rowNo).RowHeight = 60
    Err.Clear
    On Error GoTo 0
    XlInsertPicture = True
End Function

' starts the Excel of the run (or attaches to the open one when
' OWN_EXCEL_INSTANCE is False), makes it visible and adds the workbook
Private Function StartExcel(ByRef firstSheet As Object) As Boolean
    Dim e As Long, t0 As Double, told As Boolean

    StartExcel = False
    Set firstSheet = Nothing
    On Error Resume Next
    If Not OWN_EXCEL_INSTANCE Then Set mXlApp = GetObject(, "Excel.Application")
    Err.Clear
    If mXlApp Is Nothing Then
        Set mXlApp = CreateObject("Excel.Application")
        mOwnXl = Not (mXlApp Is Nothing)
    End If
    Err.Clear
    On Error GoTo 0
    If mXlApp Is Nothing Then
        Call LogLine("Excel could not be started")
        Exit Function
    End If

    ' a fresh Excel can still be loading its add-ins and refuse calls for a while
    t0 = Timer
    Do
        On Error Resume Next
        Err.Clear
        mXlApp.Visible = True
        If Err.Number = 0 And mXlBook Is Nothing Then Set mXlBook = mXlApp.Workbooks.Add
        If Err.Number = 0 And Not mXlBook Is Nothing Then Set firstSheet = mXlBook.Sheets(1)
        e = Err.Number
        Err.Clear
        On Error GoTo 0
        If e = 0 And Not firstSheet Is Nothing Then Exit Do
        Call NoteXlError(e)
        If Not IsExcelBusy(e) Then
            Call LogLine("Excel could not be prepared: error " & e)
            Exit Function
        End If
        If Not XlKeepTrying(t0, told) Then Exit Function
    Loop

    ' an Excel started by the macro stays open when the macro ends
    On Error Resume Next
    If mOwnXl Then mXlApp.UserControl = True
    Err.Clear
    On Error GoTo 0
    StartExcel = True
End Function

' Excel back to normal at the end of the run (also after an error). An own
' Excel that never got its workbook is closed again instead of being left
' running invisible.
Private Sub XlRestore()
    If mXlApp Is Nothing Then Exit Sub
    On Error Resume Next
    If mOwnXl And mXlBook Is Nothing Then
        mXlApp.Quit
        Set mXlApp = Nothing
        Err.Clear
        Exit Sub
    End If
    mXlApp.ScreenUpdating = True
    mXlApp.DisplayAlerts = True
    If mOwnXl Then mXlApp.UserControl = True
    Err.Clear
End Sub


' ==============================================================================
' HELPER: SAVING, LOG, STATUS
' ==============================================================================
' log file and save path of this run:  <folder>\BOM_<product>_<date>_<time>.txt/.xlsx
Private Sub StartRunFiles(ByVal rootPN As String)
    Dim folder As String, baseName As String, bad As Variant
    Dim k As Long, f As Integer, e As Long, ch As String

    mLogPath = ""
    mSavePath = ""
    mSavedOnce = False
    mSaveTried = False
    mLastSaveOK = False
    mLastSave = Timer
    Set mSlow = New Collection

    ' file name: no character Windows or Excel refuses
    baseName = rootPN
    For Each bad In Array("\", "/", ":", "*", "?", Chr(34), "<", ">", "|", "[", "]")
        baseName = Replace(baseName, CStr(bad), "_")
    Next bad
    For k = 0 To 31
        baseName = Replace(baseName, Chr(k), "_")
    Next k
    ' VBA's Open (the log) only takes characters of the Windows code page
    For k = 1 To Len(baseName)
        ch = Mid$(baseName, k, 1)
        If StrConv(StrConv(ch, vbFromUnicode), vbUnicode) <> ch Then Mid$(baseName, k, 1) = "_"
    Next k
    baseName = Trim$(baseName)
    If Len(baseName) = 0 Then baseName = "product"
    If Len(baseName) > 60 Then baseName = Left$(baseName, 60)
    baseName = "BOM_" & baseName & "_" & Format(Now, "yyyymmdd_hhnnss")

    ' the folder must be WRITABLE, not only exist - otherwise the user's TEMP
    folder = SAVE_FOLDER
    If Len(folder) = 0 Then folder = "C:\Temp"
    Call EnsureFolder(folder)
    ' the probe is the log file itself (the run appends to it anyway); Excel's
    ' Save As also refuses a full path of more than 218 characters
    On Error Resume Next
    Err.Clear
    f = FreeFile
    Open folder & "\" & baseName & ".txt" For Append As #f
    e = Err.Number
    Close #f
    Err.Clear
    If e = 0 And Len(folder & "\" & baseName & ".xlsx") > 218 Then e = 1
    If e <> 0 Then folder = CStr(Environ$("TEMP"))
    Err.Clear
    On Error GoTo 0

    mLogPath = folder & "\" & baseName & ".txt"
    If Len(SAVE_FOLDER) > 0 Then mSavePath = folder & "\" & baseName & ".xlsx"
End Sub

' saves the workbook (first time: Save As). force = False saves at most every
' SAVE_EVERY_SECONDS. A failed save is logged and the run goes on.
Private Sub SaveWorkbook(ByVal force As Boolean)
    Dim e As Long, d As String

    If Len(mSavePath) = 0 Or mXlBook Is Nothing Then Exit Sub
    If mXlLost Then
        mLastSaveOK = False
        Exit Sub
    End If
    ' throttled after any attempt - a save that keeps failing is not retried
    ' after every picture
    If Not force And mSaveTried And Elapsed(mLastSave) < SAVE_EVERY_SECONDS Then Exit Sub

    Call SetStatus("BOM: saving the workbook ...")
    Call XlReady
    On Error Resume Next
    ' no question from Excel may ever wait for a click in the middle of a call
    mXlApp.DisplayAlerts = False
    Err.Clear
    If mSavedOnce Then
        mXlBook.Save
    Else
        mXlBook.SaveAs mSavePath, 51           ' 51 = .xlsx
    End If
    e = Err.Number
    d = Err.Description
    Err.Clear
    mXlApp.DisplayAlerts = True
    Err.Clear
    On Error GoTo 0
    mLastSave = Timer
    mSaveTried = True
    mLastSaveOK = (e = 0)

    If e = 0 Then
        mSavedOnce = True
        mLastSaveAt = Now
        Call LogLine("saved " & mSavePath)
    Else
        Call NoteXlError(e)
        Call LogLine("saving failed: error " & e & " - " & d)
    End If
End Sub

Private Function SavedText() As String
    Dim s As String
    If mSavedOnce And mLastSaveOK Then
        s = "Saved as:" & vbCrLf & mSavePath
    ElseIf mSavedOnce Then
        s = "The LAST save did not work - only the copy of " & Format(mLastSaveAt, "hh:nn") & _
            " is saved as:" & vbCrLf & mSavePath & vbCrLf & _
            "Save the workbook in Excel yourself if it is still open."
    ElseIf Len(mSavePath) > 0 Then
        s = "The workbook could NOT be saved automatically - please save it yourself."
    Else
        s = "The workbook is not saved yet."
    End If
    If Len(mLogPath) > 0 Then s = s & vbCrLf & vbCrLf & "Log of the run:" & vbCrLf & mLogPath
    SavedText = s
End Function

' the sub-assemblies CATIA had not loaded - their content is missing in the BOM
Private Function NotLoadedText() As String
    Dim k As Variant, s As String, nA As Long, nP As Long, listA As String, listP As String

    NotLoadedText = ""
    If mNotLoaded Is Nothing Then Exit Function
    If mNotLoaded.Count = 0 Then Exit Function
    For Each k In mNotLoaded.Keys
        If mNoPartNumber.Exists(CStr(k)) Then
            nP = nP + 1
            If nP <= 6 Then listP = listP & "  " & CStr(k) & vbCrLf
        Else
            nA = nA + 1
            If nA <= 6 Then listA = listA & "  " & CStr(k) & vbCrLf
        End If
    Next k
    If nA > 6 Then listA = listA & "  ..." & vbCrLf
    If nP > 6 Then listP = listP & "  ..." & vbCrLf
    If nA > 0 Then
        s = s & nA & " sub-assembly(ies) NOT LOADED (or empty) in CATIA - their content is" & vbCrLf & _
                "missing in the BOM (column L):" & vbCrLf & listA & vbCrLf
    End If
    If nP > 0 Then
        s = s & nP & " component(s) NOT LOADED in CATIA (no part number readable) -" & vbCrLf & _
                "rows of Type ""Not loaded"" without values:" & vbCrLf & listP & vbCrLf
    End If
    NotLoadedText = s
End Function

' one line into the log file - opened and closed every time, so the file is
' complete even if CATIA has to be killed
Private Sub LogLine(ByVal txt As String)
    Dim f As Integer
    If Len(mLogPath) = 0 Then Exit Sub
    On Error Resume Next
    f = FreeFile
    Err.Clear
    Open mLogPath For Append As #f
    If Err.Number <> 0 Then
        mLogPath = ""                      ' no log possible - stop trying
    Else
        Print #f, Format(Now, "hh:nn:ss") & "  " & txt
        Close #f
    End If
    Err.Clear
End Sub

' a step that took long: into the log and onto the "Log" sheet
Private Sub NoteSlow(ByVal secs As Double, ByVal txt As String)
    Call LogLine("  SLOW " & Format(secs, "0.0") & " s: " & txt)
    If mSlow Is Nothing Then Exit Sub
    If mSlow.Count < 2000 Then mSlow.Add Array(secs, txt)
End Sub

' the "Log" sheet: run summary and every slow step
Private Sub WriteLogSheet(ByVal summary As String)
    Dim ws As Object, a() As Variant, it As Variant
    Dim n As Long, i As Long

    If mXlBook Is Nothing Then Exit Sub
    Set ws = AddSheet(mXlBook, "Log")
    If ws Is Nothing Then Exit Sub

    n = 0
    If Not mSlow Is Nothing Then n = mSlow.Count
    ReDim a(1 To n + 5, 1 To 2)
    a(1, 1) = "Run"
    a(1, 2) = summary
    a(2, 1) = "Log file"
    a(2, 2) = mLogPath
    a(4, 1) = "Seconds"
    a(4, 2) = "Slow steps (" & SLOW_STEP_SECONDS & " s or more)"
    If n = 0 Then a(5, 2) = "none"
    For i = 1 To n
        it = mSlow(i)
        a(i + 4, 1) = Round(CDbl(it(0)), 1)
        a(i + 4, 2) = CStr(it(1))
    Next i
    Call XlWriteBlock(ws, 1, 1, a)

    Call XlReady
    On Error Resume Next
    ws.Columns("A:A").ColumnWidth = 10
    ws.Columns("B:B").ColumnWidth = 110
    ws.Range("A4:B4").Font.Bold = True
    Err.Clear
End Sub

Private Sub SetStatus(ByVal txt As String)
    On Error Resume Next
    CATIA.StatusBar = txt
    Err.Clear
End Sub

' seconds since t0 - also across midnight, where Timer starts again at 0
Private Function Elapsed(ByVal t0 As Double) As Double
    Dim d As Double
    d = Timer - t0
    If d < 0 Then d = d + 86400#
    Elapsed = d
End Function

' the tree walks call this for every node: CATIA keeps its window alive
Private Sub Pump()
    mPump = mPump + 1
    If mPump Mod 50 = 0 Then DoEvents
    If mPump Mod 1000 = 0 Then Call SetStatus("BOM: walking the product tree ... " & mPump & " nodes")
End Sub

Private Sub EnsureFolder(ByVal p As String)
    Dim parentPath As String
    If Len(p) = 0 Or mFso Is Nothing Then Exit Sub
    On Error Resume Next
    If mFso.FolderExists(p) Then Exit Sub
    parentPath = mFso.GetParentFolderName(p)
    If Len(parentPath) > 0 And parentPath <> p Then Call EnsureFolder(parentPath)
    mFso.CreateFolder p
    Err.Clear
End Sub

' ==============================================================================
' HELPER: WAIT WHILE KEEPING CATIA ALIVE
' The old code used the kernel32 Sleep, which blocks the thread WITHOUT pumping
' the message queue. CATIA is single threaded: during such a sleep it cannot
' redraw, Windows paints it as "not responding", and the view we are about to
' capture never gets the chance to update either.
' ==============================================================================
Sub Settle(ByVal ms As Long)
    Dim t As Single

    If ms <= 0 Then
        DoEvents
        Exit Sub
    End If

    t = Timer
    Do
        DoEvents
        Sleep 5                      ' short nap so the loop does not spin the CPU
    Loop While (Timer - t) * 1000# < ms And Timer >= t
End Sub

' ==============================================================================
' HELPER: ONE FIXED ISOMETRIC VIEWPOINT FOR ALL PICTURES
' Without it every thumbnail inherits whatever camera the session was left in,
' so a flat part can end up as a single line.
' ==============================================================================
Sub SetIsoViewpoint(oViewer As Viewer)
    On Error Resume Next
    Dim vp As Object
    Dim sight(2) As Variant
    Dim up(2) As Variant

    If oViewer Is Nothing Then Exit Sub

    Set vp = oViewer.Viewpoint3D
    If vp Is Nothing Then Err.Clear: Exit Sub

    sight(0) = -1#: sight(1) = -1#: sight(2) = -1#     ' look from +X +Y +Z
    up(0) = 0#: up(1) = 0#: up(2) = 1#

    vp.PutSightDirection sight
    vp.PutUpDirection up
    oViewer.Viewpoint3D = vp
    oViewer.Reframe
    oViewer.Update
    Err.Clear
End Sub

' ==============================================================================
' HELPER: BATCHED SHOW / HIDE FOR A COLLECTION OF PRODUCTS
' One Selection.Add per item but only ONE SetShow call (= one visual update).
' CATIA refuses the whole SetShow when ONE element cannot be shown or hidden
' (e.g. a component that is not loaded) - then a refused range is halved again
' and again, so the refused element costs a few calls, not the whole list.
' In CATIA: SetShow(0) = SHOW, SetShow(1) = HIDE
' ==============================================================================
Sub SetShowCollection(sel As Selection, colItems As Collection, ByVal showMode As Integer)
    If colItems Is Nothing Then Exit Sub
    If colItems.Count = 0 Then Exit Sub
    Call SetShowSplit(sel, colItems, 1, colItems.Count, showMode)
End Sub

Private Sub SetShowSplit(sel As Selection, colItems As Collection, ByVal i1 As Long, _
                         ByVal i2 As Long, ByVal showMode As Integer)
    Dim m As Long
    If SetShowRange(sel, colItems, i1, i2, showMode) Then Exit Sub
    If i1 >= i2 Then Exit Sub                ' one element CATIA refuses - left as it is
    m = (i1 + i2) \ 2
    DoEvents
    Call SetShowSplit(sel, colItems, i1, m, showMode)
    Call SetShowSplit(sel, colItems, m + 1, i2, showMode)
End Sub

' items i1..i2 of the list in one selection and one SetShow; False = refused
Private Function SetShowRange(sel As Selection, colItems As Collection, ByVal i1 As Long, _
                              ByVal i2 As Long, ByVal showMode As Integer) As Boolean
    Dim i As Long, ok As Boolean

    On Error Resume Next
    sel.Clear
    For i = i1 To i2
        sel.Add colItems.Item(i)
        If i Mod 200 = 0 Then DoEvents     ' a long list must not freeze CATIA
    Next i
    Err.Clear
    sel.VisProperties.SetShow showMode
    ok = (Err.Number = 0)
    Err.Clear
    sel.Clear
    Err.Clear
    SetShowRange = ok
End Function

' ==============================================================================
' HELPER: PUT THE VISIBILITY BACK
' With PRESERVE_VISIBILITY the exact state from before the run is restored,
' otherwise everything is simply made visible again (the leaves are what the
' run hid).
' ==============================================================================
Sub RestoreVisibility(sel As Selection, colLeaves As Collection, _
                      colWasShown As Collection, colWasHidden As Collection)
    On Error Resume Next

    If PRESERVE_VISIBILITY And Not colWasShown Is Nothing Then
        Call SetShowCollection(sel, colWasShown, 0)
        Call SetShowCollection(sel, colWasHidden, 1)
    Else
        Call SetShowCollection(sel, colLeaves, 0)
    End If
    Err.Clear
End Sub

' ==============================================================================
' HELPER: SPLIT A COLLECTION INTO "WAS VISIBLE" AND "WAS HIDDEN"
' Late bound on purpose: GetShow wants a CatVisPropertyShow out-parameter, and
' going through Object keeps the module compiling even where that enum is not
' exposed.
' ==============================================================================
Sub SplitByShowState(sel As Selection, colItems As Collection, _
                     colShown As Collection, colHidden As Collection)
    On Error Resume Next
    Dim i As Long
    Dim oSel As Object
    Dim vShow As Long

    If colItems Is Nothing Then Exit Sub
    Set oSel = sel

    For i = 1 To colItems.Count
        oSel.Clear
        oSel.Add colItems.Item(i)
        vShow = 1                                  ' 1 = NoShow, assumed on error
        oSel.VisProperties.GetShow vShow
        If vShow = 0 Then                          ' 0 = catVisPropertyShowAttr
            colShown.Add colItems.Item(i)
        Else
            colHidden.Add colItems.Item(i)
        End If
        Err.Clear
    Next i
    oSel.Clear
    Err.Clear
End Sub

' ==============================================================================
' HELPER: SHOW / HIDE A SINGLE OBJECT
' ==============================================================================
Sub SetShowSingle(sel As Selection, oObj As Object, ByVal showMode As Integer)
    On Error Resume Next
    If oObj Is Nothing Then Exit Sub
    sel.Clear
    sel.Add oObj
    sel.VisProperties.SetShow showMode
    sel.Clear
    Err.Clear
End Sub

' ==============================================================================
' HELPER: SHOW / HIDE A LIST OF BODIES (by name) IN ONE SetShow CALL
' ==============================================================================
Sub SetShowBodyList(oPart As Part, partSel As Selection, bodyNames As Variant, ByVal showMode As Integer)
    On Error Resume Next
    Dim i As Integer
    Dim oBody As Body

    partSel.Clear
    For i = LBound(bodyNames) To UBound(bodyNames)
        Set oBody = GetBodyByName(oPart, CStr(bodyNames(i)))
        If Not oBody Is Nothing Then partSel.Add oBody
    Next i
    partSel.VisProperties.SetShow showMode
    partSel.Clear
    Err.Clear
End Sub

' ==============================================================================
' HELPER: TRAVERSE TREE (ROWS + LEVELS + FIRST LEVEL, structure only)
' Level 1 = direct children of root, Level 2 = grandchildren, etc.
' Also fills:
'   - colLeaves: EVERY leaf instance (used to hide them all in one call)
'   - mAssyKids: the content of every assembly part number, for the sums of
'     the assembly rows
' Nothing is loaded or measured here.
' ==============================================================================
Function TraverseTree(oProd As Product, dQty As Object, dRef As Object, dDesc As Object, _
                      dProps As Object, dLevel As Object, dBodies As Object, _
                      dPartNum As Object, dFirst As Object, dIsAssy As Object, _
                      dNodeLeaves As Object, colLeaves As Collection, _
                      ByVal currentLevel As Integer, ByVal firstLevelName As String, _
                      ByVal parentKey As String) As Collection

    Dim childProd As Product, i As Long, partNum As String
    Dim myLeaves As Collection, childLeaves As Collection
    Dim k As Long, nk As Long
    Dim sKey As String, childFirst As String, ownPn As String
    Dim kids As Collection

    Set myLeaves = New Collection
    Set TraverseTree = myLeaves          ' set early: an error still returns a list

    If oProd Is Nothing Then Exit Function
    If oProd.Products.Count = 0 Then Exit Function

    ownPn = ""
    On Error Resume Next
    ownPn = oProd.PartNumber
    Err.Clear
    On Error GoTo 0
    Set kids = New Collection

    For i = 1 To oProd.Products.Count
        Set childProd = oProd.Products.Item(i)

        ' NOTE: no ApplyWorkMode here. It is recursive: on an assembly it loads
        ' every CATPart below in one blocking call. Each PART row promotes its
        ' own part in WriteDataRows instead.
        mAllNodes.Add childProd          ' assemblies and leaves, for the show/hide baseline

        partNum = ""
        On Error Resume Next
        partNum = childProd.PartNumber
        Err.Clear
        On Error GoTo 0

        nk = childProd.Products.Count
        ' a component CATIA has not loaded yet shows no content and often no
        ' part number: its structure is loaded now (visualization mode only),
        ' so it is listed like in the CATIA tree. Still no part number: a
        ' provisional row name - pass 1 loads the part and fills in the real one.
        If nk = 0 Then
            If Len(partNum) = 0 Then
                Call LoadStructure(childProd, partNum, nk)
            ElseIf IsUnloadedAssembly(childProd) Then
                Call LoadStructure(childProd, partNum, nk)
            End If
        End If
        If Len(partNum) = 0 And nk = 0 Then partNum = ProvisionalName(childProd)
        If partNum <> "" Then kids.Add Array(partNum, (nk > 0))

        ' The direct children of the root product ARE the first level; deeper
        ' nodes inherit the name of the first level branch they sit in.
        If currentLevel = 1 Then
            childFirst = FirstLevelLabel(childProd, partNum)
        Else
            childFirst = firstLevelName
        End If

        If nk > 0 Then

            ' ---------- SUB-ASSEMBLY ----------
            mAssyNodes.Add childProd
            ' the key is needed in any case: it is the parent key of everything
            ' below this node
            sKey = MakeRowKey(parentKey, childFirst, partNum)

            If mIncludeSubassemblies And partNum <> "" Then
                If dQty.Exists(sKey) Then
                    dQty(sKey) = dQty(sKey) + 1
                    Call AddFirstLevel(dFirst, sKey, childFirst)
                Else
                    Call AddBomRow(dQty, dRef, dDesc, dProps, dLevel, dPartNum, dFirst, dIsAssy, _
                                   sKey, childProd, partNum, childFirst, currentLevel, True)
                End If
            End If

            Set childLeaves = TraverseTree(childProd, dQty, dRef, dDesc, dProps, dLevel, _
                                           dBodies, dPartNum, dFirst, dIsAssy, dNodeLeaves, _
                                           colLeaves, currentLevel + 1, childFirst, sKey)

            ' the leaves below this node are only needed for an assembly
            ' screenshot - building them always is an O(leaves x depth) copy
            If mIncludeSubassemblies And mShotAssy And partNum <> "" Then
                If Not dNodeLeaves.Exists(sKey) Then dNodeLeaves.Add sKey, childLeaves
            End If

            If mIncludeSubassemblies And mShotAssy Then
                If Not childLeaves Is Nothing Then
                    For k = 1 To childLeaves.Count
                        myLeaves.Add childLeaves.Item(k)
                    Next k
                End If
            End If

        Else

            ' ---------- LEAF PART ----------
            ' Remember every leaf instance so all can be hidden in one call
            colLeaves.Add childProd
            If mIncludeSubassemblies And mShotAssy Then myLeaves.Add childProd

            If partNum <> "" Then
                sKey = MakeRowKey(parentKey, childFirst, partNum)

                If dQty.Exists(sKey) Then
                    dQty(sKey) = dQty(sKey) + 1
                    Call AddFirstLevel(dFirst, sKey, childFirst)
                Else
                    Call AddBomRow(dQty, dRef, dDesc, dProps, dLevel, dPartNum, dFirst, dIsAssy, _
                                   sKey, childProd, partNum, childFirst, currentLevel, False)

                    ' NOTE: multi-body detection used to happen here. It needs a
                    ' loaded CATPart, so it forced the whole tree into memory
                    ' during the traversal. It is done in the row loop now.
                End If
            End If
        End If

        Set childProd = Nothing
        Call Pump
    Next i

    Call NoteAssemblyContent(ownPn, kids)
End Function

' ==============================================================================
' HELPER: KEY OF A BOM ROW - see the run mode at the top of the module
'   "PARENT" (default): the path of the parent assemblies, so every assembly
'                       lists its own content and a part that sits in several
'                       assemblies appears below each of them
'   "FIRST":            one row per first level branch
'   "GLOBAL":           one row per part number in the whole tree
' ==============================================================================
Function MakeRowKey(ByVal parentKey As String, ByVal firstLevel As String, _
                    ByVal partNum As String) As String
    Select Case UCase$(mGroupBy)
        Case "GLOBAL"
            MakeRowKey = partNum
        Case "FIRST"
            MakeRowKey = firstLevel & "|" & partNum
        Case Else
            MakeRowKey = parentKey & ">" & partNum
    End Select
End Function

' ==============================================================================
' HELPER: NAME OF A FIRST LEVEL NODE - as in the CATIA tree
' ==============================================================================
Function FirstLevelLabel(oProd As Product, ByVal partNum As String) As String
    Dim s As String

    s = ""
    If Len(partNum) > 0 Then s = TreeName(oProd, partNum)
    If Len(s) = 0 Then
        On Error Resume Next
        s = oProd.Name
        Err.Clear
        On Error GoTo 0
    End If
    FirstLevelLabel = s
End Function

' ==============================================================================
' HELPER: THE NAME OF A COMPONENT AS THE CATIA TREE SHOWS IT (TREE_NAME_FORMAT),
' without the instance name
' ==============================================================================
Function TreeName(oProd As Product, ByVal pn As String) As String
    Dim nm As String

    nm = TREE_NAME_FORMAT
    If InStr(1, nm, "{DEFINITION}", vbTextCompare) > 0 Then _
        nm = Replace(nm, "{DEFINITION}", ProductProp(oProd, "DEFINITION"), 1, -1, vbTextCompare)
    If InStr(1, nm, "{REVISION}", vbTextCompare) > 0 Then _
        nm = Replace(nm, "{REVISION}", ProductProp(oProd, "REVISION"), 1, -1, vbTextCompare)
    If InStr(1, nm, "{DESCRIPTION}", vbTextCompare) > 0 Then _
        nm = Replace(nm, "{DESCRIPTION}", ProductProp(oProd, "DESCRIPTION"), 1, -1, vbTextCompare)
    If InStr(1, nm, "{NOMENCLATURE}", vbTextCompare) > 0 Then _
        nm = Replace(nm, "{NOMENCLATURE}", ProductProp(oProd, "NOMENCLATURE"), 1, -1, vbTextCompare)
    nm = Replace(nm, "{PN}", pn, 1, -1, vbTextCompare)
    nm = Trim$(nm)
    If Len(nm) = 0 Then nm = pn
    TreeName = nm
End Function

' one property of a component - "[?]" when CATIA cannot read it, like the tree
Private Function ProductProp(oProd As Product, ByVal which As String) As String
    Dim v As String

    v = ""
    On Error Resume Next
    Err.Clear
    Select Case which
        Case "DEFINITION"
            v = oProd.Definition
        Case "REVISION"
            v = oProd.Revision
        Case "DESCRIPTION"
            v = oProd.DescriptionRef
        Case "NOMENCLATURE"
            v = oProd.Nomenclature
    End Select
    If Err.Number <> 0 Then v = "[?]"
    Err.Clear
    On Error GoTo 0
    ProductProp = v
End Function

' the text of a row in the column "CATIA tree": name + instance, and how many
' instances the row stands for
Private Function TreeRowText(oProd As Product, ByVal pn As String, ByVal qty As Long) As String
    Dim inst As String, t As String

    t = TreeName(oProd, pn)
    inst = ""
    On Error Resume Next
    inst = oProd.Name
    Err.Clear
    On Error GoTo 0
    If Len(inst) > 0 Then
        If qty > 1 Then
            t = t & " (" & inst & " + " & (qty - 1) & " more)"
        Else
            t = t & " (" & inst & ")"
        End If
    End If
    TreeRowText = t
End Function

' the sheet name of a component: its tree name; when "<name>-<part number>" is
' longer than the 31 characters Excel allows, the name alone
Private Function SheetLabel(ByVal treeNm As String, ByVal pn As String) As String
    Dim lbl As String, k As Long

    lbl = treeNm
    If Len(lbl) > 31 And Len(pn) > 0 Then
        k = InStrRev(lbl, "-" & pn, -1, vbTextCompare)
        If k > 1 Then lbl = Left$(lbl, k - 1)
    End If
    If Len(Trim$(lbl)) = 0 Then lbl = pn
    SheetLabel = lbl
End Function

' every first level node with all the properties a tree name can come from -
' to check which properties the CATIA tree shows (see TREE_NAME_FORMAT)
Private Sub LogTreeNames(rootProd As Product)
    Dim i As Long, n As Long, child As Product, pn As String

    On Error Resume Next
    n = 0
    n = rootProd.Products.Count
    Call LogLine("first level nodes (part number | Definition | Revision | DescriptionRef | Nomenclature | instance -> column CATIA tree):")
    For i = 1 To n
        If i > 60 Then
            Call LogLine("  ... " & (n - 60) & " more")
            Exit For
        End If
        Set child = Nothing
        Set child = rootProd.Products.Item(i)
        If Not child Is Nothing Then
            pn = ""
            pn = child.PartNumber
            Err.Clear
            Call LogLine("  " & pn & " | " & ProductProp(child, "DEFINITION") & " | " & _
                         ProductProp(child, "REVISION") & " | " & ProductProp(child, "DESCRIPTION") & " | " & _
                         ProductProp(child, "NOMENCLATURE") & " | " & child.Name & "  ->  " & _
                         TreeRowText(child, pn, 1))
            Err.Clear
        End If
    Next i
    Err.Clear
End Sub

' ==============================================================================
' HELPER: REGISTER ONE BOM ROW (sub-assembly or part)
' ==============================================================================
Sub AddBomRow(dQty As Object, dRef As Object, dDesc As Object, dProps As Object, _
              dLevel As Object, dPartNum As Object, dFirst As Object, dIsAssy As Object, _
              ByVal sKey As String, oProd As Product, ByVal partNum As String, _
              ByVal firstLevel As String, ByVal lvl As Integer, ByVal isAssembly As Boolean)

    Dim dMass As Double, dVol As Double, dArea As Double
    Dim sDesc As String

    dQty.Add sKey, 1
    dRef.Add sKey, oProd
    dLevel.Add sKey, lvl
    dPartNum.Add sKey, partNum
    dFirst.Add sKey, firstLevel
    dIsAssy.Add sKey, isAssembly

    sDesc = ""
    On Error Resume Next
    sDesc = oProd.DescriptionRef
    Err.Clear
    On Error GoTo 0
    dDesc.Add sKey, sDesc

    ' Mass / volume / area are NOT read here: during the traversal the CATParts
    ' may still be unloaded, and then every value would come back as 0. They are
    ' read in the row loop, after the node has been promoted to design mode.
    dProps.Add sKey, Array(0#, 0#, 0#)
End Sub

' ==============================================================================
' HELPER: COLLECT THE FIRST LEVEL NAMES OF A ROW
' Only used in the part list ("GLOBAL"): one row can then belong to several
' first level branches, which are listed comma separated.
' ==============================================================================
Sub AddFirstLevel(dFirst As Object, ByVal sKey As String, ByVal firstLevel As String)
    Dim cur As String

    If UCase$(mGroupBy) <> "GLOBAL" Then Exit Sub
    If Len(firstLevel) = 0 Then Exit Sub
    If Not dFirst.Exists(sKey) Then Exit Sub

    cur = CStr(dFirst(sKey))
    If Len(cur) = 0 Then
        dFirst(sKey) = firstLevel
    ElseIf InStr(1, ", " & cur & ", ", ", " & firstLevel & ", ", vbTextCompare) = 0 Then
        dFirst(sKey) = cur & ", " & firstLevel
    End If
End Sub

' ==============================================================================
' HELPER: MASS / VOLUME / AREA OF A PART
' Volume in m3, area in m2 - same units as before. (Assembly rows are added up
' from their parts instead - see RollUp.)
' ==============================================================================
Sub ReadProductProps(oProd As Product, ByRef dMass As Double, ByRef dVol As Double, _
                     ByRef dArea As Double)
    Dim oAnalyze As Object
    Dim oInertia As Object

    dMass = 0: dVol = 0: dArea = 0

    On Error Resume Next

    Set oAnalyze = oProd.Analyze
    If Not oAnalyze Is Nothing Then
        dMass = oAnalyze.Mass
        dVol = oAnalyze.Volume / (1000# ^ 3)
        dArea = oAnalyze.WetArea / (1000# ^ 2)
    End If

    If dMass <= 0.0000001 Then
        Set oInertia = oProd.ReferenceProduct.GetTechnologicalObject("Inertia")
        If Not oInertia Is Nothing Then dMass = oInertia.Mass
        Set oInertia = Nothing
    End If

    Set oAnalyze = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

' ==============================================================================
' HELPER: LIST NON-EMPTY (SOLID) BODY NAMES OF A PART
' Returns the count; fills names() with the body names
' ==============================================================================
Function GetSolidBodyNames(oPart As Part, ByRef names() As String) As Integer
    On Error Resume Next
    Dim i As Integer, n As Integer
    Dim oBody As Body

    GetSolidBodyNames = 0
    n = 0
    If oPart Is Nothing Then Exit Function
    If oPart.Bodies.Count = 0 Then Exit Function

    ReDim names(oPart.Bodies.Count - 1)
    For i = 1 To oPart.Bodies.Count
        Set oBody = oPart.Bodies.Item(i)
        If Not oBody Is Nothing Then
            If oBody.Shapes.Count > 0 Then
                names(n) = oBody.Name
                n = n + 1
            End If
        End If
        Set oBody = Nothing
    Next i

    If n > 0 Then
        ReDim Preserve names(n - 1)
    End If
    GetSolidBodyNames = n
    Err.Clear
End Function

' ==============================================================================
' HELPER: GET A BODY BY NAME (body names are unique inside a part)
' ==============================================================================
Function GetBodyByName(oPart As Part, sName As String) As Body
    On Error Resume Next
    Set GetBodyByName = Nothing
    Set GetBodyByName = oPart.Bodies.Item(sName)
    Err.Clear
End Function

' ==============================================================================
' HELPER: BODY MASS & DENSITY via SPAWorkbench Inertias
' (GetTechnologicalObject("Inertia") only exists on Products, not on Bodies -
'  Inertias.Add(body) is the documented way and works on all V5 releases)
' ==============================================================================
Sub GetBodyMassAndDensity(oPartDoc As PartDocument, oBody As Body, _
                          ByRef dMass As Double, ByRef dDens As Double)
    On Error Resume Next
    Dim oSPA As Object
    Dim oInertia As Object

    Set oSPA = oPartDoc.GetWorkbench("SPAWorkbench")
    If oSPA Is Nothing Then Exit Sub

    Err.Clear
    Set oInertia = oSPA.Inertias.Add(oBody)
    If Err.Number = 0 And Not oInertia Is Nothing Then
        dMass = oInertia.Mass
        dDens = oInertia.Density
        Set oInertia = Nothing
    End If
    Set oSPA = Nothing
    Err.Clear
End Sub

' ==============================================================================
' HELPER: BODY VOLUME & AREA via SPA Workbench Measurable
' Measurable returns Volume in m3 and Area in m2 (same units as part rows)
' ==============================================================================
Sub GetBodyVolumeArea(oPartDoc As PartDocument, oPart As Part, oBody As Body, _
                      ByRef dVol As Double, ByRef dArea As Double)
    On Error Resume Next
    Dim oSPA As Object
    Dim oMeas As Object
    Dim oRef As Reference

    Set oSPA = oPartDoc.GetWorkbench("SPAWorkbench")
    If oSPA Is Nothing Then Exit Sub

    Set oRef = oPart.CreateReferenceFromObject(oBody)
    If oRef Is Nothing Then Exit Sub

    Set oMeas = oSPA.GetMeasurable(oRef)
    If Not oMeas Is Nothing Then
        dVol = oMeas.Volume
        dArea = oMeas.Area
        Set oMeas = Nothing
    End If

    Set oRef = Nothing
    Set oSPA = Nothing
    Err.Clear
End Sub

' ==============================================================================
' HELPER: BODY MATERIAL (body first, part level as fallback)
' ==============================================================================
Sub GetBodyMaterial(oPart As Part, oBody As Body, ByRef matName As String, ByRef density As Double)
    On Error Resume Next

    Dim oManager As Object
    Set oManager = oPart.GetItem("CATMatManagerVBExt")

    If Not oManager Is Nothing Then
        Dim oMat As Object

        oManager.GetMaterialOnBody oBody, oMat

        If oMat Is Nothing Then
            oManager.GetMaterialOnPart oPart, oMat
        End If

        If Not oMat Is Nothing Then
            matName = oMat.Name

            If oMat.ExistAnalysisData = 1 Then
                Dim oAnalysisMat As Object
                Set oAnalysisMat = oMat.AnalysisMaterial
                If Not oAnalysisMat Is Nothing Then
                    density = oAnalysisMat.GetValue("SAMDensity")
                    Set oAnalysisMat = Nothing
                End If
            End If
            Set oMat = Nothing
        End If
        Set oManager = Nothing
    End If
    Err.Clear
End Sub

' ==============================================================================
' HELPER: TRY GET PART IF ALREADY LOADED (NO OPENING)
' ==============================================================================
Private Function TryGetLoadedPart(oProd As Product, ByRef oPart As Part, ByRef oPartDoc As PartDocument) As Boolean
    On Error Resume Next
    TryGetLoadedPart = False
    Set oPart = Nothing
    Set oPartDoc = Nothing

    If oProd Is Nothing Then Exit Function
    Dim refProd As Product
    Set refProd = oProd.ReferenceProduct
    If refProd Is Nothing Then Exit Function

    If TypeName(refProd.Parent) = "PartDocument" Then
        Set oPartDoc = refProd.Parent
        Set oPart = oPartDoc.Part
        If Not oPart Is Nothing Then TryGetLoadedPart = True
    End If
    On Error GoTo 0
End Function

' ==============================================================================
' HELPER: THE CATPART OF A PART ROW, ALSO WHEN ApplyWorkMode DID NOT BRING IT
' 1. the reference is a loaded CATPart (the usual case after design mode)
' 2. the node's shape representation - for a CATProduct whose geometry is a
'    CATPart, or when design mode did not load it; this loads that ONE file
' ==============================================================================
Private Function LoadPartDocument(oProd As Product, ByRef oPart As Part, _
                                  ByRef oPartDoc As PartDocument) As Boolean
    Dim d As Object, ok As Boolean

    ok = TryGetLoadedPart(oProd, oPart, oPartDoc)
    If Not ok Then
        On Error Resume Next
        Set d = Nothing
        Set d = oProd.ReferenceProduct.GetMasterShapeRepresentation(True)
        If d Is Nothing Then
            Err.Clear
            Set d = oProd.GetMasterShapeRepresentation(True)
        End If
        Err.Clear
        If Not d Is Nothing Then
            If TypeName(d) = "PartDocument" Then
                Set oPartDoc = d
                Set oPart = oPartDoc.Part
                ok = Not (oPart Is Nothing)
            End If
        End If
        Err.Clear
        On Error GoTo 0
    End If
    LoadPartDocument = ok
End Function

' what CATIA has of a part that gave no values - for the log
Private Function PartStateText(oProd As Product, ByVal hasPart As Boolean) As String
    Dim tn As String, pth As String, hasShape As String

    On Error Resume Next
    tn = "?"
    tn = TypeName(oProd.ReferenceProduct.Parent)
    Err.Clear
    hasShape = "?"
    hasShape = CStr(oProd.HasAMasterShapeRepresentation)
    Err.Clear
    pth = "?"
    pth = CStr(oProd.GetMasterShapeRepresentationPathName)
    Err.Clear
    On Error GoTo 0
    PartStateText = "reference document = " & tn & ", shape = " & hasShape & _
                    ", shape file = " & pth & ", CATPart loaded = " & CStr(hasPart)
End Function

' ==============================================================================
' HELPER: GET MATERIAL & DENSITY (Using VBExt) - PART LEVEL
' ==============================================================================
Sub GetMaterialAndDensity(oPart As Part, ByRef matName As String, ByRef density As Double)
    On Error Resume Next

    Dim oManager As Object
    Set oManager = oPart.GetItem("CATMatManagerVBExt")

    If Not oManager Is Nothing Then
        Dim oMat As Object
        Dim oBody As Body

        Set oBody = oPart.MainBody
        oManager.GetMaterialOnBody oBody, oMat

        If oMat Is Nothing Then
            oManager.GetMaterialOnPart oPart, oMat
        End If

        If Not oMat Is Nothing Then
            matName = oMat.Name

            If oMat.ExistAnalysisData = 1 Then
                Dim oAnalysisMat As Object
                Set oAnalysisMat = oMat.AnalysisMaterial
                If Not oAnalysisMat Is Nothing Then
                    density = oAnalysisMat.GetValue("SAMDensity")
                    Set oAnalysisMat = Nothing
                End If
            End If
            Set oMat = Nothing
        End If
        Set oBody = Nothing
        Set oManager = Nothing
    End If

    If density <= 0 Then
        Dim oInertia As Object
        Set oInertia = oPart.Inertia
        If Not oInertia Is Nothing Then
            density = oInertia.Density
            Set oInertia = Nothing
        End If
    End If
    Err.Clear
End Sub

' ==============================================================================
' HELPER: BOUNDING BOX (whole part)
' ==============================================================================
Sub GetBoundingBoxDims(oPart As Part, ByRef dDims() As Double)
    On Error Resume Next
    Dim foundValidBox As Boolean: foundValidBox = False
    Dim oDocP As PartDocument
    Dim nHB As Long, iHB As Long

    If UCase$(MEASURE_MODE) <> "BOX" Then Exit Sub
    If mBBoxDisabled Then Exit Sub

    If Not oPart.MainBody Is Nothing Then
        foundValidBox = MeasureTarget(oPart, oPart.MainBody, dDims)
    End If

    ' The geometrical sets are probed by INDEX over the count taken BEFORE the
    ' loop: MeasureTarget adds a temporary set itself, and a For Each over the
    ' live collection would then walk into the sets it is creating.
    If foundValidBox = False And Not mBBoxDisabled Then
        nHB = oPart.HybridBodies.Count
        For iHB = 1 To nHB
            If InStr(1, oPart.HybridBodies.Item(iHB).Name, "TMP_BOM_MEASURE", vbTextCompare) = 0 Then
                foundValidBox = MeasureTarget(oPart, oPart.HybridBodies.Item(iHB), dDims)
            End If
            If foundValidBox Or mBBoxDisabled Then Exit For
        Next iHB
    End If

    ' Version-proof fallback: extremum-based box on the main body
    If foundValidBox = False And Not mBBoxDisabled Then
        Set oDocP = oPart.Parent
        If Not oPart.MainBody Is Nothing Then
            foundValidBox = MeasureAABBExtremum(oPart, oDocP, oPart.MainBody, dDims)
        End If
    End If
    Err.Clear
End Sub

' ==============================================================================
' HELPER: BOUNDING BOX ON ANY TARGET (MainBody, Body, HybridBody)
' The temporary box is created and deleted inside the PART document
' ==============================================================================
Function MeasureTarget(oPart As Part, oTargetObj As Object, ByRef dDims() As Double) As Boolean
    On Error Resume Next
    Err.Clear
    MeasureTarget = False

    ' this path creates geometry inside the user's CATPart - only on request,
    ' and never again once it has failed one time (see MEASURE_MODE)
    If UCase$(MEASURE_MODE) <> "BOX" Then Exit Function
    If mBBoxDisabled Then Exit Function

    Dim oHSF As Object: Set oHSF = oPart.HybridShapeFactory
    Dim oRef As Reference: Set oRef = oPart.CreateReferenceFromObject(oTargetObj)
    Dim oBox As Object: Set oBox = oHSF.AddNewBoundingBox(oRef)
    Dim oPartSel As Selection: Set oPartSel = oPart.Parent.Selection
    Dim oTmpSet As HybridBody
    Dim d1 As Double, d2 As Double, d3 As Double

    If oBox Is Nothing Then GoTo Cleanup

    oBox.Type = 1

    ' Aggregate the box into a temporary geometrical set. Without this,
    ' UpdateObject only succeeds when the target is the in-work body.
    Err.Clear
    Set oTmpSet = oPart.HybridBodies.Add()
    If Not oTmpSet Is Nothing Then
        oTmpSet.Name = "TMP_BOM_MEASURE"
        oTmpSet.AppendHybridShape oBox
        oPart.InWorkObject = oTmpSet
    End If

    Err.Clear
    oPart.UpdateObject oBox

    If Err.Number = 0 Then
        d1 = oBox.GetLength.Value: d2 = oBox.GetWidth.Value: d3 = oBox.GetHeight.Value
        If (d1 + d2 + d3) > 0.1 Then
            Call SortThree(d1, d2, d3, dDims)
            MeasureTarget = True
        End If
    Else
        ' CATIA has just shown its own modal panel. Switch the whole geometry
        ' path off so the user does not have to click one away per component.
        mBBoxDisabled = True
        Err.Clear
    End If

    ' Restore the in-work object BEFORE deleting - deleting the set while
    ' it is (or contains) the in-work object makes CATIA raise a modal
    ' "Selected element(s) not allowed for this operation" dialog that
    ' blocks the whole macro
    Err.Clear
    oPart.InWorkObject = oPart.MainBody

    ' Delete the temporary geometrical set (takes the box with it),
    ' or just the bare box if the set could not be created
    ' Remove the box through the FACTORY. Selection.Delete dispatches CATIA's
    ' interactive Delete command, and that is what raises the modal
    ' "Selected element(s) not allowed for this operation" panel per component.
    Err.Clear
    oHSF.DeleteObjectForDatum oPart.CreateReferenceFromObject(oBox)
    If Err.Number <> 0 Then
        ' could not remove it - hide it so it does not pollute the view
        Err.Clear
        oPartSel.Clear
        oPartSel.Add oBox
        oPartSel.VisProperties.SetShow 1
    End If
    oPartSel.Clear
    Err.Clear

Cleanup:
    Set oPartSel = Nothing
    Set oTmpSet = Nothing
    Set oBox = Nothing
    Set oRef = Nothing
    Set oHSF = Nothing
    Err.Clear
End Function

' ==============================================================================
' HELPER: TRUE AXIS-ALIGNED BOUNDING DIMS VIA EXTREMUM FEATURES
' Works on ALL V5 releases (AddNewBoundingBox only exists on newer ones).
' Creates temporary min/max Extremum points along X, Y, Z in a temporary
' geometrical set, reads their coordinates, then deletes everything.
' ==============================================================================
Function MeasureAABBExtremum(oPart As Part, oPartDoc As PartDocument, _
                             oTargetObj As Object, ByRef dDims() As Double) As Boolean
    On Error Resume Next
    Err.Clear
    MeasureAABBExtremum = False

    If UCase$(MEASURE_MODE) <> "BOX" Then Exit Function
    If mBBoxDisabled Then Exit Function

    Dim oHSF As Object: Set oHSF = oPart.HybridShapeFactory
    Dim oSPA As Object: Set oSPA = oPartDoc.GetWorkbench("SPAWorkbench")
    Dim oRef As Reference: Set oRef = oPart.CreateReferenceFromObject(oTargetObj)
    If oHSF Is Nothing Or oSPA Is Nothing Or oRef Is Nothing Then GoTo CleanupE

    Dim oTmpSet As HybridBody
    Err.Clear
    Set oTmpSet = oPart.HybridBodies.Add()
    If oTmpSet Is Nothing Then mBBoxDisabled = True: GoTo CleanupE
    oTmpSet.Name = "TMP_BOM_MEASURE"
    oPart.InWorkObject = oTmpSet

    Dim spans(2) As Double
    Dim k As Integer
    Dim okAll As Boolean: okAll = True
    Dim vMin As Double, vMax As Double
    Dim oDir As Object
    Dim dx As Double, dy As Double, dz As Double

    For k = 0 To 2
        dx = 0#: dy = 0#: dz = 0#
        If k = 0 Then dx = 1#
        If k = 1 Then dy = 1#
        If k = 2 Then dz = 1#

        Set oDir = oHSF.AddNewDirectionByCoord(dx, dy, dz)

        If GetExtremumCoord(oPart, oSPA, oTmpSet, oHSF, oRef, oDir, 1, k, vMax) And _
           GetExtremumCoord(oPart, oSPA, oTmpSet, oHSF, oRef, oDir, 0, k, vMin) Then
            spans(k) = vMax - vMin
        Else
            okAll = False
        End If

        Set oDir = Nothing
        If okAll = False Then mBBoxDisabled = True: Exit For
    Next k

    ' Restore the in-work object BEFORE deleting - deleting the set while
    ' it is the in-work object raises a modal CATIA error dialog that
    ' blocks the whole macro
    Err.Clear
    oPart.InWorkObject = oPart.MainBody

    ' Remove the temporary set through the FACTORY (no interactive Delete
    ' command, so no modal panel). If that is refused, hide it - it then stays
    ' in the part until it is closed without saving, the price of "BOX" mode.
    Dim oPartSel As Selection
    Err.Clear
    oHSF.DeleteObjectForDatum oPart.CreateReferenceFromObject(oTmpSet)
    If Err.Number <> 0 Then
        Err.Clear
        Set oPartSel = oPartDoc.Selection
        oPartSel.Clear
        oPartSel.Add oTmpSet
        oPartSel.VisProperties.SetShow 1
        oPartSel.Clear
        Set oPartSel = Nothing
    End If
    Err.Clear

    If okAll Then
        If (spans(0) + spans(1) + spans(2)) > 0.1 Then
            Call SortThree(spans(0), spans(1), spans(2), dDims)
            MeasureAABBExtremum = True
        End If
    End If

CleanupE:
    Set oTmpSet = Nothing
    Set oRef = Nothing
    Set oSPA = Nothing
    Set oHSF = Nothing
    Err.Clear
End Function

' ==============================================================================
' HELPER: BUILD ONE EXTREMUM POINT AND READ ITS COORDINATE ON ONE AXIS
' minMax: 1 = maximum, 0 = minimum ; axisIndex: 0 = X, 1 = Y, 2 = Z
' ==============================================================================
Private Function GetExtremumCoord(oPart As Part, oSPA As Object, oTmpSet As HybridBody, _
                                  oHSF As Object, oRef As Reference, oDir As Object, _
                                  ByVal minMax As Long, ByVal axisIndex As Integer, _
                                  ByRef outCoord As Double) As Boolean
    On Error Resume Next
    GetExtremumCoord = False

    Dim oExt As Object
    Err.Clear
    Set oExt = oHSF.AddNewExtremum(oRef, oDir, minMax)
    If oExt Is Nothing Then Err.Clear: Exit Function

    oTmpSet.AppendHybridShape oExt

    Err.Clear
    oPart.UpdateObject oExt
    If Err.Number <> 0 Then mBBoxDisabled = True: Err.Clear: Exit Function

    Dim oMeas As Object
    Dim oExtRef As Reference
    Set oExtRef = oPart.CreateReferenceFromObject(oExt)
    Set oMeas = oSPA.GetMeasurable(oExtRef)
    If oMeas Is Nothing Then Err.Clear: Exit Function

    Dim coords(2)
    Err.Clear
    oMeas.GetPoint coords
    If Err.Number <> 0 Then Err.Clear: Exit Function

    outCoord = CDbl(coords(axisIndex))
    GetExtremumCoord = True

    Set oMeas = Nothing
    Set oExtRef = Nothing
    Set oExt = Nothing
    Err.Clear
End Function

' ==============================================================================
' HELPER: BODY DIMENSIONS FROM ITS OWN INERTIA (last fallback)
' Gives the equivalent-box dimensions computed from the principal moments,
' using SPAWorkbench Inertias on the single body
' ==============================================================================
Sub GetBodyInertiaDims(oPartDoc As PartDocument, oBody As Body, ByRef dDims() As Double)
    On Error Resume Next
    Dim oSPA As Object
    Dim oInertia As Object

    Set oSPA = oPartDoc.GetWorkbench("SPAWorkbench")
    If oSPA Is Nothing Then Exit Sub

    Err.Clear
    Set oInertia = oSPA.Inertias.Add(oBody)
    If Err.Number <> 0 Or oInertia Is Nothing Then Err.Clear: Exit Sub

    Dim dMass As Double: dMass = oInertia.Mass
    If dMass <= 0 Then
        Set oInertia = Nothing
        Exit Sub
    End If
    Dim Matrix(8)
    oInertia.GetPrincipalMoments Matrix
    Dim A As Double, B As Double, C As Double
    A = (6 * (Matrix(1) + Matrix(2) - Matrix(0)) / dMass)
    B = (6 * (Matrix(0) + Matrix(2) - Matrix(1)) / dMass)
    C = (6 * (Matrix(0) + Matrix(1) - Matrix(2)) / dMass)
    If A < 0 Then A = 0
    If B < 0 Then B = 0
    If C < 0 Then C = 0
    Call SortThree(Sqr(A) * 1000, Sqr(B) * 1000, Sqr(C) * 1000, dDims)
    Set oInertia = Nothing
    Err.Clear
End Sub

' ==============================================================================
' HELPER: DIMENSIONS FROM INERTIA (fallback when bounding box fails)
' ==============================================================================
Sub GetInertiaDims(oProd As Product, ByRef dDims() As Double)
    On Error Resume Next
    Dim oInertia As Object: Set oInertia = oProd.GetTechnologicalObject("Inertia")
    If oInertia Is Nothing Then Exit Sub
    Dim dMass As Double: dMass = oInertia.Mass
    If dMass <= 0 Then
        Set oInertia = Nothing
        Exit Sub
    End If
    Dim Matrix(8)
    oInertia.GetPrincipalMoments Matrix
    Dim A As Double, B As Double, C As Double
    A = (6 * (Matrix(1) + Matrix(2) - Matrix(0)) / dMass)
    B = (6 * (Matrix(0) + Matrix(2) - Matrix(1)) / dMass)
    C = (6 * (Matrix(0) + Matrix(1) - Matrix(2)) / dMass)
    If A < 0 Then A = 0
    If B < 0 Then B = 0
    If C < 0 Then C = 0
    Call SortThree(Sqr(A) * 1000, Sqr(B) * 1000, Sqr(C) * 1000, dDims)
    Set oInertia = Nothing
    Err.Clear
End Sub

' ==============================================================================
' HELPER: SORT THREE VALUES (largest to smallest)
' ==============================================================================
Sub SortThree(a As Double, b As Double, c As Double, ByRef arr() As Double)
    Dim temp As Double
    Dim vals(2) As Double
    vals(0) = a: vals(1) = b: vals(2) = c
    Dim i As Integer, j As Integer
    For i = 0 To 1
        For j = i + 1 To 2
            If vals(i) < vals(j) Then
                temp = vals(i): vals(i) = vals(j): vals(j) = temp
            End If
        Next j
    Next i
    arr(0) = vals(0): arr(1) = vals(1): arr(2) = vals(2)
End Sub
