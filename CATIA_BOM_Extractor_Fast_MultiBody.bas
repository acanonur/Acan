Option Explicit

' ==============================================================================
' MACRO: MASTER BOM v5 (SPLIT BY LEVEL + NUMBERS FIRST, PICTURES LAST,
'                       NO LONG FREEZES, OWN EXCEL, AUTO-SAVE AND LOG)
'
' HOW A RUN GOES:
'   1. You are asked at which level the BOM is split. The question shows how
'      many components and sub-assemblies every level of the product has.
'        0 = no split: one sheet for the whole product
'        n = one worksheet per sub-assembly on level n. Each one is opened in a
'            window of its own, so CATIA only has to draw that sub-assembly while
'            it is measured and pictured. Parts that are not inside any of them
'            get one extra sheet, and the "Index" sheet links to every sheet and
'            shows its level, quantity, path, number of rows, status and time.
'   2. You are asked whether pictures are wanted. "No" is the fastest run.
'   3. Every sheet is filled in TWO PASSES:
'        pass 1 - all numbers: the display is frozen and nothing is shown or
'                 hidden, CATIA only loads and measures the PARTS. The assembly
'                 rows are added up from their parts afterwards. The sheet is
'                 complete and visible in Excel before the first picture.
'        pass 2 - all pictures: one show / capture / hide per row, every picture
'                 goes into Excel straight away.
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
'     the row belongs to, so the BOM can be filtered per first level
'   - Column O "Type": Assembly / Part / Body
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

' False = column N shows the part number of the first level node,
' True  = column N shows its description
Private Const FIRST_LEVEL_USE_DESCRIPTION As Boolean = False

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
Private Const ASK_SPLIT_LEVEL As Boolean = True
' False = every component is done in the main window (no extra windows).
' A component that is not the root of its own CATProduct file (e.g. a
' "Product1" inserted inside another assembly) is always done there.
Private Const OPEN_UNITS_IN_NEW_WINDOW As Boolean = True

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
Private mBoldRows As Collection  ' sheet rows written in bold (assemblies)
Private mAssyRows As Collection  ' Array(sheet row, part number, product) of every assembly row
Private mSheetName As String     ' name of the sheet being written

' --- the whole run -------------------------------------------------------------
Private mXlApp As Object, mXlBook As Object, mFso As Object
Private mOwnXl As Boolean        ' mXlApp was started by this run
Private mXlLost As Boolean       ' Excel was closed or crashed - nothing can be written any more
Private mXlSkipAt As Double      ' when the user last chose to skip a write to a busy Excel (-1 = never)
Private mPromoted As Object      ' part numbers already promoted to design mode in this run
Private mTempDir As String
Private mSheetNo As Long
Private mSheetNames As Object    ' every sheet name used in the workbook
Private mShotParts As Boolean    ' pictures of the part (and body) rows in this run
Private mShotAssy As Boolean     ' pictures of the assembly rows in this run
Private mLogPath As String, mSavePath As String
Private mSavedOnce As Boolean, mLastSave As Double
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

    startTime = Timer
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
    Set mXlApp = Nothing                 ' never an Excel of an earlier run
    Set mXlBook = Nothing
    mOwnXl = False
    mXlLost = False
    mXlSkipAt = -1
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
    On Error Resume Next
    If Not OWN_EXCEL_INSTANCE Then Set mXlApp = GetObject(, "Excel.Application")
    Err.Clear
    If mXlApp Is Nothing Then
        Set mXlApp = CreateObject("Excel.Application")
        mOwnXl = Not (mXlApp Is Nothing)
    End If
    Err.Clear
    On Error GoTo Fail
    If mXlApp Is Nothing Then Err.Raise vbObjectError + 513, , "Excel could not be started"

    mXlApp.Visible = True
    On Error Resume Next
    ' an Excel started by the macro stays open when the macro ends
    If mOwnXl Then mXlApp.UserControl = True
    Err.Clear
    On Error GoTo Fail
    Set mXlBook = mXlApp.Workbooks.Add
    Call KeepOneSheet(mXlBook)
    Set firstSheet = mXlBook.Sheets(1)

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
        If Not unitOK Then MsgBox "The run stopped early:" & vbCrLf & statusTxt, vbExclamation

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
        Call WriteIndexHeader(idxSheet, splitLevel)
        idxRow = 3

        ' --- the components that are CATProducts of their own: own window ------
        For Each unitKey In dUnitQty.Keys
            nUnits = nUnits + 1
            unitStart = Timer
            nRows = 0: nAssy = 0: nParts = 0
            statusTxt = "not started"
            Set xlSheet = Nothing
            Call CheckExcel                  ' Excel gone: stop the whole run here
            ' an error in here ends this component only, never the whole run
            On Error GoTo UnitFail

            Set unitInst = dUnitRef(unitKey)
            Call LogLine("=== component " & nUnits & " of " & dUnitQty.Count & ": " & CStr(unitKey))
            Set xlSheet = AddSheet(mXlBook, Format(nUnits, "00") & " " & CStr(unitKey))
            If xlSheet Is Nothing Then Err.Raise vbObjectError + 514, , "Excel could not add a sheet"

            Call WriteIndexRow(idxSheet, idxRow, nUnits, CStr(unitKey), CStr(dUnitDesc(unitKey)), _
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

                Call WriteIndexResult(idxSheet, idxRow, nRows, statusTxt, Elapsed(unitStart), "own window")
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
                Call WriteIndexRow(idxSheet, idxRow, nUnits, CStr(unitKey), CStr(dUnitDesc(unitKey)), _
                                   splitLevel, CLng(dUnitQty(unitKey)), CStr(dUnitFirst(unitKey)), _
                                   CStr(dUnitPath(unitKey)), "")
            End If
            Call WriteIndexResult(idxSheet, idxRow, nRows, statusTxt, Elapsed(unitStart), "-")
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

                Call WriteIndexResult(idxSheet, CLng(colFbIdxRow(i)), nRows, statusTxt, _
                                      Elapsed(unitStart), "main window")
                Call LogLine("component " & CStr(unitKey) & ": " & statusTxt & ", " & nRows & " rows, " & _
                             Format(Elapsed(unitStart), "0") & " s")
                totRows = totRows + nRows: totAssy = totAssy + nAssy: totParts = totParts + nParts
                If unitOK Then nUnitsOK = nUnitsOK + 1
                Call SaveWorkbook(False)
                DoEvents
            Next i

            If colLooseInst.Count > 0 Then
                Call CheckExcel
                unitStart = Timer
                Call LogLine("=== above level " & splitLevel)
                Set xlSheet = AddSheet(mXlBook, "Above level " & splitLevel)
                If Not xlSheet Is Nothing Then
                    Call WriteIndexRow(idxSheet, idxRow, 0, "(above level " & splitLevel & ")", _
                                       "parts that are not inside any component of level " & splitLevel & _
                                       IIf(mIncludeSubassemblies, ", and the assemblies above it", ""), _
                                       0, colLooseParts.Count, "", "", mSheetName)
                    unitOK = ProcessLoose(colLooseInst, colLooseLevel, colLooseFirst, colLooseKey, _
                                          colLooseIsAssy, rootSel, rootViewer, xlSheet, _
                                          nRows, nAssy, nParts, statusTxt)
                    Call WriteIndexResult(idxSheet, idxRow, nRows, statusTxt, Elapsed(unitStart), "main window")
                    Call LogLine("above level " & splitLevel & ": " & statusTxt & ", " & nRows & " rows, " & _
                                 Format(Elapsed(unitStart), "0") & " s")
                    totRows = totRows + nRows: totAssy = totAssy + nAssy: totParts = totParts + nParts
                End If
            End If

            ' everything visible again
            Call SetShowCollection(rootSel, colUnitInst, 0)
            Call SetShowCollection(rootSel, colLooseParts, 0)
            rootSceneSet = False
        End If

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
        MsgBox "BOM exported  (" & modeName & ")" & vbCrLf & vbCrLf & _
               "Rows written:     " & totRows & vbCrLf & _
               "  sub-assemblies: " & totAssy & vbCrLf & _
               "  parts:          " & totParts & vbCrLf & _
               "Time: " & Format(Elapsed(startTime) / 60#, "0.0") & " min" & vbCrLf & vbCrLf & _
               SavedText(), vbInformation
    Else
        MsgBox "BOM exported  (" & modeName & ", split at level " & splitLevel & ")" & vbCrLf & vbCrLf & _
               "Component sheets: " & nUnits & "  (" & nUnitsOK & " complete)" & vbCrLf & _
               "Rows written:     " & totRows & vbCrLf & _
               "Time: " & Format(Elapsed(startTime) / 60#, "0.0") & " min" & vbCrLf & vbCrLf & _
               "The Index sheet links to every component sheet." & vbCrLf & vbCrLf & _
               SavedText(), vbInformation
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

    MsgBox failMsg & vbCrLf & vbCrLf & _
           "Visibility, spec tree and compass were restored." & vbCrLf & _
           "The rows written so far are in the Excel workbook." & vbCrLf & vbCrLf & _
           SavedText(), vbCritical

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

    ProcessTree = False
    nRows = 0: nAssy = 0: nParts = 0
    r = 2
    statusTxt = ""
    t0 = Timer
    On Error GoTo SheetFail

    Call ResetSheetData
    Call WriteHeader(xlSheet)

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
    Call FlushSheet(xlSheet)

    ' the numbers are complete - make them visible before the pictures start
    Call FinishSheetLayout(xlSheet)
    Call LogLine(mSheetName & ": numbers done after " & Format(Elapsed(t0), "0") & " s")
    Call SaveWorkbook(False)

    ' --- PASS 2: the pictures ---
    Call CheckExcel
    Call TakePictures(sel, oViewer, xlSheet)

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
    If Not mFlushed Then Call FlushSheet(xlSheet)
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

    ProcessLoose = False
    nRows = 0: nAssy = 0: nParts = 0
    r = 2
    statusTxt = ""
    On Error GoTo LooseFail

    Call ResetSheetData
    Call WriteHeader(xlSheet)

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
    Call FlushSheet(xlSheet)
    Call FinishSheetLayout(xlSheet)
    Call SaveWorkbook(False)
    Call CheckExcel
    Call TakePictures(sel, oViewer, xlSheet)

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
    If Not mFlushed Then Call FlushSheet(xlSheet)
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

        If isAssy Then
            ' ---------- SUB-ASSEMBLY: nothing is loaded or measured here ----------
            ' (ApplyWorkMode / Analyze on an assembly work on its WHOLE subtree in
            '  one blocking call - the values are added up from its parts instead)
            nAssy = nAssy + 1
            Call BufSet(r, 15, "Assembly")
            mBoldRows.Add r
            mAssyRows.Add Array(r, strPartNum, oPartProd)
            If mShotAssy Then
                Call AddJob(r, "ASSY", CStr(uniqueKey), "")
            Else
                Call BufSet(r, 1, "-")
            End If

        Else
            ' ---------- PART ----------
            nParts = nParts + 1
            Call BufSet(r, 15, "Part")

            Call SetStatus("BOM " & mSheetName & ": numbers, row " & nDone & " of " & _
                           nTotal & " - " & strPartNum)
            ' written BEFORE the part is loaded: if CATIA hangs, the last line of
            ' the log names the part
            Call LogLine(mSheetName & "  row " & nDone & "/" & nTotal & "  " & strPartNum)
            tRow = Timer

            ' promote THIS part to design mode first - mass, volume, area and the
            ' inertia box are all 0 on a component in visualization mode. The
            ' CATPart document is shared by all instances of the part, so once
            ' per part number is enough.
            If Not mPromoted.Exists(strPartNum) Then
                On Error Resume Next
                oPartProd.ApplyWorkMode 2
                Err.Clear
                On Error GoTo 0
                mPromoted.Add strPartNum, True
            End If
            tLoad = Elapsed(tRow)

            hasPart = False
            Set oPart = Nothing
            Set oPartDoc = Nothing
            hasPart = TryGetLoadedPart(oPartProd, oPart, oPartDoc)

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
                    Err.Clear
                    On Error GoTo 0
                End If
                If dims(0) > 0 Then mDimCache.Add strPartNum, Array(dims(0), dims(1), dims(2))
            End If
            Call PutNumber(r, 9, dims(0))
            Call PutNumber(r, 10, dims(1))
            Call PutNumber(r, 11, dims(2))

            ' --- the picture of this row is taken in pass 2 ---
            If mShotParts Then
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
                        Call BufSet(r, 4, "Body of " & strPartNum)
                        Call BufSet(r, 5, mQty(uniqueKey))
                        Call BufSet(r, 14, mFirst(uniqueKey))
                        Call BufSet(r, 15, "Body")

                        bMass = 0: bVol = 0: bArea = 0
                        bDens = 0: bDensInertia = 0
                        bMat = "N/A"

                        Call GetBodyMaterial(oPart, oBody, bMat, bDens)
                        Call GetBodyMassAndDensity(oPartDoc, oBody, bMass, bDensInertia)
                        If bDens <= 0 Then bDens = bDensInertia
                        Call GetBodyVolumeArea(oPartDoc, oPart, oBody, bVol, bArea)
                        If bMass <= 0 And bDens > 0 And bVol > 0 Then bMass = bVol * bDens

                        Call PutNumber(r, 6, bMass)
                        Call PutNumber(r, 7, bVol)
                        Call PutNumber(r, 8, bArea)

                        dims(0) = 0: dims(1) = 0: dims(2) = 0
                        If UCase$(MEASURE_MODE) = "BOX" Then
                            If Not MeasureTarget(oPart, oBody, dims) Then
                                Call MeasureAABBExtremum(oPart, oPartDoc, oBody, dims)
                            End If
                        End If
                        If dims(0) < 0.1 Then Call GetBodyInertiaDims(oPartDoc, oBody, dims)
                        Call PutNumber(r, 9, dims(0))
                        Call PutNumber(r, 10, dims(1))
                        Call PutNumber(r, 11, dims(2))

                        Call BufSet(r, 12, bMat)
                        Call PutNumber(r, 13, bDens)

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
        ElseIf ASSY_DIMS_MAX_PARTS > 0 And v(3) > 0 And v(3) <= ASSY_DIMS_MAX_PARTS Then
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

        Call BufSet(r, 12, "N/A")
        Call BufSet(r, 13, "N/A")
        Set oProd = Nothing
        DoEvents
    Next it
End Sub


' ==============================================================================
' HELPER: MASS / VOLUME / AREA OF ONE PIECE OF AN ASSEMBLY, ADDED UP
' Returns Array(mass, volume, area, part instances, part instances without values)
' over every part instance below it. Pure bookkeeping - nothing is asked from
' CATIA. The content of every assembly part number is noted while the trees are
' walked (mAssyKids); the part values are the ones pass 1 has read (mPropCache).
' ==============================================================================
Private Function RollUp(ByVal pn As String, ByVal depth As Long) As Variant
    Dim kids As Collection, k As Variant, v As Variant, res As Variant
    Dim m As Double, vo As Double, a As Double
    Dim np As Long, nm As Long

    If mRollCache.Exists(pn) Then
        RollUp = mRollCache(pn)
        Exit Function
    End If
    If depth > 60 Or Not mAssyKids.Exists(pn) Then
        RollUp = Array(0#, 0#, 0#, 0&, 1&)          ' content unknown
        Exit Function
    End If

    Set kids = mAssyKids(pn)
    For Each k In kids
        If k(1) Then
            v = RollUp(CStr(k(0)), depth + 1)
            m = m + v(0): vo = vo + v(1): a = a + v(2)
            np = np + v(3): nm = nm + v(4)
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

    res = Array(m, vo, a, np, nm)
    ' a complete sum never changes again; an incomplete one may, later in the run
    If nm = 0 Then mRollCache.Add pn, res
    RollUp = res
End Function

' notes the content of an assembly part number once (all its instances share it)
Private Sub NoteAssemblyContent(ByVal pn As String, kids As Collection)
    If Len(pn) = 0 Then Exit Sub
    If mAssyKids Is Nothing Then Exit Sub
    If Not mAssyKids.Exists(pn) Then mAssyKids.Add pn, kids
End Sub



' ==============================================================================
' PASS 2 - ALL PICTURES OF ONE SHEET (after all its numbers are written)
' 1. scene: every node of the sheet visible, every leaf hidden
' 2. per picture: show -> reframe -> capture to a JPEG file -> hide, the file
'    goes into Excel at once and is deleted (Excel repaint off meanwhile)
' ==============================================================================
Private Sub TakePictures(sel As Selection, oViewer As Viewer, xlSheet As Object)
    Dim j As Long, rowNo As Long
    Dim kind As String, key As String, bodyName As String
    Dim picPath As String, what As String
    Dim oProd As Product, colNode As Collection
    Dim oPart As Part, oPartDoc As PartDocument, partSel As Selection
    Dim oBody As Body, bodyNames As Variant
    Dim colWasShown As Collection, colWasHidden As Collection
    Dim t As Double, hasFile As Boolean

    If mJobRow.Count = 0 Then Exit Sub

    If oViewer Is Nothing Then
        For j = 1 To mJobRow.Count
            Call XlPutCell(xlSheet, mJobRow(j), 1, "No Viewer")
        Next j
        Exit Sub
    End If

    Call LogLine(mSheetName & ": " & mJobRow.Count & " pictures")

    ' --- 1. the scene ---
    If PRESERVE_VISIBILITY Then
        Set colWasShown = New Collection
        Set colWasHidden = New Collection
        Call SplitByShowState(sel, mAllNodes, colWasShown, colWasHidden)
    End If
    ' every assembly visible (a hidden one hides its whole subtree), every part
    ' hidden - showing the parts first only to hide them again is wasted work
    Call SetShowCollection(sel, mAssyNodes, 0)
    Call SetShowCollection(sel, mLeaves, 1)
    Call Settle(100)

    On Error Resume Next
    mXlApp.ScreenUpdating = False
    Err.Clear
    On Error GoTo 0

    ' --- 2. per picture: capture into a file, put it into Excel, delete it ---
    '        (so there is only ever one picture file on disk)
    For j = 1 To mJobRow.Count
        Call CheckExcel                       ' no hour of pictures into a closed Excel
        rowNo = mJobRow(j)
        kind = mJobKind(j)
        key = mJobKey(j)
        bodyName = mJobBody(j)
        picPath = mTempDir & "\s" & mSheetNo & "_r" & rowNo & ".jpg"

        what = kind
        If Len(key) > 0 Then
            If mPartNum.Exists(key) Then what = kind & " " & mPartNum(key)
        End If
        If Len(bodyName) > 0 Then what = what & " / " & bodyName
        Call SetStatus("BOM " & mSheetName & ": picture " & j & " of " & mJobRow.Count & " - " & what)
        Call LogLine(mSheetName & "  picture " & j & "/" & mJobRow.Count & "  " & what)
        t = Timer

        ' a skipped job must never pick up an old file
        On Error Resume Next
        If mFso.FileExists(picPath) Then mFso.DeleteFile picPath, True
        Err.Clear
        On Error GoTo 0

        Select Case kind

            Case "UNIT"                     ' the whole component of the sheet
                If mLeaves.Count > 0 And mLeaves.Count <= MAX_ASSY_LEAVES_FOR_SHOT * 5 Then
                    Call SetShowCollection(sel, mLeaves, 0)
                    Call Settle(ASSY_SETTLE_MS)
                    Call CaptureViewToFile(oViewer, picPath)
                    Call SetShowCollection(sel, mLeaves, 1)
                End If

            Case "ASSY"                     ' a sub-assembly: all leaves below it
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

            Case "PART"                     ' one part
                Set oProd = mRef(key)
                Call SetShowSingle(sel, oProd, 0)
                Call Settle(PART_SETTLE_MS)
                Call CaptureViewToFile(oViewer, picPath)
                Call SetShowSingle(sel, oProd, 1)

            Case "BODY"                     ' one body of a multi-body part
                Set oProd = mRef(key)
                Set oPart = Nothing
                Set oPartDoc = Nothing
                If TryGetLoadedPart(oProd, oPart, oPartDoc) Then
                    Set oBody = GetBodyByName(oPart, bodyName)
                    If Not oBody Is Nothing And mBodies.Exists(mPartNum(key)) Then
                        bodyNames = mBodies(mPartNum(key))
                        Set partSel = oPartDoc.Selection
                        Call SetShowSingle(sel, oProd, 0)
                        Call SetShowBodyList(oPart, partSel, bodyNames, 1)
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
        End Select

        ' --- into Excel straight away ---
        hasFile = False
        On Error Resume Next
        hasFile = mFso.FileExists(picPath)
        Err.Clear
        On Error GoTo 0
        If hasFile Then
            If Not XlInsertPicture(xlSheet, rowNo, picPath) Then
                Call XlPutCell(xlSheet, rowNo, 1, "No Preview")
            End If
            On Error Resume Next
            mFso.DeleteFile picPath, True
            Err.Clear
            On Error GoTo 0
        Else
            Call XlPutCell(xlSheet, rowNo, 1, "No Preview")
        End If

        If Elapsed(t) >= SLOW_STEP_SECONDS Then
            Call NoteSlow(Elapsed(t), mSheetName & "  picture of row " & rowNo & "  " & what)
        End If
        Call SaveWorkbook(False)                  ' at most every SAVE_EVERY_SECONDS
        DoEvents
    Next j

    ' --- back to normal ---
    Call RestoreVisibility(sel, mLeaves, colWasShown, colWasHidden)

    On Error Resume Next
    mXlApp.ScreenUpdating = True
    CATIA.StatusBar = ""
    Err.Clear
    On Error GoTo 0
End Sub


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
    ReDim mBuf(1 To 15, 1 To 256)
    mBufRows = 0
    mFlushed = False
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
    If i < 1 Or c < 1 Or c > 15 Then Exit Sub
    n = UBound(mBuf, 2)
    If i > n Then
        Do While n < i
            n = n * 2
        Loop
        ReDim Preserve mBuf(1 To 15, 1 To n)
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
Private Sub FlushSheet(xlSheet As Object)
    Const BLOCK_ROWS As Long = 500
    Dim arr() As Variant
    Dim i0 As Long, n As Long, i As Long, c As Long
    Dim rr As Variant

    mFlushed = True
    If mBufRows = 0 Then Exit Sub

    Call SetStatus("BOM " & mSheetName & ": writing " & mBufRows & " rows to Excel ...")
    For i0 = 1 To mBufRows Step BLOCK_ROWS
        n = mBufRows - i0 + 1
        If n > BLOCK_ROWS Then n = BLOCK_ROWS
        ReDim arr(1 To n, 1 To 15)
        For i = 1 To n
            For c = 1 To 15
                arr(i, c) = mBuf(c, i0 + i - 1)
            Next c
        Next i
        Call XlWriteBlock(xlSheet, i0 + 1, 1, arr)
    Next i0

    ' assembly rows in bold - only looks, so a failure here is ignored
    If mBoldRows.Count > 0 Then
        Call XlReady
        On Error Resume Next
        For Each rr In mBoldRows
            xlSheet.Range(xlSheet.Cells(rr, 2), xlSheet.Cells(rr, 15)).Font.Bold = True
        Next rr
        Err.Clear
        On Error GoTo 0
    End If
End Sub


' ==============================================================================
' HELPER: THE COLUMN HEADERS OF A BOM SHEET
' (identical on every sheet, so every sheet imports into the ACE cost model)
' ==============================================================================
Private Sub WriteHeader(xlSheet As Object)
    Dim h As Variant, a(1 To 1, 1 To 15) As Variant
    Dim i As Long

    h = Array("Thumbnail", "Level", "Part Number", "Description", "Qty", _
              "Mass (kg)", "Volume (m3)", "Area (m2)", "Length (mm)", "Width (mm)", _
              "Height (mm)", "Material", "Density (kg/m3)", "First Level", "Type")
    For i = 0 To 14
        a(1, i + 1) = h(i)
    Next i

    ' the formats first: part numbers stay text (leading zeros, "1E5" ...)
    Call XlReady
    On Error Resume Next
    With xlSheet
        .Columns("A:A").ColumnWidth = 15
        .Columns("B:B").ColumnWidth = 8
        .Columns("C:D").ColumnWidth = 20
        .Columns("C:C").NumberFormat = "@"
        .Columns("F:H").NumberFormat = "0.000000000"
        .Columns("I:K").NumberFormat = "0.0"
        .Columns("M:M").NumberFormat = "0.000"
        .Columns("N:N").NumberFormat = "@"
        .Columns("N:N").ColumnWidth = 22
        .Columns("O:O").ColumnWidth = 10
        .Range("A1:O1").Font.Bold = True
        .Range("A1:O1").Interior.Color = RGB(220, 220, 220)
    End With
    Err.Clear
    On Error GoTo 0

    Call XlWriteBlock(xlSheet, 1, 1, a)
End Sub


' ==============================================================================
' HELPER: FINISH A BOM SHEET (column widths + filter) AND SHOW IT
' ==============================================================================
Private Sub FinishSheetLayout(xlSheet As Object)
    Call XlReady
    On Error Resume Next
    xlSheet.Columns("B:O").AutoFit
    xlSheet.Columns("A:A").ColumnWidth = 15
    xlSheet.Columns("N:N").ColumnWidth = 22
    If Not xlSheet.AutoFilterMode Then xlSheet.Range("B1:O1").AutoFilter
    xlSheet.Activate
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
Private Sub WriteIndexHeader(idxSheet As Object, ByVal splitLevel As Long)
    Dim a(1 To 2, 1 To 12) As Variant

    a(1, 1) = "BOM split at level " & splitLevel & _
              "  -  quantities on a component sheet are per ONE piece of that component; " & _
              "column E says how many pieces the product contains"
    a(2, 1) = "#"
    a(2, 2) = "Component"
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

    Call XlReady
    On Error Resume Next
    idxSheet.Columns("B:B").NumberFormat = "@"
    Err.Clear
    On Error GoTo 0
    Call XlWriteBlock(idxSheet, 1, 1, a)

    Call XlReady
    On Error Resume Next
    With idxSheet
        .Cells(1, 1).Font.Italic = True
        .Range("A2:L2").Font.Bold = True
        .Range("A2:L2").Interior.Color = RGB(220, 220, 220)
    End With
    Err.Clear
End Sub

Private Sub WriteIndexRow(idxSheet As Object, ByVal idxRow As Long, ByVal no As Long, _
                          ByVal pn As String, ByVal desc As String, ByVal lvl As Long, _
                          ByVal qty As Long, ByVal firstLevel As String, ByVal pathTxt As String, _
                          ByVal sheetName As String)
    Dim a(1 To 1, 1 To 10) As Variant

    If no > 0 Then a(1, 1) = no
    a(1, 2) = pn
    a(1, 3) = desc
    If lvl > 0 Then a(1, 4) = lvl
    a(1, 5) = qty
    a(1, 6) = firstLevel
    a(1, 7) = pathTxt
    a(1, 10) = "waiting ..."
    Call XlWriteBlock(idxSheet, idxRow, 1, a)

    If Len(sheetName) = 0 Then Exit Sub
    Call XlReady
    On Error Resume Next
    idxSheet.Hyperlinks.Add idxSheet.Cells(idxRow, 8), "", _
                            "'" & Replace(sheetName, "'", "''") & "'!A1", "", sheetName
    Err.Clear
End Sub

Private Sub WriteIndexResult(idxSheet As Object, ByVal idxRow As Long, ByVal nRows As Long, _
                             ByVal statusTxt As String, ByVal secs As Double, ByVal whereTxt As String)
    Dim a(1 To 1, 1 To 4) As Variant

    a(1, 1) = nRows
    a(1, 2) = statusTxt
    a(1, 3) = Round(secs, 1)
    a(1, 4) = whereTxt
    Call XlWriteBlock(idxSheet, idxRow, 9, a)

    If statusTxt = "OK" Then Exit Sub
    Call XlReady
    On Error Resume Next
    idxSheet.Range(idxSheet.Cells(idxRow, 1), idxSheet.Cells(idxRow, 12)).Interior.Color = RGB(255, 220, 200)
    Err.Clear
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
    Dim msg As String, ans As String
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
          "n = one sheet per sub-assembly on level n," & vbCrLf & _
          "     each one worked in its own window" & vbCrLf & vbCrLf & _
          "Level:  components  (sub-assemblies)" & vbCrLf
    For lvl = 1 To maxDepth
        If lvl > 15 Then
            msg = msg & "  ... deeper levels down to " & maxDepth & vbCrLf
            Exit For
        End If
        msg = msg & "  " & lvl & ":   " & cntAll(lvl) & "  (" & cntAssy(lvl) & ")" & vbCrLf
    Next lvl

    Do
        ans = Trim$(InputBox(msg, "BOM - split level", "0"))
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
            rawPn = ""
            rawPn = child.PartNumber
            pn = rawPn
            If Len(pn) = 0 Then pn = child.Name

            If lvl = 1 Then
                childFirst = FirstLevelLabel(child, pn)
            Else
                childFirst = firstName
            End If
            rowKey = MakeRowKey(parentKey, childFirst, pn)

            nk = 0
            nk = child.Products.Count
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
                If Len(pathTxt) > 0 Then childPath = pathTxt & " > " & pn Else childPath = pn
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
    t0 = Timer
    Do
        If Not XlReady() Then Exit Function
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
    If mXlLost Then Err.Raise vbObjectError + 515, , _
        "Excel was closed or does not answer any more - the run is stopped"
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
    ws.Rows(rowNo).RowHeight = 60
    Err.Clear
    On Error GoTo 0
    XlInsertPicture = True
End Function

' Excel back to normal at the end of the run (also after an error)
Private Sub XlRestore()
    If mXlApp Is Nothing Then Exit Sub
    On Error Resume Next
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

    mLogPath = ""
    mSavePath = ""
    mSavedOnce = False
    mLastSave = Timer
    Set mSlow = New Collection

    folder = SAVE_FOLDER
    If Len(folder) = 0 Then folder = "C:\Temp"
    Call EnsureFolder(folder)
    On Error Resume Next
    If Not mFso.FolderExists(folder) Then folder = CStr(Environ$("TEMP"))
    Err.Clear
    On Error GoTo 0

    baseName = rootPN
    For Each bad In Array("\", "/", ":", "*", "?", Chr(34), "<", ">", "|")
        baseName = Replace(baseName, CStr(bad), "_")
    Next bad
    If Len(baseName) > 60 Then baseName = Left$(baseName, 60)
    baseName = folder & "\BOM_" & baseName & "_" & Format(Now, "yyyymmdd_hhnnss")

    mLogPath = baseName & ".txt"
    If Len(SAVE_FOLDER) > 0 Then mSavePath = baseName & ".xlsx"
End Sub

' saves the workbook (first time: Save As). force = False saves at most every
' SAVE_EVERY_SECONDS. A failed save is logged and the run goes on.
Private Sub SaveWorkbook(ByVal force As Boolean)
    Dim e As Long, d As String

    If Len(mSavePath) = 0 Or mXlBook Is Nothing Or mXlLost Then Exit Sub
    If Not force And mSavedOnce And Elapsed(mLastSave) < SAVE_EVERY_SECONDS Then Exit Sub

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

    If e = 0 Then
        mSavedOnce = True
        Call LogLine("saved " & mSavePath)
    Else
        Call LogLine("saving failed: error " & e & " - " & d)
    End If
End Sub

Private Function SavedText() As String
    Dim s As String
    If mSavedOnce Then
        s = "Saved as:" & vbCrLf & mSavePath
    ElseIf Len(mSavePath) > 0 Then
        s = "The workbook could NOT be saved automatically - please save it yourself."
    Else
        s = "The workbook is not saved yet."
    End If
    If Len(mLogPath) > 0 Then s = s & vbCrLf & vbCrLf & "Log of the run:" & vbCrLf & mLogPath
    SavedText = s
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
' One Selection.Add per item but only ONE SetShow call (= one visual update)
' In CATIA: SetShow(0) = SHOW, SetShow(1) = HIDE
' ==============================================================================
Sub SetShowCollection(sel As Selection, colItems As Collection, ByVal showMode As Integer)
    On Error Resume Next
    Dim i As Long

    If colItems Is Nothing Then Exit Sub
    If colItems.Count = 0 Then Exit Sub

    sel.Clear
    For i = 1 To colItems.Count
        sel.Add colItems.Item(i)
        If i Mod 200 = 0 Then DoEvents     ' a long list must not freeze CATIA
    Next i
    sel.VisProperties.SetShow showMode
    sel.Clear
    Err.Clear
End Sub

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
' HELPER: NAME OF A FIRST LEVEL NODE (part number, or description on request)
' ==============================================================================
Function FirstLevelLabel(oProd As Product, ByVal partNum As String) As String
    Dim s As String
    Dim d As String

    s = partNum

    If FIRST_LEVEL_USE_DESCRIPTION Then
        d = ""
        On Error Resume Next
        d = oProd.DescriptionRef
        Err.Clear
        On Error GoTo 0
        If Len(Trim$(d)) > 0 Then s = d
    End If

    If Len(s) = 0 Then
        On Error Resume Next
        s = oProd.Name
        Err.Clear
        On Error GoTo 0
    End If

    FirstLevelLabel = s
End Function

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
