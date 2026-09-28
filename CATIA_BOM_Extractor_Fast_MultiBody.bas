Option Explicit

' ==============================================================================
' MACRO: MASTER BOM v4 (SPLIT BY LEVEL + NUMBERS FIRST, PICTURES LAST)
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
'   2. Every sheet is filled in TWO PASSES:
'        pass 1 - all numbers: the display is frozen and nothing is shown or
'                 hidden, CATIA only loads and measures. The sheet is complete
'                 and visible in Excel before the first picture is taken.
'        pass 2 - all pictures: one show / capture / hide per row into picture
'                 FILES, then all files are inserted into Excel together.
'   3. A component that fails is marked on the Index sheet and the run goes on
'      with the next one.
'
' WHY THIS VERSION IS FAST AND QUIET:
'   - NO temporary geometry is created in the CATParts (MEASURE_MODE). Creating
'     a bounding box, running the update engine on it and deleting it again is
'     what made CATIA show an error panel for EVERY component. The dimensions
'     come from the inertia, which only reads.
'   - The specification tree is switched off through the WINDOW LAYOUT, which
'     works in every language and immediately.
'   - Every wait pumps the message queue (Settle) instead of blocking the thread,
'     so CATIA stays responsive and actually redraws before a capture.
'   - HSOSynchronized is off for the run, so Selection.Add does not repaint the
'     tree every time.
'   - Every node is promoted to design mode in its own row instead of one
'     blocking call over the whole tree.
'   - Mass, material, dimensions and body names are read ONCE per part number
'     for the whole run, over all sheets.
'
' IF IT IS STILL TOO SLOW: split at a deeper level first. Otherwise the
' settings block below has the switches - each of them trades content for speed:
'   CAPTURE_ASSEMBLY_SHOTS = False  no pictures on the assembly rows
'   CAPTURE_PART_SHOTS = False      no pictures at all - the biggest single win
'   ASSEMBLY_VOLUME_AREA = False    assembly rows keep the mass, lose volume/area
'   ASSEMBLY_METRICS = False        assembly rows have no mass/volume/area
'   DETECT_MULTIBODY = False        no extra row per body of a multi-body part
'   MEASURE_MODE = "BOX"           (the opposite: exact box, but slow and noisy)
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
' NEW FEATURE (multi-body CATPart):
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
' Assembly thumbnails have to show and hide every leaf below the node. That is
' the most expensive single thing in the run, but without it the assembly rows
' have no picture at all - so it stays ON, with a cap: a node with more leaves
' than the limit is listed without a picture instead of stalling the run.
' Set CAPTURE_ASSEMBLY_SHOTS = False if you want the run as fast as possible.
Private Const CAPTURE_ASSEMBLY_SHOTS As Boolean = True
Private Const MAX_ASSY_LEAVES_FOR_SHOT As Long = 400
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
' not need one row per body: on a big tree that alone saves most of the loading.
Private Const DETECT_MULTIBODY As Boolean = True
' Product.ApplyWorkMode is RECURSIVE: one call on the root promotes every
' CATPart of the assembly in a single blocking COM call - that is the long
' "CATIA not responding" phase before the first row appears. Parts are promoted
' one at a time in the row loop instead, which spreads the same work over a
' responsive UI. Set this to True only if your parts must all be loaded first.
Private Const FORCE_DESIGN_MODE As Boolean = False
' Sub-assembly rows are promoted too, because their mass/volume/area and their
' inertia box are 0 otherwise. ApplyWorkMode is recursive, so the first level 1
' assembly row loads its whole branch in one blocking call - set this to False
' if you would rather have empty assembly rows than that wait.
Private Const PROMOTE_ASSEMBLY_ROWS As Boolean = True

' A component in NoShow hides its WHOLE subtree in the 3D view, whatever the
' children say. So a leaf below a sub-assembly that the user had hidden would be
' captured as an empty picture. The run therefore shows every node once and then
' hides only the leaves. Side effect: components that were hidden before the run
' are visible after it. Set PRESERVE_VISIBILITY = True to have the macro record
' the previous state of every node and put it back - correct, but it costs one
' selection round trip per node before the first row.
Private Const PRESERVE_VISIBILITY As Boolean = False

' Mass / volume / area of a sub-assembly row. Reading them means a recursive
' evaluation over the whole subtree, so they are the expensive part of the
' assembly structure version - but without them the assembly rows are empty,
' so they stay ON.
Private Const ASSEMBLY_METRICS As Boolean = True
' ... of the three, Volume and WetArea are the costly ones (a full B-Rep
' integration over the subtree). Set this to False to keep only the mass on
' assembly rows if the run is too slow.
Private Const ASSEMBLY_VOLUME_AREA As Boolean = True

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

' --- caches (whole run, keyed by part number) --------------------------------
' A part that appears on several sheets is read only once.
Private mDimCache As Object      ' part number -> Array(L, W, H)
Private mMatCache As Object      ' part number -> Array(material, density)
Private mBodyCache As Object     ' part number -> body names, or Empty for one body
Private mPropCache As Object     ' part number -> Array(mass, volume, area)
Private mBBoxDisabled As Boolean ' the temp geometry path failed once - stop it

' --- the sheet being written (reset for every sheet) --------------------------
Private mQty As Object, mRef As Object, mDesc As Object, mProps As Object
Private mLevel As Object, mBodies As Object, mPartNum As Object, mFirst As Object
Private mIsAssy As Object, mNodeLeaves As Object
Private mLeaves As Collection    ' every leaf instance of the sheet
Private mAllNodes As Collection  ' every instance of the sheet: assemblies AND leaves
Private mJobRow As Collection    ' pictures to take in pass 2: row ...
Private mJobKind As Collection   ' ... "PART" / "ASSY" / "BODY" / "UNIT"
Private mJobKey As Collection    ' ... row key
Private mJobBody As Collection   ' ... body name for "BODY"

' --- the whole run -------------------------------------------------------------
Private mXlApp As Object, mXlBook As Object, mFso As Object
Private mTempDir As String
Private mSheetNo As Long


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
'   1. asks at which level the BOM is split (0 = one sheet, as before)
'   2. split level 0:  one sheet for the whole product
'      split level n:  one sheet per sub-assembly on level n, each worked in its
'                      own CATIA window; parts above level n on an extra sheet;
'                      an "Index" sheet with a link to every component sheet
'   3. every sheet is filled in two passes: first ALL numbers (display frozen,
'      no show/hide), then all pictures in one pass, inserted into Excel together
' ==============================================================================
Private Sub RunBOM()

    Dim catDoc As Document
    Dim rootProd As Product
    Dim rootSel As Selection
    Dim rootWin As Object
    Dim rootViewer As Viewer
    Dim rootLayout As Long, rootLayoutSet As Boolean
    Dim compassToggled As Boolean

    Dim splitLevel As Long
    Dim xlSheet As Object, idxSheet As Object
    Dim startTime As Double, unitStart As Double
    Dim nRows As Long, nAssy As Long, nParts As Long
    Dim totRows As Long, totAssy As Long, totParts As Long
    Dim nUnits As Long, nUnitsOK As Long
    Dim statusTxt As String, unitOK As Boolean

    ' components on the split level, keyed by part number
    Dim dUnitQty As Object, dUnitRef As Object, dUnitFirst As Object
    Dim dUnitDesc As Object, dUnitPath As Object
    Dim colUnitInst As Collection      ' every instance on the split level
    Dim colAbove As Collection         ' every assembly instance above it
    Dim colLooseInst As Collection     ' parts that are not inside any component
    Dim colLooseLevel As Collection, colLooseFirst As Collection

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

    ' --- 2. WHERE TO SPLIT ------------------------------------------------------
    splitLevel = AskSplitLevel(rootProd)
    If splitLevel < 0 Then Exit Sub                     ' cancelled

    ' --- 3. SESSION --------------------------------------------------------------
    On Error Resume Next
    CATIA.DisplayFileAlerts = False
    CATIA.HSOSynchronized = False    ' otherwise every Selection.Add repaints the tree
    If FORCE_DESIGN_MODE Then rootProd.ApplyWorkMode 2
    Err.Clear
    On Error GoTo Fail

    Set mFso = CreateObject("Scripting.FileSystemObject")
    If Not mFso.FolderExists("C:\Temp") Then mFso.CreateFolder "C:\Temp"
    mTempDir = "C:\Temp\catia_bom_pictures"
    If Not mFso.FolderExists(mTempDir) Then mFso.CreateFolder mTempDir

    ' caches: the same part number is measured once for the whole run, over all
    ' sheets - a part that is used in ten components is read once
    Set mDimCache = CreateObject("Scripting.Dictionary")
    Set mMatCache = CreateObject("Scripting.Dictionary")
    Set mBodyCache = CreateObject("Scripting.Dictionary")
    Set mPropCache = CreateObject("Scripting.Dictionary")
    mBBoxDisabled = False
    mSheetNo = 0

    ' --- 4. EXCEL ----------------------------------------------------------------
    On Error Resume Next
    Set mXlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Or mXlApp Is Nothing Then
        Err.Clear
        Set mXlApp = CreateObject("Excel.Application")
    End If
    Err.Clear
    On Error GoTo Fail

    mXlApp.Visible = True
    Set mXlBook = mXlApp.Workbooks.Add
    Call KeepOneSheet(mXlBook)

    ' --- 5. MAIN WINDOW: geometry only, fixed camera ----------------------------
    Call PrepareWindow(rootWin, rootViewer, rootLayout, rootLayoutSet)
    compassToggled = ToggleCompass()

    If splitLevel = 0 Then

        ' ===== ONE SHEET FOR THE WHOLE PRODUCT ==================================
        Set xlSheet = mXlBook.Sheets(1)
        xlSheet.Name = "BOM"
        unitOK = ProcessTree(rootProd, rootSel, rootViewer, xlSheet, 0, "", "", 0, _
                             nRows, nAssy, nParts, statusTxt)
        totRows = nRows: totAssy = nAssy: totParts = nParts
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
        Set colFbKey = New Collection
        Set colFbSheet = New Collection
        Set colFbIdxRow = New Collection

        On Error Resume Next
        CATIA.StatusBar = "BOM: collecting the components on level " & splitLevel & " ..."
        Err.Clear
        On Error GoTo Fail
        Call CollectUnits(rootProd, 1, splitLevel, "", "", dUnitQty, dUnitRef, dUnitFirst, _
                          dUnitDesc, dUnitPath, colUnitInst, colAbove, _
                          colLooseInst, colLooseLevel, colLooseFirst)

        Set idxSheet = mXlBook.Sheets(1)
        idxSheet.Name = "Index"
        Call WriteIndexHeader(idxSheet, splitLevel)
        idxRow = 3

        ' --- the components that are CATProducts of their own: own window ------
        For Each unitKey In dUnitQty.Keys
            nUnits = nUnits + 1
            unitStart = Timer
            Set unitInst = dUnitRef(unitKey)
            Set xlSheet = AddSheet(mXlBook, Format(nUnits, "00") & " " & CStr(unitKey))

            Call WriteIndexRow(idxSheet, idxRow, nUnits, CStr(unitKey), CStr(dUnitDesc(unitKey)), _
                               splitLevel, CLng(dUnitQty(unitKey)), CStr(dUnitFirst(unitKey)), _
                               CStr(dUnitPath(unitKey)), xlSheet.Name)

            ' promote the component, then look for a document of its own
            On Error Resume Next
            CATIA.StatusBar = "BOM: component " & nUnits & " of " & dUnitQty.Count & _
                              " - loading " & CStr(unitKey) & " ..."
            unitInst.ApplyWorkMode 2
            Err.Clear
            On Error GoTo Fail

            ownWindow = False
            Set unitDoc = Nothing
            Set unitWin = Nothing
            If OPEN_UNITS_IN_NEW_WINDOW Then Set unitDoc = OwnDocumentOf(unitInst, CStr(unitKey), catDoc)

            If Not unitDoc Is Nothing Then
                On Error Resume Next
                Set unitWin = unitDoc.NewWindow
                If Not unitWin Is Nothing Then unitWin.Activate
                Err.Clear
                On Error GoTo Fail
                ownWindow = Not (unitWin Is Nothing)
            End If

            If ownWindow Then
                Set unitViewer = Nothing
                Call PrepareWindow(unitWin, unitViewer, unitLayout, unitLayoutSet)
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
                On Error GoTo Fail
                Set unitWin = Nothing

                Call WriteIndexResult(idxSheet, idxRow, nRows, statusTxt, Timer - unitStart, "own window")
                totRows = totRows + nRows: totAssy = totAssy + nAssy: totParts = totParts + nParts
                If unitOK Then nUnitsOK = nUnitsOK + 1
            Else
                ' done later in the main window, once the rest of the scene is hidden
                colFbKey.Add unitKey
                colFbSheet.Add xlSheet
                colFbIdxRow.Add idxRow
            End If

            idxRow = idxRow + 1
            DoEvents
        Next unitKey

        ' --- the rest is done in the main window ---------------------------------
        If colFbKey.Count > 0 Or colLooseInst.Count > 0 Then
            ' scene: every assembly above the split level visible, every component
            ' and every loose part hidden - then one thing at a time is shown
            Call SetShowCollection(rootSel, colAbove, 0)
            Call SetShowCollection(rootSel, colUnitInst, 1)
            Call SetShowCollection(rootSel, colLooseInst, 1)
            rootSceneSet = True

            For i = 1 To colFbKey.Count
                unitKey = colFbKey(i)
                unitStart = Timer
                Set unitInst = dUnitRef(unitKey)

                Set xlSheet = colFbSheet(i)
                Call SetShowSingle(rootSel, unitInst, 0)         ' only this branch
                unitOK = ProcessTree(unitInst, rootSel, rootViewer, xlSheet, splitLevel, _
                                     CStr(dUnitFirst(unitKey)), CStr(unitKey), CLng(dUnitQty(unitKey)), _
                                     nRows, nAssy, nParts, statusTxt)
                Call SetShowSingle(rootSel, unitInst, 1)

                Call WriteIndexResult(idxSheet, CLng(colFbIdxRow(i)), nRows, statusTxt, _
                                      Timer - unitStart, "main window")
                totRows = totRows + nRows: totAssy = totAssy + nAssy: totParts = totParts + nParts
                If unitOK Then nUnitsOK = nUnitsOK + 1
                DoEvents
            Next i

            If colLooseInst.Count > 0 Then
                unitStart = Timer
                Set xlSheet = AddSheet(mXlBook, "Parts above level " & splitLevel)
                unitOK = ProcessLoose(colLooseInst, colLooseLevel, colLooseFirst, rootSel, rootViewer, _
                                      xlSheet, nRows, nParts, statusTxt)
                Call WriteIndexRow(idxSheet, idxRow, 0, "(parts above level " & splitLevel & ")", _
                                   "parts that are not inside any component of level " & splitLevel, _
                                   0, colLooseInst.Count, "", "", xlSheet.Name)
                Call WriteIndexResult(idxSheet, idxRow, nRows, statusTxt, Timer - unitStart, "main window")
                totRows = totRows + nRows: totParts = totParts + nParts
            End If

            ' everything visible again
            Call SetShowCollection(rootSel, colUnitInst, 0)
            Call SetShowCollection(rootSel, colLooseInst, 0)
            rootSceneSet = False
        End If

        Call FinishIndexLayout(idxSheet)
        idxSheet.Activate
    End If

    ' --- 6. RESTORE THE SESSION --------------------------------------------------
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
    mXlApp.ScreenUpdating = True
    Err.Clear
    On Error GoTo Fail

    Dim modeName As String
    If mIncludeSubassemblies Then modeName = "Assembly structure" Else modeName = "Part list"

    If splitLevel = 0 Then
        MsgBox "BOM exported  (" & modeName & ")" & vbCrLf & vbCrLf & _
               "Rows written:     " & totRows & vbCrLf & _
               "  sub-assemblies: " & totAssy & vbCrLf & _
               "  parts:          " & totParts & vbCrLf & _
               "Time: " & Format(Timer - startTime, "0.0") & " s", vbInformation
    Else
        MsgBox "BOM exported  (" & modeName & ", split at level " & splitLevel & ")" & vbCrLf & vbCrLf & _
               "Component sheets: " & nUnits & "  (" & nUnitsOK & " complete)" & vbCrLf & _
               "Rows written:     " & totRows & vbCrLf & _
               "Time: " & Format(Timer - startTime, "0.0") & " s" & vbCrLf & vbCrLf & _
               "The Index sheet links to every component sheet.", vbInformation
    End If
    Exit Sub

' ------------------------------------------------------------------------------
' Something went wrong outside a component: put CATIA and Excel back into a
' usable state before reporting. (Errors INSIDE a component are handled there -
' that component is marked on the Index sheet and the run goes on.)
' ------------------------------------------------------------------------------
Fail:
    Dim failMsg As String
    failMsg = "The BOM run stopped with error " & Err.Number & " - " & Err.Description

    On Error Resume Next
    If Not unitWin Is Nothing Then unitWin.Close
    If Not rootWin Is Nothing Then rootWin.Activate
    If rootSceneSet Then
        Call SetShowCollection(rootSel, colUnitInst, 0)
        Call SetShowCollection(rootSel, colLooseInst, 0)
    End If
    If rootLayoutSet Then rootWin.Layout = rootLayout
    If compassToggled Then CATIA.StartCommand COMPASS_COMMAND
    If Not rootViewer Is Nothing Then rootViewer.Reframe
    CATIA.HSOSynchronized = True
    CATIA.RefreshDisplay = True
    CATIA.DisplayFileAlerts = True
    CATIA.StatusBar = ""
    If Not rootSel Is Nothing Then rootSel.Clear
    If Not mXlApp Is Nothing Then mXlApp.ScreenUpdating = True
    Err.Clear

    MsgBox failMsg & vbCrLf & vbCrLf & _
           "Visibility, spec tree and compass were restored." & vbCrLf & _
           "The rows written so far are in the Excel workbook.", vbCritical

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
    Dim r As Long

    ProcessTree = False
    nRows = 0: nAssy = 0: nParts = 0
    r = 2
    statusTxt = ""
    On Error GoTo SheetFail

    Call ResetSheetData
    Call WriteHeader(xlSheet)

    ' --- the tree of this sheet (numbers are NOT read here) ---
    On Error Resume Next
    CATIA.StatusBar = "BOM: reading the tree of " & xlSheet.Name & " ..."
    Err.Clear
    On Error GoTo SheetFail

    Call TraverseTree(treeRoot, mQty, mRef, mDesc, mProps, mLevel, mBodies, mPartNum, mFirst, _
                      mIsAssy, mNodeLeaves, mLeaves, levelOffset + 1, firstLevel, unitPartNum)

    ' --- the component itself is the first row of its sheet ---
    If levelOffset > 0 Then
        Call WriteUnitRow(treeRoot, xlSheet, r, levelOffset, unitPartNum, unitQty, firstLevel)
        nAssy = nAssy + 1
        r = r + 1
    End If

    ' --- PASS 1: all numbers, display frozen, nothing shown or hidden ---
    mXlApp.ScreenUpdating = False           ' Excel repaints are the slow part of cell writes
    Call WriteDataRows(xlSheet, r, nAssy, nParts)
    nRows = r - 2

    ' the numbers are complete - make them visible before the pictures start
    Call FinishSheetLayout(xlSheet)

    ' --- PASS 2: the pictures ---
    Call TakePictures(sel, oViewer, xlSheet)

    ProcessTree = True
    statusTxt = "OK"
    Exit Function

SheetFail:
    statusTxt = "stopped: " & Err.Number & " - " & Err.Description
    nRows = r - 2
    On Error Resume Next
    Call SetShowCollection(sel, mLeaves, 0)
    CATIA.RefreshDisplay = True
    mXlApp.ScreenUpdating = True
    xlSheet.Cells(r + 1, 3).Value = "!! " & statusTxt
    Err.Clear
End Function


' ==============================================================================
' THE SHEET OF THE PARTS ABOVE THE SPLIT LEVEL
' (parts that do not sit inside any component of the split level)
' ==============================================================================
Private Function ProcessLoose(colInst As Collection, colLevel As Collection, _
                              colFirst As Collection, sel As Selection, oViewer As Viewer, _
                              xlSheet As Object, ByRef nRows As Long, ByRef nParts As Long, _
                              ByRef statusTxt As String) As Boolean
    Dim i As Long, r As Long, dummyAssy As Long
    Dim oProd As Product, pn As String, sKey As String

    ProcessLoose = False
    nRows = 0: nParts = 0
    r = 2
    statusTxt = ""
    On Error GoTo LooseFail

    Call ResetSheetData
    Call WriteHeader(xlSheet)

    For i = 1 To colInst.Count
        Set oProd = colInst(i)
        pn = ""
        On Error Resume Next
        pn = oProd.PartNumber
        Err.Clear
        On Error GoTo LooseFail

        mLeaves.Add oProd
        mAllNodes.Add oProd
        If Len(pn) > 0 Then
            sKey = CStr(colFirst(i)) & "|" & pn
            If mQty.Exists(sKey) Then
                mQty(sKey) = mQty(sKey) + 1
            Else
                Call AddBomRow(mQty, mRef, mDesc, mProps, mLevel, mPartNum, mFirst, mIsAssy, _
                               sKey, oProd, pn, CStr(colFirst(i)), CInt(colLevel(i)), False)
            End If
        End If
    Next i

    mXlApp.ScreenUpdating = False
    Call WriteDataRows(xlSheet, r, dummyAssy, nParts)
    nRows = r - 2
    Call FinishSheetLayout(xlSheet)
    Call TakePictures(sel, oViewer, xlSheet)

    ProcessLoose = True
    statusTxt = "OK"
    Exit Function

LooseFail:
    statusTxt = "stopped: " & Err.Number & " - " & Err.Description
    nRows = r - 2
    On Error Resume Next
    CATIA.RefreshDisplay = True
    mXlApp.ScreenUpdating = True
    Err.Clear
End Function


' ==============================================================================
' PASS 1 - ALL NUMBERS OF ONE SHEET
' Nothing is shown, hidden or captured here and the display is frozen, so CATIA
' only loads and measures. Every picture that will be needed is noted as a job.
' Errors are not handled here on purpose: they go up to the sheet's handler.
' ==============================================================================
Private Sub WriteDataRows(xlSheet As Object, ByRef r As Long, ByRef nAssy As Long, _
                          ByRef nParts As Long)
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
    Dim nDone As Long

    On Error Resume Next
    CATIA.RefreshDisplay = False
    Err.Clear
    On Error GoTo 0

    For Each uniqueKey In mQty.Keys

        nDone = nDone + 1
        strPartNum = mPartNum(uniqueKey)
        Set oPartProd = mRef(uniqueKey)
        isAssy = mIsAssy(uniqueKey)

        If isAssy Then nAssy = nAssy + 1 Else nParts = nParts + 1

        xlSheet.Cells(r, 2).Value = mLevel(uniqueKey)
        xlSheet.Cells(r, 3).Value = strPartNum
        xlSheet.Cells(r, 4).Value = mDesc(uniqueKey)
        xlSheet.Cells(r, 5).Value = mQty(uniqueKey)
        xlSheet.Cells(r, 14).Value = mFirst(uniqueKey)
        If isAssy Then
            xlSheet.Cells(r, 15).Value = "Assembly"
            xlSheet.Range(xlSheet.Cells(r, 2), xlSheet.Cells(r, 15)).Font.Bold = True
        Else
            xlSheet.Cells(r, 15).Value = "Part"
        End If

        On Error Resume Next
        CATIA.StatusBar = "BOM " & xlSheet.Name & ": numbers, row " & nDone & " of " & _
                          mQty.Count & " - " & strPartNum
        Err.Clear
        On Error GoTo 0

        ' promote THIS node to design mode first - mass, volume, area and the
        ' inertia box are all 0 on a component in visualization mode
        If (Not isAssy) Or PROMOTE_ASSEMBLY_ROWS Then
            On Error Resume Next
            oPartProd.ApplyWorkMode 2
            Err.Clear
            On Error GoTo 0
        End If

        hasPart = False
        Set oPart = Nothing
        Set oPartDoc = Nothing
        If Not isAssy Then hasPart = TryGetLoadedPart(oPartProd, oPart, oPartDoc)

        ' multi-body parts: the body names are cached per part number, so a part
        ' that appears on several sheets gets its body rows on every one of them
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
            If (Not isAssy) Or ASSEMBLY_METRICS Then
                Call ReadProductProps(oPartProd, dMass, dVol, dArea, isAssy)
            End If
            propsArray = Array(dMass, dVol, dArea)
            If dMass > 0 Or dVol > 0 Or dArea > 0 Then mPropCache.Add strPartNum, propsArray
        End If

        If propsArray(0) > 0 Then xlSheet.Cells(r, 6).Value = propsArray(0) Else xlSheet.Cells(r, 6).Value = "N/A"
        If propsArray(1) > 0 Then xlSheet.Cells(r, 7).Value = propsArray(1) Else xlSheet.Cells(r, 7).Value = "N/A"
        If propsArray(2) > 0 Then xlSheet.Cells(r, 8).Value = propsArray(2) Else xlSheet.Cells(r, 8).Value = "N/A"

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

        xlSheet.Cells(r, 12).Value = strMatName
        If dDensity > 0 Then xlSheet.Cells(r, 13).Value = dDensity Else xlSheet.Cells(r, 13).Value = "N/A"

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

        If dims(0) > 0 Then xlSheet.Cells(r, 9).Value = dims(0) Else xlSheet.Cells(r, 9).Value = "N/A"
        If dims(1) > 0 Then xlSheet.Cells(r, 10).Value = dims(1) Else xlSheet.Cells(r, 10).Value = "N/A"
        If dims(2) > 0 Then xlSheet.Cells(r, 11).Value = dims(2) Else xlSheet.Cells(r, 11).Value = "N/A"

        ' --- the picture of this row is taken in pass 2 ---
        If isAssy Then
            If CAPTURE_ASSEMBLY_SHOTS Then
                Call AddJob(r, "ASSY", CStr(uniqueKey), "")
            Else
                xlSheet.Cells(r, 1).Value = "-"
            End If
        ElseIf CAPTURE_PART_SHOTS Then
            Call AddJob(r, "PART", CStr(uniqueKey), "")
        Else
            xlSheet.Cells(r, 1).Value = "-"
        End If

        ' --- MULTI-BODY SUB-ROWS ---
        If hasPart And mBodies.Exists(strPartNum) Then
            bodyNames = mBodies(strPartNum)
            For bi = LBound(bodyNames) To UBound(bodyNames)
                Set oBody = GetBodyByName(oPart, CStr(bodyNames(bi)))
                If Not oBody Is Nothing Then
                    r = r + 1

                    xlSheet.Cells(r, 2).Value = mLevel(uniqueKey) + 1
                    xlSheet.Cells(r, 3).Value = "    " & CStr(bodyNames(bi))
                    xlSheet.Cells(r, 4).Value = "Body of " & strPartNum
                    xlSheet.Cells(r, 5).Value = mQty(uniqueKey)
                    xlSheet.Cells(r, 14).Value = mFirst(uniqueKey)
                    xlSheet.Cells(r, 15).Value = "Body"

                    bMass = 0: bVol = 0: bArea = 0
                    bDens = 0: bDensInertia = 0
                    bMat = "N/A"

                    Call GetBodyMaterial(oPart, oBody, bMat, bDens)
                    Call GetBodyMassAndDensity(oPartDoc, oBody, bMass, bDensInertia)
                    If bDens <= 0 Then bDens = bDensInertia
                    Call GetBodyVolumeArea(oPartDoc, oPart, oBody, bVol, bArea)
                    If bMass <= 0 And bDens > 0 And bVol > 0 Then bMass = bVol * bDens

                    If bMass > 0 Then xlSheet.Cells(r, 6).Value = bMass Else xlSheet.Cells(r, 6).Value = "N/A"
                    If bVol > 0 Then xlSheet.Cells(r, 7).Value = bVol Else xlSheet.Cells(r, 7).Value = "N/A"
                    If bArea > 0 Then xlSheet.Cells(r, 8).Value = bArea Else xlSheet.Cells(r, 8).Value = "N/A"

                    dims(0) = 0: dims(1) = 0: dims(2) = 0
                    If UCase$(MEASURE_MODE) = "BOX" Then
                        If Not MeasureTarget(oPart, oBody, dims) Then
                            Call MeasureAABBExtremum(oPart, oPartDoc, oBody, dims)
                        End If
                    End If
                    If dims(0) < 0.1 Then Call GetBodyInertiaDims(oPartDoc, oBody, dims)

                    If dims(0) > 0 Then xlSheet.Cells(r, 9).Value = dims(0) Else xlSheet.Cells(r, 9).Value = "N/A"
                    If dims(1) > 0 Then xlSheet.Cells(r, 10).Value = dims(1) Else xlSheet.Cells(r, 10).Value = "N/A"
                    If dims(2) > 0 Then xlSheet.Cells(r, 11).Value = dims(2) Else xlSheet.Cells(r, 11).Value = "N/A"

                    xlSheet.Cells(r, 12).Value = bMat
                    If bDens > 0 Then xlSheet.Cells(r, 13).Value = bDens Else xlSheet.Cells(r, 13).Value = "N/A"

                    If CAPTURE_PART_SHOTS Then
                        Call AddJob(r, "BODY", CStr(uniqueKey), CStr(bodyNames(bi)))
                    Else
                        xlSheet.Cells(r, 1).Value = "-"
                    End If
                End If
                Set oBody = Nothing
            Next bi
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
' ==============================================================================
Private Sub WriteUnitRow(treeRoot As Product, xlSheet As Object, ByVal r As Long, _
                         ByVal lvl As Long, ByVal pn As String, ByVal qty As Long, _
                         ByVal firstLevel As String)
    Dim dMass As Double, dVol As Double, dArea As Double
    Dim dims(2) As Double
    Dim sDesc As String

    sDesc = ""
    On Error Resume Next
    sDesc = treeRoot.DescriptionRef
    Err.Clear
    On Error GoTo 0

    xlSheet.Cells(r, 2).Value = lvl
    xlSheet.Cells(r, 3).Value = pn
    xlSheet.Cells(r, 4).Value = sDesc
    xlSheet.Cells(r, 5).Value = qty
    xlSheet.Cells(r, 14).Value = firstLevel
    xlSheet.Cells(r, 15).Value = "Assembly"
    xlSheet.Range(xlSheet.Cells(r, 2), xlSheet.Cells(r, 15)).Font.Bold = True

    If ASSEMBLY_METRICS Then
        On Error Resume Next
        treeRoot.ApplyWorkMode 2
        Err.Clear
        On Error GoTo 0
        Call ReadProductProps(treeRoot, dMass, dVol, dArea, True)
    End If
    If dMass > 0 Then xlSheet.Cells(r, 6).Value = dMass Else xlSheet.Cells(r, 6).Value = "N/A"
    If dVol > 0 Then xlSheet.Cells(r, 7).Value = dVol Else xlSheet.Cells(r, 7).Value = "N/A"
    If dArea > 0 Then xlSheet.Cells(r, 8).Value = dArea Else xlSheet.Cells(r, 8).Value = "N/A"

    dims(0) = 0: dims(1) = 0: dims(2) = 0
    On Error Resume Next
    Call GetInertiaDims(treeRoot.ReferenceProduct, dims)
    Err.Clear
    On Error GoTo 0
    If dims(0) > 0 Then xlSheet.Cells(r, 9).Value = dims(0) Else xlSheet.Cells(r, 9).Value = "N/A"
    If dims(1) > 0 Then xlSheet.Cells(r, 10).Value = dims(1) Else xlSheet.Cells(r, 10).Value = "N/A"
    If dims(2) > 0 Then xlSheet.Cells(r, 11).Value = dims(2) Else xlSheet.Cells(r, 11).Value = "N/A"

    xlSheet.Cells(r, 12).Value = "N/A"
    xlSheet.Cells(r, 13).Value = "N/A"

    ' the whole component, pictured in pass 2
    Call AddJob(r, "UNIT", "", "")
End Sub


' ==============================================================================
' PASS 2 - ALL PICTURES OF ONE SHEET
' 1. scene: every node of the sheet visible, every leaf hidden
' 2. per job: show -> reframe -> capture to a FILE -> hide  (CATIA only)
' 3. all files are inserted into Excel together (Excel only)
' ==============================================================================
Private Sub TakePictures(sel As Selection, oViewer As Viewer, xlSheet As Object)
    Dim j As Long, rowNo As Long
    Dim kind As String, key As String, bodyName As String
    Dim picPath As String
    Dim oProd As Product, colNode As Collection
    Dim oPart As Part, oPartDoc As PartDocument, partSel As Selection
    Dim oBody As Body, bodyNames As Variant
    Dim shp As Object
    Dim colWasShown As Collection, colWasHidden As Collection

    If mJobRow.Count = 0 Then Exit Sub

    If oViewer Is Nothing Then
        For j = 1 To mJobRow.Count
            xlSheet.Cells(mJobRow(j), 1).Value = "No Viewer"
        Next j
        Exit Sub
    End If

    ' --- 1. the scene ---
    If PRESERVE_VISIBILITY Then
        Set colWasShown = New Collection
        Set colWasHidden = New Collection
        Call SplitByShowState(sel, mAllNodes, colWasShown, colWasHidden)
    End If
    Call SetShowCollection(sel, mAllNodes, 0)
    Call SetShowCollection(sel, mLeaves, 1)
    Call Settle(100)

    ' --- 2. capture every picture into a file ---
    For j = 1 To mJobRow.Count
        rowNo = mJobRow(j)
        kind = mJobKind(j)
        key = mJobKey(j)
        bodyName = mJobBody(j)
        picPath = mTempDir & "\s" & mSheetNo & "_r" & rowNo & ".jpg"

        On Error Resume Next
        CATIA.StatusBar = "BOM " & xlSheet.Name & ": picture " & j & " of " & mJobRow.Count
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

        DoEvents
    Next j

    ' --- back to normal ---
    Call RestoreVisibility(sel, mLeaves, colWasShown, colWasHidden)

    ' --- 3. insert all pictures into Excel together ---
    On Error Resume Next
    CATIA.StatusBar = "BOM " & xlSheet.Name & ": inserting the pictures ..."
    mXlApp.ScreenUpdating = False
    For j = 1 To mJobRow.Count
        rowNo = mJobRow(j)
        picPath = mTempDir & "\s" & mSheetNo & "_r" & rowNo & ".jpg"
        If mFso.FileExists(picPath) Then
            Set shp = Nothing
            Set shp = xlSheet.Shapes.AddPicture(picPath, False, True, _
                        xlSheet.Cells(rowNo, 1).Left + 2, xlSheet.Cells(rowNo, 1).Top + 2, -1, -1)
            If Not shp Is Nothing Then
                shp.Height = 50
                If shp.Width > 90 Then shp.Width = 90
            End If
            xlSheet.Rows(rowNo).RowHeight = 60
            mFso.DeleteFile picPath
        Else
            xlSheet.Cells(rowNo, 1).Value = "No Preview"
        End If
    Next j
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
    oViewer.CaptureToFile 4, picPath
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
    Set mJobRow = New Collection
    Set mJobKind = New Collection
    Set mJobKey = New Collection
    Set mJobBody = New Collection
    mSheetNo = mSheetNo + 1
End Sub


' ==============================================================================
' HELPER: THE COLUMN HEADERS OF A BOM SHEET
' (identical on every sheet, so every sheet imports into the ACE cost model)
' ==============================================================================
Private Sub WriteHeader(xlSheet As Object)
    With xlSheet
        .Cells(1, 1).Value = "Thumbnail"
        .Cells(1, 2).Value = "Level"
        .Cells(1, 3).Value = "Part Number"
        .Cells(1, 4).Value = "Description"
        .Cells(1, 5).Value = "Qty"
        .Cells(1, 6).Value = "Mass (kg)"
        .Cells(1, 7).Value = "Volume (m3)"
        .Cells(1, 8).Value = "Area (m2)"
        .Cells(1, 9).Value = "Length (mm)"
        .Cells(1, 10).Value = "Width (mm)"
        .Cells(1, 11).Value = "Height (mm)"
        .Cells(1, 12).Value = "Material"
        .Cells(1, 13).Value = "Density (kg/m3)"
        .Cells(1, 14).Value = "First Level"
        .Cells(1, 15).Value = "Type"

        .Range("A1:O1").Font.Bold = True
        .Range("A1:O1").Interior.Color = RGB(220, 220, 220)
        .Columns("A:A").ColumnWidth = 15
        .Columns("B:B").ColumnWidth = 8
        .Columns("C:D").ColumnWidth = 20
        .Columns("F:H").NumberFormat = "0.000000000"
        .Columns("I:K").NumberFormat = "0.0"
        .Columns("M:M").NumberFormat = "0.000"
        .Columns("N:N").ColumnWidth = 22
        .Columns("O:O").ColumnWidth = 10
    End With
End Sub


' ==============================================================================
' HELPER: FINISH A BOM SHEET (column widths + filter) AND SHOW IT
' ==============================================================================
Private Sub FinishSheetLayout(xlSheet As Object)
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


' ==============================================================================
' HELPER: THE INDEX SHEET
' ==============================================================================
Private Sub WriteIndexHeader(idxSheet As Object, ByVal splitLevel As Long)
    With idxSheet
        .Cells(1, 1).Value = "BOM split at level " & splitLevel & _
                             "  -  quantities on a component sheet are per ONE piece of that component; " & _
                             "column E says how many pieces the product contains"
        .Cells(1, 1).Font.Italic = True
        .Cells(2, 1).Value = "#"
        .Cells(2, 2).Value = "Component"
        .Cells(2, 3).Value = "Description"
        .Cells(2, 4).Value = "Level"
        .Cells(2, 5).Value = "Qty in product"
        .Cells(2, 6).Value = "First Level"
        .Cells(2, 7).Value = "Path"
        .Cells(2, 8).Value = "Sheet"
        .Cells(2, 9).Value = "Rows"
        .Cells(2, 10).Value = "Status"
        .Cells(2, 11).Value = "Seconds"
        .Cells(2, 12).Value = "Worked in"
        .Range("A2:L2").Font.Bold = True
        .Range("A2:L2").Interior.Color = RGB(220, 220, 220)
    End With
End Sub

Private Sub WriteIndexRow(idxSheet As Object, ByVal idxRow As Long, ByVal no As Long, _
                          ByVal pn As String, ByVal desc As String, ByVal lvl As Long, _
                          ByVal qty As Long, ByVal firstLevel As String, ByVal pathTxt As String, _
                          ByVal sheetName As String)
    On Error Resume Next
    With idxSheet
        If no > 0 Then .Cells(idxRow, 1).Value = no
        .Cells(idxRow, 2).Value = pn
        .Cells(idxRow, 3).Value = desc
        If lvl > 0 Then .Cells(idxRow, 4).Value = lvl
        .Cells(idxRow, 5).Value = qty
        .Cells(idxRow, 6).Value = firstLevel
        .Cells(idxRow, 7).Value = pathTxt
        .Hyperlinks.Add .Cells(idxRow, 8), "", "'" & sheetName & "'!A1", "", sheetName
        .Cells(idxRow, 10).Value = "waiting ..."
    End With
    Err.Clear
End Sub

Private Sub WriteIndexResult(idxSheet As Object, ByVal idxRow As Long, ByVal nRows As Long, _
                             ByVal statusTxt As String, ByVal secs As Double, ByVal whereTxt As String)
    On Error Resume Next
    With idxSheet
        .Cells(idxRow, 9).Value = nRows
        .Cells(idxRow, 10).Value = statusTxt
        .Cells(idxRow, 11).Value = Round(secs, 1)
        .Cells(idxRow, 12).Value = whereTxt
        If statusTxt <> "OK" Then .Range(.Cells(idxRow, 1), .Cells(idxRow, 12)).Interior.Color = RGB(255, 220, 200)
    End With
    Err.Clear
End Sub

Private Sub FinishIndexLayout(idxSheet As Object)
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

    AskSplitLevel = 0
    If Not ASK_SPLIT_LEVEL Then Exit Function

    On Error Resume Next
    CATIA.StatusBar = "BOM: reading the product structure ..."
    Err.Clear
    On Error GoTo 0
    Call ScanLevels(rootProd, 1, cntAll, cntAssy, maxDepth)
    On Error Resume Next
    CATIA.StatusBar = ""
    Err.Clear
    On Error GoTo 0

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

        If IsNumeric(ans) Then
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
            nk = 0
            nk = child.Products.Count
            If nk > 0 Then
                cntAssy(lvl) = cntAssy(lvl) + 1
                Call ScanLevels(child, lvl + 1, cntAll, cntAssy, maxDepth)
            End If
        End If
    Next i
    Err.Clear
End Sub


' ==============================================================================
' HELPER: THE COMPONENTS ON THE SPLIT LEVEL
'   dUnit*        one entry per part number (first occurrence + count)
'   colUnitInst   EVERY instance on the split level (to hide them all)
'   colAbove      every assembly instance above the split level
'   colLoose*     parts above (or on) the split level that sit in no component
' Structure only - nothing is loaded here.
' ==============================================================================
Private Sub CollectUnits(oProd As Product, ByVal lvl As Long, ByVal target As Long, _
                         ByVal firstName As String, ByVal pathTxt As String, _
                         dUnitQty As Object, dUnitRef As Object, dUnitFirst As Object, _
                         dUnitDesc As Object, dUnitPath As Object, _
                         colUnitInst As Collection, colAbove As Collection, _
                         colLooseInst As Collection, colLooseLevel As Collection, _
                         colLooseFirst As Collection)
    Dim i As Long, n As Long, nk As Long
    Dim child As Product
    Dim pn As String, childFirst As String, sDesc As String, childPath As String

    On Error Resume Next
    n = 0
    n = oProd.Products.Count

    For i = 1 To n
        Set child = Nothing
        Set child = oProd.Products.Item(i)
        If Not child Is Nothing Then
            pn = ""
            pn = child.PartNumber
            If Len(pn) = 0 Then pn = child.Name

            If lvl = 1 Then
                childFirst = FirstLevelLabel(child, pn)
            Else
                childFirst = firstName
            End If

            nk = 0
            nk = child.Products.Count

            If nk > 0 And lvl = target Then
                ' ---- a component: one sheet per part number ----
                colUnitInst.Add child
                If dUnitQty.Exists(pn) Then
                    dUnitQty(pn) = dUnitQty(pn) + 1
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
                ' ---- an assembly above the split level: go deeper ----
                colAbove.Add child
                If Len(pathTxt) > 0 Then childPath = pathTxt & " > " & pn Else childPath = pn
                Call CollectUnits(child, lvl + 1, target, childFirst, childPath, _
                                  dUnitQty, dUnitRef, dUnitFirst, dUnitDesc, dUnitPath, _
                                  colUnitInst, colAbove, colLooseInst, colLooseLevel, colLooseFirst)

            Else
                ' ---- a part that is not inside any component ----
                colLooseInst.Add child
                colLooseLevel.Add lvl
                colLooseFirst.Add childFirst
            End If
        End If
    Next i
    Err.Clear
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
    On Error Resume Next
    xlBook.Application.DisplayAlerts = False
    For i = xlBook.Sheets.Count To 2 Step -1
        xlBook.Sheets(i).Delete
    Next i
    xlBook.Application.DisplayAlerts = True
    Err.Clear
End Sub

Private Function AddSheet(xlBook As Object, ByVal baseName As String) As Object
    Dim ws As Object
    Set ws = xlBook.Sheets.Add(, xlBook.Sheets(xlBook.Sheets.Count))
    ws.Name = SafeSheetName(xlBook, baseName)
    Set AddSheet = ws
End Function

' Excel sheet names: max 31 characters, none of  : \ / ? * [ ]  and unique
Private Function SafeSheetName(xlBook As Object, ByVal baseName As String) As String
    Dim s As String, cand As String, bad As Variant
    Dim k As Long, n As Long

    s = baseName
    For Each bad In Array(":", "\", "/", "?", "*", "[", "]")
        s = Replace(s, CStr(bad), "_")
    Next bad
    s = Trim$(s)
    If Len(s) = 0 Then s = "Sheet"
    If Len(s) > 31 Then s = Left$(s, 31)

    cand = s
    n = 1
    Do While SheetNameUsed(xlBook, cand)
        n = n + 1
        k = Len(CStr(n)) + 1
        cand = Left$(s, 31 - k) & "~" & n
    Loop
    SafeSheetName = cand
End Function

Private Function SheetNameUsed(xlBook As Object, ByVal nm As String) As Boolean
    Dim ws As Object
    SheetNameUsed = False
    For Each ws In xlBook.Sheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then
            SheetNameUsed = True
            Exit Function
        End If
    Next ws
End Function

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
' HELPER: TRAVERSE TREE (CALCULATES MASS + TRACKS LEVEL + DETECTS MULTI-BODY)
' Level 1 = direct children of root, Level 2 = grandchildren, etc.
' Also fills:
'   - colLeaves: EVERY leaf instance (used to hide them all in one call)
'   - dBodies: partNum -> array of solid body names (only when more than one)
' ==============================================================================
Function TraverseTree(oProd As Product, dQty As Object, dRef As Object, dDesc As Object, _
                      dProps As Object, dLevel As Object, dBodies As Object, _
                      dPartNum As Object, dFirst As Object, dIsAssy As Object, _
                      dNodeLeaves As Object, colLeaves As Collection, _
                      ByVal currentLevel As Integer, ByVal firstLevelName As String, _
                      ByVal parentKey As String) As Collection

    Dim childProd As Product, i As Integer, partNum As String
    Dim myLeaves As Collection, childLeaves As Collection
    Dim k As Long
    Dim sKey As String, childFirst As String

    Set myLeaves = New Collection
    Set TraverseTree = myLeaves          ' set early: an error still returns a list

    If oProd Is Nothing Then Exit Function
    If oProd.Products.Count = 0 Then Exit Function

    For i = 1 To oProd.Products.Count
        Set childProd = oProd.Products.Item(i)

        ' NOTE: no ApplyWorkMode here any more. Doing it per node forced every
        ' CATPart of the tree into design mode during the traversal, which is
        ' what made CATIA go "not responding" on deep trees. Each row promotes
        ' its own node in RunBOM instead.
        mAllNodes.Add childProd          ' assemblies and leaves, for the show/hide baseline

        partNum = ""
        On Error Resume Next
        partNum = childProd.PartNumber
        On Error GoTo 0

        ' The direct children of the root product ARE the first level; deeper
        ' nodes inherit the name of the first level branch they sit in.
        If currentLevel = 1 Then
            childFirst = FirstLevelLabel(childProd, partNum)
        Else
            childFirst = firstLevelName
        End If

        If childProd.Products.Count > 0 Then

            ' ---------- SUB-ASSEMBLY ----------
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
            If mIncludeSubassemblies And CAPTURE_ASSEMBLY_SHOTS And partNum <> "" Then
                If Not dNodeLeaves.Exists(sKey) Then dNodeLeaves.Add sKey, childLeaves
            End If

            If mIncludeSubassemblies And CAPTURE_ASSEMBLY_SHOTS Then
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
            If mIncludeSubassemblies And CAPTURE_ASSEMBLY_SHOTS Then myLeaves.Add childProd

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
    Next i
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
' HELPER: MASS / VOLUME / AREA OF A PRODUCT (works for parts and assemblies)
' Volume in m3, area in m2 - same units as before
' ==============================================================================
Sub ReadProductProps(oProd As Product, ByRef dMass As Double, ByRef dVol As Double, _
                     ByRef dArea As Double, ByVal isAssembly As Boolean)
    Dim oAnalyze As Object
    Dim oInertia As Object

    dMass = 0: dVol = 0: dArea = 0

    On Error Resume Next

    Set oAnalyze = oProd.Analyze
    If Not oAnalyze Is Nothing Then
        dMass = oAnalyze.Mass
        ' Volume and WetArea integrate over the whole subtree - on an assembly
        ' row that is the expensive part, so it can be switched off
        If (Not isAssembly) Or ASSEMBLY_VOLUME_AREA Then
            dVol = oAnalyze.Volume / (1000# ^ 3)
            dArea = oAnalyze.WetArea / (1000# ^ 2)
        End If
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
