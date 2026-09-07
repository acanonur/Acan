Option Explicit

' ==============================================================================
' MACRO: MASTER BOM v3 (FAST SCREENSHOTS + MULTI-PARTBODY SUPPORT)
'
' WHY THIS VERSION IS FAST AND QUIET:
'   - NO temporary geometry is created in the CATParts any more (MEASURE_MODE).
'     Creating a bounding box, running the update engine on it and deleting it
'     again is what made CATIA show an error panel for EVERY component - VBA
'     cannot suppress those panels, they come from CATIA itself. The dimensions
'     now come from the inertia, which only reads.
'   - The specification tree is switched off through the WINDOW LAYOUT, which
'     works in every language. The old StartCommand "Specification Tree Display"
'     only works on an English CATIA and silently did nothing anywhere else,
'     which is why the tree was still in every picture.
'   - Every wait pumps the message queue (Settle) instead of blocking the thread
'     with the kernel32 Sleep, so CATIA stays responsive and actually redraws
'     the view before it is captured.
'   - Selection.Add is used thousands of times, so HSOSynchronized is switched
'     off for the run - without that every single Add repaints the tree.
'   - Design mode is applied ONCE to the root, not once per node.
'   - Material and dimensions are measured ONCE per part number and cached.
'   - All leaf parts are hidden ONCE with a single batched SetShow call; per
'     part only the target is shown and hidden again.
'   - Excel ScreenUpdating is off while rows are written.
'
' IF IT IS TOO SLOW, the settings block below has the switches - each of them
' trades content for speed:
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

' Mass / volume / area of a sub-assembly row. Reading them means a recursive
' evaluation over the whole subtree, so they are the expensive part of the
' assembly structure version - but without them the assembly rows are empty,
' so they stay ON.
Private Const ASSEMBLY_METRICS As Boolean = True
' ... of the three, Volume and WetArea are the costly ones (a full B-Rep
' integration over the subtree). Set this to False to keep only the mass on
' assembly rows if the run is too slow.
Private Const ASSEMBLY_VOLUME_AREA As Boolean = True

' --- caches (filled during the run) ------------------------------------------
Private mDimCache As Object      ' part number -> Array(L, W, H)
Private mMatCache As Object      ' part number -> Array(material, density)
Private mBodyCache As Object     ' part number -> Variant array of body names
Private mPropCache As Object     ' part number -> Array(mass, volume, area)
Private mBBoxDisabled As Boolean ' the temp geometry path failed once - stop it


' ==============================================================================
' ENTRY POINT 1: PART LIST  (the classic BOM - same result as before)
' One row per part number, Qty = instances in the whole product, no assembly
' rows. Column N lists the first level branch(es) the part is used in.
' ==============================================================================
Sub GenerateMasterBOM()
    mIncludeSubassemblies = False
    mGroupBy = "GLOBAL"
    Call RunBOM
End Sub

' ==============================================================================
' ENTRY POINT 2: ASSEMBLY STRUCTURE
' Sub-assemblies get their own row and every assembly lists its own content,
' so a part can be located inside the tree. Longer list, same columns.
' ==============================================================================
Sub GenerateAssemblyBOM()
    mIncludeSubassemblies = True
    mGroupBy = "PARENT"
    Call RunBOM
End Sub


' ==============================================================================
' THE WORKER USED BY BOTH ENTRY POINTS
' ==============================================================================
Private Sub RunBOM()

    ' --- 1. SETUP & VARIABLES ---
    Dim catDoc As Document
    Dim rootProd As Product
    Dim sel As Selection

    ' Dictionaries (all keyed by the row key: first level + part number)
    Dim dictQty As Object, dictRef As Object, dictDesc As Object, dictProps As Object
    Dim dictLevel As Object, dictBodies As Object
    Dim dictPartNum As Object, dictFirst As Object
    Dim dictIsAssy As Object, dictNodeLeaves As Object

    ' Every leaf instance in the tree (needed to hide them all in one shot)
    Dim colLeaves As Collection

    ' Excel
    Dim xlApp As Object, xlBook As Object, xlSheet As Object

    ' Loop Vars
    Dim uniqueKey As Variant, r As Long, strPartNum As String
    Dim oPartProd As Product
    Dim propsArray As Variant
    Dim isAssy As Boolean
    Dim colNodeLeaves As Collection
    Dim nAssyRows As Long, nPartRows As Long
    Dim cacheArr As Variant
    Dim solidNames() As String
    Dim nSolid As Integer
    Dim dMass As Double, dVol As Double, dArea As Double

    ' Part data
    Dim oPart As Part
    Dim oPartDoc As PartDocument
    Dim hasPart As Boolean
    Dim dims(2) As Double
    Dim strMatName As String, dDensity As Double

    ' Body sub-rows
    Dim bodyNames As Variant
    Dim bi As Integer
    Dim oBody As Body
    Dim partSel As Selection
    Dim bMass As Double, bVol As Double, bArea As Double
    Dim bDens As Double, bDensInertia As Double
    Dim bMat As String

    ' Screenshot
    Dim tempPicPath As String, fso As Object
    Dim oViewer As Viewer

    ' Timing
    Dim startTime As Double
    startTime = Timer

    ' Window handling: the spec tree is switched off through the layout, so the
    ' handler has to know the previous layout to put it back
    Dim oWin As Object
    Dim savedLayout As Long
    Dim viewToggled As Boolean, compassToggled As Boolean

    ' Any unexpected error must not leave the model completely hidden
    On Error GoTo Fail

    ' --- 2. INITIALIZATION ---
    On Error Resume Next
    Set catDoc = CATIA.ActiveDocument
    If Err.Number <> 0 Then MsgBox "No Document.", vbCritical: Exit Sub
    On Error GoTo Fail

    If InStr(catDoc.Name, ".CATProduct") = 0 Then MsgBox "Open Assembly.", vbExclamation: Exit Sub

    Set rootProd = catDoc.Product
    Set sel = catDoc.Selection

    ' Suppress CATIA alert dialogs during the run (restored at the end)
    On Error Resume Next
    CATIA.DisplayFileAlerts = False
    On Error GoTo Fail

    Set fso = CreateObject("Scripting.FileSystemObject")
    tempPicPath = "C:\Temp\catia_bom_shot.jpg"
    If Not fso.FolderExists("C:\Temp") Then fso.CreateFolder ("C:\Temp")

    Set dictQty = CreateObject("Scripting.Dictionary")
    Set dictRef = CreateObject("Scripting.Dictionary")
    Set dictDesc = CreateObject("Scripting.Dictionary")
    Set dictProps = CreateObject("Scripting.Dictionary")
    Set dictLevel = CreateObject("Scripting.Dictionary")
    Set dictBodies = CreateObject("Scripting.Dictionary")
    Set dictPartNum = CreateObject("Scripting.Dictionary")
    Set dictFirst = CreateObject("Scripting.Dictionary")
    Set dictIsAssy = CreateObject("Scripting.Dictionary")
    Set dictNodeLeaves = CreateObject("Scripting.Dictionary")
    Set colLeaves = New Collection

    ' caches: the same part number is measured once, however often it is listed
    Set mDimCache = CreateObject("Scripting.Dictionary")
    Set mMatCache = CreateObject("Scripting.Dictionary")
    Set mBodyCache = CreateObject("Scripting.Dictionary")
    Set mPropCache = CreateObject("Scripting.Dictionary")
    mBBoxDisabled = False

    ' Selection.Add is used thousands of times; without this every single one
    ' repaints the tree and the highlight, which is what makes CATIA crawl
    On Error Resume Next
    CATIA.HSOSynchronized = False
    ' design mode ONCE for the whole tree instead of once per node
    If FORCE_DESIGN_MODE Then rootProd.ApplyWorkMode 2
    Err.Clear
    On Error GoTo Fail

    ' --- 3. TRAVERSE TREE (Count Parts, Calc Mass, Detect Multi-Body Parts) ---
    On Error Resume Next
    CATIA.RefreshDisplay = False
    On Error GoTo Fail

    Call TraverseTree(rootProd, dictQty, dictRef, dictDesc, dictProps, dictLevel, _
                      dictBodies, dictPartNum, dictFirst, dictIsAssy, dictNodeLeaves, _
                      colLeaves, 1, "", "")

    On Error Resume Next
    CATIA.RefreshDisplay = True
    On Error GoTo Fail

    If dictQty.Count = 0 Then
        ' nothing to write - but the two SESSION flags this run changed must be
        ' put back, otherwise CATIA stays without file alerts and without HSO
        ' synchronisation until it is restarted
        On Error Resume Next
        CATIA.HSOSynchronized = True
        CATIA.DisplayFileAlerts = True
        Err.Clear
        On Error GoTo Fail
        MsgBox "No parts found.", vbInformation
        Exit Sub
    End If

    ' --- 4. EXCEL SETUP ---
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then Set xlApp = CreateObject("Excel.Application")
    On Error GoTo Fail

    xlApp.Visible = True
    xlApp.ScreenUpdating = False
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)

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
    r = 2

    ' --- 5. SCREENSHOT SESSION SETUP (done ONCE, not per part) ---
    ' The specification tree is switched off through the window LAYOUT. The old
    ' StartCommand "Specification Tree Display" only works on an English CATIA -
    ' on any other language it silently does nothing and the tree stays in every
    ' picture.
    Set oViewer = Nothing
    On Error Resume Next
    Set oWin = CATIA.ActiveWindow
    Set oViewer = oWin.ActiveViewer

    savedLayout = -1
    savedLayout = oWin.Layout
    oWin.Layout = catWindowGeomOnly            ' geometry only - no spec tree
    viewToggled = (Err.Number = 0)
    Err.Clear

    ' the compass has no automation property, only a translated command name
    If Len(COMPASS_COMMAND) > 0 Then
        CATIA.StartCommand COMPASS_COMMAND
        compassToggled = (Err.Number = 0)
        Err.Clear
    End If

    If SET_ISO_VIEW Then Call SetIsoViewpoint(oViewer)
    Call Settle(100)
    On Error GoTo Fail

    ' Hide ALL leaf parts with ONE SetShow call.
    ' Parent nodes stay untouched, so showing one leaf later is enough.
    Call SetShowCollection(sel, colLeaves, 1)
    Call Settle(100)

    ' --- 6. PROCESS UNIQUE ROWS (sub-assemblies and parts, in tree order) ---
    For Each uniqueKey In dictQty.Keys

        strPartNum = dictPartNum(uniqueKey)
        Set oPartProd = dictRef(uniqueKey)
        isAssy = dictIsAssy(uniqueKey)

        If isAssy Then nAssyRows = nAssyRows + 1 Else nPartRows = nPartRows + 1

        ' Write Level + Basic Data + Pre-Calculated Mass
        xlSheet.Cells(r, 2).Value = dictLevel(uniqueKey)
        xlSheet.Cells(r, 3).Value = strPartNum
        xlSheet.Cells(r, 4).Value = dictDesc(uniqueKey)
        xlSheet.Cells(r, 5).Value = dictQty(uniqueKey)
        xlSheet.Cells(r, 14).Value = dictFirst(uniqueKey)
        If isAssy Then
            xlSheet.Cells(r, 15).Value = "Assembly"
            xlSheet.Range(xlSheet.Cells(r, 2), xlSheet.Cells(r, 15)).Font.Bold = True
        Else
            xlSheet.Cells(r, 15).Value = "Part"
        End If

        ' --- MASS / VOLUME / AREA (read once per part number, after loading) ---
        If mPropCache.Exists(strPartNum) Then
            propsArray = mPropCache(strPartNum)
        Else
            dMass = 0: dVol = 0: dArea = 0
            If (Not isAssy) Or ASSEMBLY_METRICS Then
                Call ReadProductProps(oPartProd, dMass, dVol, dArea, isAssy)
            End If
            propsArray = Array(dMass, dVol, dArea)
            mPropCache.Add strPartNum, propsArray
        End If

        If propsArray(0) > 0 Then xlSheet.Cells(r, 6).Value = propsArray(0) Else xlSheet.Cells(r, 6).Value = "N/A"
        If propsArray(1) > 0 Then xlSheet.Cells(r, 7).Value = propsArray(1) Else xlSheet.Cells(r, 7).Value = "N/A"
        If propsArray(2) > 0 Then xlSheet.Cells(r, 8).Value = propsArray(2) Else xlSheet.Cells(r, 8).Value = "N/A"

        ' Promote THIS node to design mode: mass, volume, area and the inertia
        ' box are all 0 on a component that is still in visualization mode.
        ' Doing it row by row instead of once over the whole tree keeps the
        ' interface alive between the rows.
        On Error Resume Next
        oPartProd.ApplyWorkMode 2
        Err.Clear
        On Error GoTo Fail

        ' Try to read part data without opening any window
        ' (a sub-assembly has no CATPart of its own)
        hasPart = False
        Set oPart = Nothing
        Set oPartDoc = Nothing
        If Not isAssy Then hasPart = TryGetLoadedPart(oPartProd, oPart, oPartDoc)

        ' multi-body detection moved out of the traversal: it needs a loaded
        ' CATPart, and doing it here keeps the traversal free of any loading
        If hasPart And DETECT_MULTIBODY And Not mBodyCache.Exists(strPartNum) Then
            mBodyCache.Add strPartNum, True
            nSolid = GetSolidBodyNames(oPart, solidNames)
            If nSolid > 1 Then
                If Not dictBodies.Exists(strPartNum) Then dictBodies.Add strPartNum, solidNames
            End If
        End If

        ' --- MATERIAL & DENSITY (measured once per part number) ---
        strMatName = "N/A"
        dDensity = 0
        If hasPart Then
            If mMatCache.Exists(strPartNum) Then
                cacheArr = mMatCache(strPartNum)
                strMatName = CStr(cacheArr(0))
                dDensity = CDbl(cacheArr(1))
            Else
                Call GetMaterialAndDensity(oPart, strMatName, dDensity)
                mMatCache.Add strPartNum, Array(strMatName, dDensity)
            End If
        End If

        xlSheet.Cells(r, 12).Value = strMatName
        If dDensity > 0 Then
            xlSheet.Cells(r, 13).Value = dDensity
        Else
            xlSheet.Cells(r, 13).Value = "N/A"
        End If

        ' --- DIMENSIONS (measured once per part number) ---
        dims(0) = 0: dims(1) = 0: dims(2) = 0

        If mDimCache.Exists(strPartNum) Then
            cacheArr = mDimCache(strPartNum)
            dims(0) = CDbl(cacheArr(0)): dims(1) = CDbl(cacheArr(1)): dims(2) = CDbl(cacheArr(2))
        Else
            ' only in "BOX" mode: this creates temporary geometry in the CATPart
            If hasPart And UCase$(MEASURE_MODE) = "BOX" Then Call GetBoundingBoxDims(oPart, dims)

            If dims(0) < 0.1 Then
                ' equivalent box from the inertia - reads only, creates nothing.
                ' ReferenceProduct itself can raise on a node that is not loaded,
                ' and that would abort the whole run with everything still hidden
                On Error Resume Next
                Call GetInertiaDims(oPartProd.ReferenceProduct, dims)
                Err.Clear
                On Error GoTo Fail
            End If
            mDimCache.Add strPartNum, Array(dims(0), dims(1), dims(2))
        End If

        ' numbers, not Format() strings - otherwise the columns are text
        If dims(0) > 0 Then xlSheet.Cells(r, 9).Value = dims(0) Else xlSheet.Cells(r, 9).Value = "N/A"
        If dims(1) > 0 Then xlSheet.Cells(r, 10).Value = dims(1) Else xlSheet.Cells(r, 10).Value = "N/A"
        If dims(2) > 0 Then xlSheet.Cells(r, 11).Value = dims(2) Else xlSheet.Cells(r, 11).Value = "N/A"

        ' --- SCREENSHOT ------------------------------------------------------
        ' Part: show only that instance. Sub-assembly: show all leaf instances
        ' below it (they were hidden one by one, so the node alone stays empty).
        If isAssy Then
            Set colNodeLeaves = Nothing
            If dictNodeLeaves.Exists(uniqueKey) Then Set colNodeLeaves = dictNodeLeaves(uniqueKey)

            If Not CAPTURE_ASSEMBLY_SHOTS Then
                xlSheet.Cells(r, 1).Value = "-"
            ElseIf colNodeLeaves Is Nothing Then
                xlSheet.Cells(r, 1).Value = "No Preview"
            ElseIf colNodeLeaves.Count = 0 Then
                ' nothing with geometry below this node - the scene would be empty
                xlSheet.Cells(r, 1).Value = "No Preview"
            ElseIf colNodeLeaves.Count > MAX_ASSY_LEAVES_FOR_SHOT Then
                ' showing and hiding thousands of leaves for one picture would
                ' cost more than the whole rest of the run
                xlSheet.Cells(r, 1).Value = colNodeLeaves.Count & " parts"
            Else
                Call SetShowCollection(sel, colNodeLeaves, 0)   ' 0 = SHOW
                Call Settle(ASSY_SETTLE_MS)                     ' a whole branch has to redraw
                Call CaptureViewToExcel(oViewer, xlSheet, r, tempPicPath, fso)
                Call SetShowCollection(sel, colNodeLeaves, 1)   ' 1 = HIDE again
            End If
        ElseIf Not CAPTURE_PART_SHOTS Then
            xlSheet.Cells(r, 1).Value = "-"
        ElseIf Not oPartProd Is Nothing Then
            Call SetShowSingle(sel, oPartProd, 0)   ' 0 = SHOW
            Call Settle(PART_SETTLE_MS)
            Call CaptureViewToExcel(oViewer, xlSheet, r, tempPicPath, fso)
        Else
            xlSheet.Cells(r, 1).Value = "No Preview"
        End If

        ' --- MULTI-BODY SUB-ROWS (part stays visible while its bodies cycle) ---
        If hasPart And dictBodies.Exists(strPartNum) And Not oPartProd Is Nothing Then

            bodyNames = dictBodies(strPartNum)
            Set partSel = oPartDoc.Selection

            ' Hide all solid bodies of this part in one shot
            Call SetShowBodyList(oPart, partSel, bodyNames, 1)

            For bi = LBound(bodyNames) To UBound(bodyNames)
                Set oBody = GetBodyByName(oPart, CStr(bodyNames(bi)))
                If Not oBody Is Nothing Then
                    r = r + 1

                    ' Tree info: body behaves like a sub-part, one level deeper
                    xlSheet.Cells(r, 2).Value = dictLevel(uniqueKey) + 1
                    xlSheet.Cells(r, 3).Value = "    " & CStr(bodyNames(bi))
                    xlSheet.Cells(r, 4).Value = "Body of " & strPartNum
                    xlSheet.Cells(r, 5).Value = dictQty(uniqueKey)
                    xlSheet.Cells(r, 14).Value = dictFirst(uniqueKey)
                    xlSheet.Cells(r, 15).Value = "Body"

                    ' Show ONLY this body and capture it
                    ' (it stays visible until measurements are done - the
                    '  bounding box cannot be built on a hidden body)
                    partSel.Clear
                    partSel.Add oBody
                    partSel.VisProperties.SetShow 0   ' 0 = SHOW
                    Call Settle(PART_SETTLE_MS)
                    Call CaptureViewToExcel(oViewer, xlSheet, r, tempPicPath, fso)

                    ' --- BODY MEASUREMENTS ---
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
                            ' Fallback: true axis-aligned box from extremum points
                            Call MeasureAABBExtremum(oPart, oPartDoc, oBody, dims)
                        End If
                    End If
                    ' equivalent box from the body inertia - creates nothing
                    If dims(0) < 0.1 Then Call GetBodyInertiaDims(oPartDoc, oBody, dims)

                    If dims(0) > 0 Then xlSheet.Cells(r, 9).Value = dims(0) Else xlSheet.Cells(r, 9).Value = "N/A"
                    If dims(1) > 0 Then xlSheet.Cells(r, 10).Value = dims(1) Else xlSheet.Cells(r, 10).Value = "N/A"
                    If dims(2) > 0 Then xlSheet.Cells(r, 11).Value = dims(2) Else xlSheet.Cells(r, 11).Value = "N/A"

                    xlSheet.Cells(r, 12).Value = bMat
                    If bDens > 0 Then
                        xlSheet.Cells(r, 13).Value = bDens
                    Else
                        xlSheet.Cells(r, 13).Value = "N/A"
                    End If

                    ' Hide this body again before moving to the next one
                    partSel.Clear
                    partSel.Add oBody
                    partSel.VisProperties.SetShow 1   ' 1 = HIDE
                End If
                Set oBody = Nothing
            Next bi

            ' Restore all bodies of this part to visible
            Call SetShowBodyList(oPart, partSel, bodyNames, 0)
        End If

        ' Hide the part again before moving to the next one.
        ' A sub-assembly node is NOT hidden here - its leaves were already
        ' hidden again right after the screenshot, and hiding the node itself
        ' would leave it invisible in the session.
        If Not isAssy And Not oPartProd Is Nothing Then
            Call SetShowSingle(sel, oPartProd, 1)   ' 1 = HIDE
        End If

        r = r + 1

        ' keep CATIA alive and show progress - a run without this looks hung
        On Error Resume Next
        CATIA.StatusBar = "BOM: row " & (r - 2) & " of " & dictQty.Count & " ..."
        Err.Clear
        On Error GoTo Fail
        DoEvents
    Next uniqueKey

    ' --- 7. RESTORE SCENE (done ONCE) ---
    Call SetShowCollection(sel, colLeaves, 0)   ' show all leaf parts again

    On Error Resume Next
    If viewToggled And savedLayout >= 0 Then oWin.Layout = savedLayout
    If compassToggled And Len(COMPASS_COMMAND) > 0 Then CATIA.StartCommand COMPASS_COMMAND
    If Not oViewer Is Nothing Then oViewer.Reframe
    CATIA.HSOSynchronized = True
    CATIA.DisplayFileAlerts = True
    CATIA.StatusBar = ""
    Err.Clear
    On Error GoTo Fail
    sel.Clear

    ' --- 8. FINALIZE EXCEL ---
    xlApp.ScreenUpdating = True
    xlSheet.Columns("B:O").AutoFit
    xlSheet.Columns("A:A").ColumnWidth = 15
    xlSheet.Columns("N:N").ColumnWidth = 22

    ' filter on the header row: column N = first level, column O = Assembly/Part
    On Error Resume Next
    xlSheet.Range("B1:O1").AutoFilter
    On Error GoTo Fail

    Dim modeName As String
    If mIncludeSubassemblies Then modeName = "Assembly structure" Else modeName = "Part list"

    MsgBox "BOM Exported Successfully!  (" & modeName & ")" & vbCrLf & _
           "Rows written:     " & (r - 2) & vbCrLf & _
           "  sub-assemblies: " & nAssyRows & vbCrLf & _
           "  parts:          " & nPartRows & vbCrLf & _
           "Time elapsed: " & Format(Timer - startTime, "0.0") & " seconds", vbInformation
    Exit Sub

' ------------------------------------------------------------------------------
' Something went wrong: put the model, CATIA and Excel back into a usable state
' before reporting. Without this the run would leave every part hidden.
' ------------------------------------------------------------------------------
Fail:
    Dim failMsg As String
    failMsg = "The BOM run stopped with error " & Err.Number & " - " & Err.Description

    On Error Resume Next
    Call SetShowCollection(sel, colLeaves, 0)         ' show all leaf parts again
    If viewToggled And savedLayout >= 0 Then oWin.Layout = savedLayout
    If compassToggled And Len(COMPASS_COMMAND) > 0 Then CATIA.StartCommand COMPASS_COMMAND
    If Not oViewer Is Nothing Then oViewer.Reframe
    CATIA.HSOSynchronized = True
    CATIA.RefreshDisplay = True
    CATIA.DisplayFileAlerts = True
    CATIA.StatusBar = ""
    If Not sel Is Nothing Then sel.Clear
    If Not xlApp Is Nothing Then xlApp.ScreenUpdating = True
    Err.Clear

    MsgBox failMsg & vbCrLf & vbCrLf & _
           "Visibility, spec tree and compass were restored." & vbCrLf & _
           "The rows written so far are in the Excel sheet.", vbCritical

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
    Next i
    sel.VisProperties.SetShow showMode
    sel.Clear
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
' HELPER: REFRAME + CAPTURE CURRENT VIEW INTO EXCEL CELL A(row)
' ==============================================================================
Sub CaptureViewToExcel(oViewer As Viewer, xlSheet As Object, ByVal row As Long, _
                       tempPicPath As String, fso As Object)
    On Error Resume Next
    Dim shp As Object

    If oViewer Is Nothing Then
        xlSheet.Cells(row, 1).Value = "No Viewer"
        Exit Sub
    End If

    oViewer.Reframe
    oViewer.Update
    Call Settle(PART_SETTLE_MS)

    If fso.FileExists(tempPicPath) Then fso.DeleteFile tempPicPath
    oViewer.CaptureToFile 4, tempPicPath

    If fso.FileExists(tempPicPath) Then
        Set shp = xlSheet.Shapes.AddPicture(tempPicPath, False, True, _
            xlSheet.Cells(row, 1).Left + 2, xlSheet.Cells(row, 1).Top + 2, -1, -1)
        If Not shp Is Nothing Then
            shp.Height = 50
            If shp.Width > 90 Then shp.Width = 90
        End If
        xlSheet.Rows(row).RowHeight = 60
    Else
        xlSheet.Cells(row, 1).Value = "No Preview"
    End If
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
        ' what made CATIA go "not responding" on deep trees. It is done once on
        ' the root in RunBOM instead.
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
