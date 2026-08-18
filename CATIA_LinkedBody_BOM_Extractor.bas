Option Explicit

' ==============================================================================
' MACRO: LINKED-BODY BOM EXTRACTOR (for a single CATPart built from links)
'
' For CATParts whose tree is a flat list of bodies that were pasted with link
' from an assembly, so every body carries the instance path of its source part
' in its name, e.g.
'
'   2DP000064819.1\2DP000033345.1\2DP000033352.1\2DP000033416.1\ ...
'                                     ... \2DP000034948.1\PartBody
'   2DP000064819.1\2DP000033345.1\ ... \2DP000033444.1\Inner-Winding
'   2DP000064819.1\2DP000033345.1\ ... \2DP000033444.1\PANEL-INTERFACE-WINDING.1
'
' The macro produces EXACTLY the same sheet as the assembly extractor
' (CATIA_BOM_Extractor_Fast_MultiBody.bas), so the result can be imported into
' the ACE cost model with the ACE_BOM_Import macro without any change:
'
'   A Thumbnail | B Level | C Part Number | D Description | E Qty |
'   F Mass (kg) | G Volume (m3) | H Area (m2) |
'   I Length (mm) | J Width (mm) | K Height (mm) |
'   L Material | M Density (kg/m3)
'
' How the tree is read:
'   * Level        = number of part numbers in the body name path
'                    (the example above -> Level 6)
'   * Part Number  = deepest part number of the path, without the instance
'                    suffix (2DP000034948.1 -> 2DP000034948)
'   * Description  = name of the body itself (PartBody, Inner-Winding,
'                    PANEL-INTERFACE-WINDING, ...)
'   * Qty          = number of identical bodies that were merged into the row
'   * bodies without a path in their name are reported on Level 1 with the
'     part number of the CATPart itself
'
' Identical bodies (same source path, same body name, same mass and volume)
' are merged into ONE row with Qty > 1 - that is what turns the six
' PANEL-INTERFACE-WINDING.1 ... .6 bodies into one row with Qty 6.
' Set GROUP_IDENTICAL_BODIES = False for one row per body.
'
' Mass, volume, area, bounding box and material are measured per body with the
' same routines and the same fallbacks as the assembly extractor.
'
' Usage:  open the CATPart in CATIA -> Tools -> Macro -> ExtractLinkedBodyBOM
' ==============================================================================

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' --- settings -----------------------------------------------------------------
Private Const GROUP_IDENTICAL_BODIES As Boolean = True   ' merge equal bodies, count Qty
Private Const GROUP_BY_GEOMETRY As Boolean = True        ' only merge if mass+volume match
Private Const CAPTURE_THUMBNAILS As Boolean = True       ' False = much faster, no pictures
Private Const INCLUDE_GEOMETRICAL_SETS As Boolean = False ' also list surface sets
Private Const WRITE_ROOT_ROW As Boolean = False          ' extra Level 1 row for the CATPart
Private Const WRITE_SOURCE_PATH As Boolean = False       ' full body name into column N
Private Const TEMP_FOLDER As String = "C:\Temp"


' ==============================================================================
' MAIN
' ==============================================================================
Sub ExtractLinkedBodyBOM()

    ' --- CATIA ---
    Dim catDoc As Document
    Dim partDoc As PartDocument
    Dim oPart As Part
    Dim sel As Selection
    Dim oViewer As Viewer
    Dim oItem As Object

    ' --- Excel ---
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim r As Long

    ' --- body table (parallel arrays, index 1..nBodies) ---
    Dim colBodies As Collection
    Dim aFullName() As String, aPartNum() As String, aDesc() As String
    Dim aPathKey() As String, aKey() As String
    Dim aLevel() As Long, aQty() As Long
    Dim aMass() As Double, aVol() As Double, aArea() As Double
    Dim aDimL() As Double, aDimW() As Double, aDimH() As Double
    Dim aMat() As String, aDens() As Double
    Dim aPic() As String
    Dim aIsRep() As Boolean

    Dim nBodies As Long, i As Long, nRows As Long, nSkipped As Long
    Dim dims(2) As Double
    Dim dMass As Double, dVol As Double, dArea As Double
    Dim dDens As Double, dDensInertia As Double, dMatDens As Double
    Dim sMat As String
    Dim rootPartNum As String

    Dim dictGroup As Object
    Dim repIndex As Long
    Dim sKey As String

    Dim fso As Object
    Dim picPath As String

    Dim startTime As Double
    startTime = Timer

    ' --- 1. DOCUMENT CHECK ----------------------------------------------------
    On Error Resume Next
    Set catDoc = CATIA.ActiveDocument
    If Err.Number <> 0 Or catDoc Is Nothing Then
        MsgBox "No document open.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    If InStr(1, catDoc.Name, ".CATPart", vbTextCompare) = 0 Then
        MsgBox "Please open the CATPart that contains the linked bodies." & vbCrLf & vbCrLf & _
               "For a CATProduct use the assembly extractor (GenerateMasterBOM).", vbExclamation
        Exit Sub
    End If

    Set partDoc = catDoc
    Set oPart = partDoc.Part
    Set sel = catDoc.Selection
    rootPartNum = LBRootPartNumber(partDoc, oPart)

    ' --- 2. COLLECT THE BODIES ------------------------------------------------
    Set colBodies = LBCollectBodies(oPart)
    nBodies = colBodies.Count

    If nBodies = 0 Then
        MsgBox "No bodies with geometry found in this CATPart.", vbInformation
        Exit Sub
    End If

    ReDim aFullName(1 To nBodies): ReDim aPartNum(1 To nBodies): ReDim aDesc(1 To nBodies)
    ReDim aPathKey(1 To nBodies): ReDim aKey(1 To nBodies)
    ReDim aLevel(1 To nBodies): ReDim aQty(1 To nBodies)
    ReDim aMass(1 To nBodies): ReDim aVol(1 To nBodies): ReDim aArea(1 To nBodies)
    ReDim aDimL(1 To nBodies): ReDim aDimW(1 To nBodies): ReDim aDimH(1 To nBodies)
    ReDim aMat(1 To nBodies): ReDim aDens(1 To nBodies)
    ReDim aPic(1 To nBodies): ReDim aIsRep(1 To nBodies)

    ' --- 3. SCREENSHOT SESSION SETUP (once for the whole run) -----------------
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(TEMP_FOLDER) Then fso.CreateFolder (TEMP_FOLDER)

    On Error Resume Next
    CATIA.DisplayFileAlerts = False
    Set oViewer = CATIA.ActiveWindow.ActiveViewer
    If CAPTURE_THUMBNAILS Then
        CATIA.StartCommand "Specification Tree Display"   ' toggle off
        CATIA.StartCommand "Compass"                      ' toggle off
        Sleep 100
    End If
    On Error GoTo 0

    ' hide every body once - afterwards only the measured one is shown
    Call LBSetShowCollection(sel, colBodies, 1)
    Sleep 100

    ' --- 4. MEASURE EVERY BODY ------------------------------------------------
    For i = 1 To nBodies
        Set oItem = colBodies.Item(i)

        aFullName(i) = LBNameOf(oItem)
        Call LBParseBodyName(aFullName(i), rootPartNum, _
                             aPartNum(i), aDesc(i), aPathKey(i), aLevel(i))

        ' show only this body
        Call LBSetShowSingle(sel, oItem, 0)
        Sleep 50

        ' screenshot FIRST - the measurements below build temporary geometry
        If CAPTURE_THUMBNAILS Then
            picPath = TEMP_FOLDER & "\catia_lbom_" & Format(i, "0000") & ".jpg"
            If LBCaptureViewToFile(oViewer, picPath, fso) Then aPic(i) = picPath
        End If

        ' mass / density / volume / area
        dMass = 0: dVol = 0: dArea = 0
        dDens = 0: dDensInertia = 0: dMatDens = 0
        sMat = "N/A"

        Call LBGetBodyMaterial(oPart, oItem, sMat, dMatDens)
        Call LBGetBodyMassAndDensity(partDoc, oItem, dMass, dDensInertia)
        dDens = dMatDens
        If dDens <= 0 Then dDens = dDensInertia
        Call LBGetBodyVolumeArea(partDoc, oPart, oItem, dVol, dArea)
        If dMass <= 0 And dDens > 0 And dVol > 0 Then dMass = dVol * dDens

        aMass(i) = dMass: aVol(i) = dVol: aArea(i) = dArea
        aMat(i) = sMat: aDens(i) = dDens

        ' bounding box with the same three-step fallback as the assembly macro
        dims(0) = 0: dims(1) = 0: dims(2) = 0
        If Not LBMeasureTarget(oPart, oItem, dims) Then
            If Not LBMeasureAABBExtremum(oPart, partDoc, oItem, dims) Then
                Call LBGetBodyInertiaDims(partDoc, oItem, dims)
            End If
        End If
        aDimL(i) = dims(0): aDimW(i) = dims(1): aDimH(i) = dims(2)

        ' hide it again
        Call LBSetShowSingle(sel, oItem, 1)

        On Error Resume Next
        CATIA.StatusBar = "BOM: body " & i & " of " & nBodies & " ..."
        On Error GoTo 0
        DoEvents
    Next i

    ' --- 5. RESTORE THE SCENE (once) ------------------------------------------
    Call LBSetShowCollection(sel, colBodies, 0)

    On Error Resume Next
    If CAPTURE_THUMBNAILS Then
        CATIA.StartCommand "Specification Tree Display"   ' toggle on
        CATIA.StartCommand "Compass"                      ' toggle on
    End If
    If Not oViewer Is Nothing Then oViewer.Reframe
    CATIA.DisplayFileAlerts = True
    CATIA.StatusBar = ""
    On Error GoTo 0
    sel.Clear

    ' --- 6. GROUP IDENTICAL BODIES -------------------------------------------
    Set dictGroup = CreateObject("Scripting.Dictionary")

    For i = 1 To nBodies
        aQty(i) = 1
        aIsRep(i) = True

        If GROUP_IDENTICAL_BODIES Then
            sKey = aPathKey(i) & "|" & UCase$(LBBaseName(aDesc(i)))
            If GROUP_BY_GEOMETRY Then
                sKey = sKey & "|" & Format(aMass(i), "0.000000") & _
                              "|" & Format(aVol(i), "0.000000000")
            End If
            aKey(i) = sKey

            If dictGroup.Exists(sKey) Then
                repIndex = dictGroup(sKey)
                aQty(repIndex) = aQty(repIndex) + 1
                aQty(i) = 0
                aIsRep(i) = False
                nSkipped = nSkipped + 1
            Else
                dictGroup.Add sKey, i
                ' the row stands for a group -> show the name without the
                ' instance number (PANEL-INTERFACE-WINDING.1 -> ...WINDING)
                aDesc(i) = LBBaseName(aDesc(i))
            End If
        End If
    Next i

    ' --- 7. EXCEL -------------------------------------------------------------
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Or xlApp Is Nothing Then Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0

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
        If WRITE_SOURCE_PATH Then .Cells(1, 14).Value = "Source"

        .Range("A1:M1").Font.Bold = True
        .Range("A1:M1").Interior.Color = RGB(220, 220, 220)
        .Columns("A:A").ColumnWidth = 15
        .Columns("B:B").ColumnWidth = 8
        .Columns("C:D").ColumnWidth = 20
        .Columns("F:H").NumberFormat = "0.000000000"
        .Columns("M:M").NumberFormat = "0.000"
    End With
    r = 2

    ' optional row for the CATPart itself
    If WRITE_ROOT_ROW Then
        xlSheet.Cells(r, 2).Value = 1
        xlSheet.Cells(r, 3).Value = rootPartNum
        xlSheet.Cells(r, 4).Value = LBBaseFileName(catDoc.Name)
        xlSheet.Cells(r, 5).Value = 1
        xlSheet.Cells(r, 6).Value = LBTotalMass(aMass, aQty, nBodies)
        xlSheet.Cells(r, 12).Value = "N/A"
        xlSheet.Cells(r, 13).Value = "N/A"
        r = r + 1
    End If

    For i = 1 To nBodies
        If aIsRep(i) Then
            xlSheet.Cells(r, 2).Value = aLevel(i)
            xlSheet.Cells(r, 3).Value = aPartNum(i)
            xlSheet.Cells(r, 4).Value = aDesc(i)
            xlSheet.Cells(r, 5).Value = aQty(i)

            If aMass(i) > 0 Then xlSheet.Cells(r, 6).Value = aMass(i) Else xlSheet.Cells(r, 6).Value = "N/A"
            If aVol(i) > 0 Then xlSheet.Cells(r, 7).Value = aVol(i) Else xlSheet.Cells(r, 7).Value = "N/A"
            If aArea(i) > 0 Then xlSheet.Cells(r, 8).Value = aArea(i) Else xlSheet.Cells(r, 8).Value = "N/A"

            If aDimL(i) > 0 Then xlSheet.Cells(r, 9).Value = Format(aDimL(i), "0.0") Else xlSheet.Cells(r, 9).Value = "N/A"
            If aDimW(i) > 0 Then xlSheet.Cells(r, 10).Value = Format(aDimW(i), "0.0") Else xlSheet.Cells(r, 10).Value = "N/A"
            If aDimH(i) > 0 Then xlSheet.Cells(r, 11).Value = Format(aDimH(i), "0.0") Else xlSheet.Cells(r, 11).Value = "N/A"

            xlSheet.Cells(r, 12).Value = aMat(i)
            If aDens(i) > 0 Then
                xlSheet.Cells(r, 13).Value = aDens(i)
            Else
                xlSheet.Cells(r, 13).Value = "N/A"
            End If

            If WRITE_SOURCE_PATH Then xlSheet.Cells(r, 14).Value = aFullName(i)

            If CAPTURE_THUMBNAILS Then
                Call LBInsertPicture(xlSheet, r, aPic(i), fso)
            End If

            nRows = nRows + 1
            r = r + 1
        End If
    Next i

    ' --- 8. FINALIZE ----------------------------------------------------------
    xlApp.ScreenUpdating = True
    xlSheet.Columns("B:M").AutoFit
    xlSheet.Columns("A:A").ColumnWidth = 15

    ' remove the temporary screenshots
    On Error Resume Next
    For i = 1 To nBodies
        If Len(aPic(i)) > 0 Then
            If fso.FileExists(aPic(i)) Then fso.DeleteFile aPic(i)
        End If
    Next i
    Err.Clear
    On Error GoTo 0

    MsgBox "BOM Exported Successfully!" & vbCrLf & _
           "Bodies found:  " & nBodies & vbCrLf & _
           "Rows written:  " & nRows & vbCrLf & _
           "Merged as Qty: " & nSkipped & vbCrLf & _
           "Time elapsed:  " & Format(Timer - startTime, "0.0") & " seconds", vbInformation

End Sub


' ==============================================================================
' HELPER: ALL BODIES WITH GEOMETRY (solid bodies, optionally surface sets)
' ==============================================================================
Private Function LBCollectBodies(oPart As Part) As Collection
    On Error Resume Next
    Dim col As Collection
    Dim i As Long
    Dim oBody As Body
    Dim oHB As HybridBody

    Set col = New Collection

    For i = 1 To oPart.Bodies.Count
        Set oBody = oPart.Bodies.Item(i)
        If Not oBody Is Nothing Then
            If oBody.Shapes.Count > 0 Then col.Add oBody
        End If
        Set oBody = Nothing
    Next i

    If INCLUDE_GEOMETRICAL_SETS Then
        For i = 1 To oPart.HybridBodies.Count
            Set oHB = oPart.HybridBodies.Item(i)
            If Not oHB Is Nothing Then
                If InStr(1, oHB.Name, "TMP_BOM_MEASURE", vbTextCompare) = 0 Then
                    If oHB.HybridShapes.Count > 0 Then col.Add oHB
                End If
            End If
            Set oHB = Nothing
        Next i
    End If

    Set LBCollectBodies = col
    Err.Clear
End Function


' ==============================================================================
' HELPER: SPLIT A LINKED BODY NAME INTO PART NUMBER / DESCRIPTION / LEVEL
'
'   "2DP000064819.1\2DP000033345.1\2DP000034948.1\PartBody"
'      -> partNum = "2DP000034948"
'      -> desc    = "PartBody"
'      -> pathKey = "2DP000064819\2DP000033345\2DP000034948"   (for grouping)
'      -> level   = 3
'
' A body without a path (a normal, locally modelled body) is reported on
' Level 1 with the part number of the CATPart itself.
' ==============================================================================
Private Sub LBParseBodyName(ByVal fullName As String, ByVal rootPartNum As String, _
                            ByRef partNum As String, ByRef desc As String, _
                            ByRef pathKey As String, ByRef level As Long)
    Dim s As String
    Dim segs As Variant
    Dim i As Long, n As Long

    s = Trim$(fullName)
    s = Replace(s, "/", "\")

    If InStr(s, "\") = 0 Then
        partNum = rootPartNum
        desc = s
        pathKey = rootPartNum
        level = 1
        Exit Sub
    End If

    segs = Split(s, "\")
    n = UBound(segs)                      ' last element = body name

    desc = Trim$(CStr(segs(n)))
    partNum = LBBaseName(Trim$(CStr(segs(n - 1))))
    level = n                             ' number of part numbers in the path

    pathKey = ""
    For i = 0 To n - 1
        If i > 0 Then pathKey = pathKey & "\"
        pathKey = pathKey & LBBaseName(Trim$(CStr(segs(i))))
    Next i
End Sub


' ==============================================================================
' HELPER: NAME WITHOUT THE TRAILING INSTANCE NUMBER
'   "2DP000034948.1"              -> "2DP000034948"
'   "PANEL-INTERFACE-WINDING.6"   -> "PANEL-INTERFACE-WINDING"
'   "LN9499_09200_1_4301_INSERT"  -> unchanged
' ==============================================================================
Private Function LBBaseName(ByVal s As String) As String
    Dim p As Long, tail As String

    LBBaseName = s
    p = InStrRev(s, ".")
    If p <= 1 Then Exit Function

    tail = Mid$(s, p + 1)
    If Len(tail) = 0 Then Exit Function
    If Not IsNumeric(tail) Then Exit Function
    If InStr(tail, ",") > 0 Or InStr(tail, ".") > 0 Then Exit Function

    LBBaseName = Left$(s, p - 1)
End Function


' ==============================================================================
' HELPER: PART NUMBER OF THE OPEN CATPART
' ==============================================================================
Private Function LBRootPartNumber(partDoc As PartDocument, oPart As Part) As String
    On Error Resume Next
    Dim s As String

    s = ""
    s = partDoc.Product.PartNumber
    If Len(s) = 0 Then s = oPart.Name
    If Len(s) = 0 Then s = LBBaseFileName(partDoc.Name)

    LBRootPartNumber = s
    Err.Clear
End Function


' ==============================================================================
' HELPER: FILE NAME WITHOUT EXTENSION
' ==============================================================================
Private Function LBBaseFileName(ByVal s As String) As String
    Dim p As Long

    p = InStrRev(s, ".")
    If p > 1 Then
        LBBaseFileName = Left$(s, p - 1)
    Else
        LBBaseFileName = s
    End If
End Function


' ==============================================================================
' HELPER: NAME OF A BODY / HYBRID BODY
' ==============================================================================
Private Function LBNameOf(oItem As Object) As String
    On Error Resume Next
    LBNameOf = ""
    LBNameOf = oItem.Name
    Err.Clear
End Function


' ==============================================================================
' HELPER: TOTAL MASS OF ALL ROWS (for the optional root row)
' ==============================================================================
Private Function LBTotalMass(aMass() As Double, aQty() As Long, ByVal n As Long) As Double
    Dim i As Long, t As Double

    For i = 1 To n
        If aQty(i) > 0 Then t = t + aMass(i) * aQty(i)
    Next i
    LBTotalMass = t
End Function


' ==============================================================================
' HELPER: BATCHED SHOW / HIDE   (SetShow 0 = SHOW, 1 = HIDE)
' ==============================================================================
Private Sub LBSetShowCollection(sel As Selection, colItems As Collection, ByVal showMode As Integer)
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
Private Sub LBSetShowSingle(sel As Selection, oObj As Object, ByVal showMode As Integer)
    On Error Resume Next
    If oObj Is Nothing Then Exit Sub
    sel.Clear
    sel.Add oObj
    sel.VisProperties.SetShow showMode
    sel.Clear
    Err.Clear
End Sub


' ==============================================================================
' HELPER: REFRAME + CAPTURE THE CURRENT VIEW INTO A FILE
' ==============================================================================
Private Function LBCaptureViewToFile(oViewer As Viewer, ByVal picPath As String, fso As Object) As Boolean
    On Error Resume Next
    LBCaptureViewToFile = False

    If oViewer Is Nothing Then Exit Function

    oViewer.Reframe
    oViewer.Update
    Sleep 50

    If fso.FileExists(picPath) Then fso.DeleteFile picPath
    oViewer.CaptureToFile 4, picPath

    LBCaptureViewToFile = fso.FileExists(picPath)
    Err.Clear
End Function


' ==============================================================================
' HELPER: PUT A CAPTURED PICTURE INTO CELL A(row)
' Same size as the assembly extractor: height 50, max width 90, row height 60
' ==============================================================================
Private Sub LBInsertPicture(xlSheet As Object, ByVal row As Long, _
                            ByVal picPath As String, fso As Object)
    On Error Resume Next
    Dim shp As Object

    If Len(picPath) = 0 Then
        xlSheet.Cells(row, 1).Value = "No Preview"
        Exit Sub
    End If
    If Not fso.FileExists(picPath) Then
        xlSheet.Cells(row, 1).Value = "No Preview"
        Exit Sub
    End If

    Set shp = xlSheet.Shapes.AddPicture(picPath, False, True, _
        xlSheet.Cells(row, 1).Left + 2, xlSheet.Cells(row, 1).Top + 2, -1, -1)
    If Not shp Is Nothing Then
        shp.Height = 50
        If shp.Width > 90 Then shp.Width = 90
    End If
    xlSheet.Rows(row).RowHeight = 60
    Err.Clear
End Sub


' ==============================================================================
' HELPER: BODY MASS & DENSITY via SPAWorkbench Inertias
' ==============================================================================
Private Sub LBGetBodyMassAndDensity(oPartDoc As PartDocument, oBody As Object, _
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
' HELPER: BODY VOLUME & AREA via SPA Workbench Measurable (m3 / m2)
' ==============================================================================
Private Sub LBGetBodyVolumeArea(oPartDoc As PartDocument, oPart As Part, oBody As Object, _
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
Private Sub LBGetBodyMaterial(oPart As Part, oBody As Object, _
                              ByRef matName As String, ByRef density As Double)
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
' HELPER: BOUNDING BOX ON A BODY (temporary box inside a temporary set)
' ==============================================================================
Private Function LBMeasureTarget(oPart As Part, oTargetObj As Object, ByRef dDims() As Double) As Boolean
    On Error Resume Next
    Err.Clear
    LBMeasureTarget = False

    Dim oHSF As Object: Set oHSF = oPart.HybridShapeFactory
    Dim oRef As Reference: Set oRef = oPart.CreateReferenceFromObject(oTargetObj)
    Dim oBox As Object: Set oBox = oHSF.AddNewBoundingBox(oRef)
    Dim oPartSel As Selection: Set oPartSel = oPart.Parent.Selection
    Dim oTmpSet As HybridBody
    Dim d1 As Double, d2 As Double, d3 As Double

    If oBox Is Nothing Then GoTo Cleanup

    oBox.Type = 1

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
            Call LBSortThree(d1, d2, d3, dDims)
            LBMeasureTarget = True
        End If
    End If

    ' restore the in-work object BEFORE deleting - otherwise CATIA shows a
    ' modal error dialog that blocks the macro
    Err.Clear
    oPart.InWorkObject = oPart.MainBody

    Err.Clear
    oPartSel.Clear
    If Not oTmpSet Is Nothing Then
        oPartSel.Add oTmpSet
    Else
        oPartSel.Add oBox
    End If
    oPartSel.Delete
    If Err.Number <> 0 Then
        Err.Clear
        oPartSel.Clear
        If Not oTmpSet Is Nothing Then oPartSel.Add oTmpSet Else oPartSel.Add oBox
        oPartSel.VisProperties.SetShow 1
    End If
    oPartSel.Clear

Cleanup:
    Set oPartSel = Nothing
    Set oTmpSet = Nothing
    Set oBox = Nothing
    Set oRef = Nothing
    Set oHSF = Nothing
    Err.Clear
End Function


' ==============================================================================
' HELPER: AXIS-ALIGNED DIMENSIONS VIA EXTREMUM POINTS
' (for releases without AddNewBoundingBox)
' ==============================================================================
Private Function LBMeasureAABBExtremum(oPart As Part, oPartDoc As PartDocument, _
                                       oTargetObj As Object, ByRef dDims() As Double) As Boolean
    On Error Resume Next
    Err.Clear
    LBMeasureAABBExtremum = False

    Dim oHSF As Object: Set oHSF = oPart.HybridShapeFactory
    Dim oSPA As Object: Set oSPA = oPartDoc.GetWorkbench("SPAWorkbench")
    Dim oRef As Reference: Set oRef = oPart.CreateReferenceFromObject(oTargetObj)
    If oHSF Is Nothing Or oSPA Is Nothing Or oRef Is Nothing Then GoTo CleanupE

    Dim oTmpSet As HybridBody
    Err.Clear
    Set oTmpSet = oPart.HybridBodies.Add()
    If oTmpSet Is Nothing Then GoTo CleanupE
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

        If LBGetExtremumCoord(oPart, oSPA, oTmpSet, oHSF, oRef, oDir, 1, k, vMax) And _
           LBGetExtremumCoord(oPart, oSPA, oTmpSet, oHSF, oRef, oDir, 0, k, vMin) Then
            spans(k) = vMax - vMin
        Else
            okAll = False
        End If

        Set oDir = Nothing
        If okAll = False Then Exit For
    Next k

    Err.Clear
    oPart.InWorkObject = oPart.MainBody

    Dim oPartSel As Selection
    Err.Clear
    Set oPartSel = oPartDoc.Selection
    oPartSel.Clear
    oPartSel.Add oTmpSet
    oPartSel.Delete
    If Err.Number <> 0 Then
        Err.Clear
        oPartSel.Clear
        oPartSel.Add oTmpSet
        oPartSel.VisProperties.SetShow 1
    End If
    oPartSel.Clear
    Set oPartSel = Nothing

    If okAll Then
        If (spans(0) + spans(1) + spans(2)) > 0.1 Then
            Call LBSortThree(spans(0), spans(1), spans(2), dDims)
            LBMeasureAABBExtremum = True
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
' HELPER: ONE EXTREMUM POINT, ONE AXIS
' minMax: 1 = maximum, 0 = minimum ; axisIndex: 0 = X, 1 = Y, 2 = Z
' ==============================================================================
Private Function LBGetExtremumCoord(oPart As Part, oSPA As Object, oTmpSet As HybridBody, _
                                    oHSF As Object, oRef As Reference, oDir As Object, _
                                    ByVal minMax As Long, ByVal axisIndex As Integer, _
                                    ByRef outCoord As Double) As Boolean
    On Error Resume Next
    LBGetExtremumCoord = False

    Dim oExt As Object
    Err.Clear
    Set oExt = oHSF.AddNewExtremum(oRef, oDir, minMax)
    If oExt Is Nothing Then Err.Clear: Exit Function

    oTmpSet.AppendHybridShape oExt

    Err.Clear
    oPart.UpdateObject oExt
    If Err.Number <> 0 Then Err.Clear: Exit Function

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
    LBGetExtremumCoord = True

    Set oMeas = Nothing
    Set oExtRef = Nothing
    Set oExt = Nothing
    Err.Clear
End Function


' ==============================================================================
' HELPER: DIMENSIONS FROM THE BODY INERTIA (last fallback)
' ==============================================================================
Private Sub LBGetBodyInertiaDims(oPartDoc As PartDocument, oBody As Object, ByRef dDims() As Double)
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
    Call LBSortThree(Sqr(A) * 1000, Sqr(B) * 1000, Sqr(C) * 1000, dDims)
    Set oInertia = Nothing
    Err.Clear
End Sub


' ==============================================================================
' HELPER: SORT THREE VALUES (largest to smallest)
' ==============================================================================
Private Sub LBSortThree(A As Double, B As Double, C As Double, ByRef arr() As Double)
    Dim temp As Double
    Dim vals(2) As Double
    vals(0) = A: vals(1) = B: vals(2) = C
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
