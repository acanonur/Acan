Option Explicit

' ==============================================================================
' MACRO: MASTER BOM - PRODUCT LEVEL ONLY - TWO PHASE (FAST)
'
' Phase 1: traverses the assembly and writes the COMPLETE Excel table
'          (level, part number, description, qty, mass, volume, area,
'          dimensions, material, density) - NO screenshots, no visibility
'          changes, so it runs fast and CATIA stays usable.
'
' Phase 2: one dedicated screenshot pass at the end:
'          - compass and spec tree toggled off ONCE
'          - all leaf parts hidden ONCE with a single batched SetShow call
'          - per part: show -> reframe -> capture -> hide (2 cheap calls)
'          - everything restored ONCE at the end
'
' Measurement code is identical to the original working version
' (CATIA_BOM_Extractor_Final). No PartBody handling in this macro.
' ==============================================================================

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub GenerateMasterBOM()

    ' --- 1. SETUP & VARIABLES ---
    Dim catDoc As Document
    Dim rootProd As Product
    Dim sel As Selection

    ' Dictionaries
    Dim dictQty As Object, dictRef As Object, dictDesc As Object, dictProps As Object
    Dim dictLevel As Object, dictRow As Object

    ' Every leaf instance in the tree (to hide them all in one call)
    Dim colLeaves As Collection

    ' Excel
    Dim xlApp As Object, xlBook As Object, xlSheet As Object

    ' Loop Vars
    Dim uniqueKey As Variant, r As Long, strPartNum As String
    Dim oPartProd As Product
    Dim propsArray As Variant
    Dim nDone As Long, nTotal As Long

    ' Part data
    Dim oPart As Part
    Dim oPartDoc As PartDocument
    Dim hasPart As Boolean
    Dim dims(2) As Double
    Dim strMatName As String, dDensity As Double

    ' Screenshot
    Dim tempPicPath As String, fso As Object
    Dim oViewer As Viewer

    ' Timing
    Dim startTime As Double
    startTime = Timer

    ' --- 2. INITIALIZATION ---
    On Error Resume Next
    Set catDoc = CATIA.ActiveDocument
    If Err.Number <> 0 Then MsgBox "No Document.", vbCritical: Exit Sub
    On Error GoTo 0

    If InStr(catDoc.Name, ".CATProduct") = 0 Then MsgBox "Open Assembly.", vbExclamation: Exit Sub

    Set rootProd = catDoc.Product
    Set sel = catDoc.Selection

    Set fso = CreateObject("Scripting.FileSystemObject")
    tempPicPath = "C:\Temp\catia_bom_shot.jpg"
    If Not fso.FolderExists("C:\Temp") Then fso.CreateFolder ("C:\Temp")

    Set dictQty = CreateObject("Scripting.Dictionary")
    Set dictRef = CreateObject("Scripting.Dictionary")
    Set dictDesc = CreateObject("Scripting.Dictionary")
    Set dictProps = CreateObject("Scripting.Dictionary")
    Set dictLevel = CreateObject("Scripting.Dictionary")
    Set dictRow = CreateObject("Scripting.Dictionary")
    Set colLeaves = New Collection

    ' --- 3. TRAVERSE TREE (Count Parts & Calc Mass) ---
    Call TraverseTree(rootProd, dictQty, dictRef, dictDesc, dictProps, dictLevel, colLeaves, 1)

    If dictQty.Count = 0 Then MsgBox "No parts found.", vbInformation: Exit Sub

    ' --- 4. EXCEL SETUP ---
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then Set xlApp = CreateObject("Excel.Application")
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

        .Range("A1:M1").Font.Bold = True
        .Range("A1:M1").Interior.Color = RGB(220, 220, 220)
        .Columns("A:A").ColumnWidth = 15
        .Columns("B:B").ColumnWidth = 8
        .Columns("C:D").ColumnWidth = 20
        .Columns("F:H").NumberFormat = "0.000000000"
        .Columns("M:M").NumberFormat = "0.000"
    End With

    nTotal = dictQty.Count

    ' ==========================================================================
    ' PHASE 1: WRITE ALL DATA (NO SCREENSHOTS)
    ' ==========================================================================
    r = 2
    nDone = 0
    For Each uniqueKey In dictQty.Keys

        strPartNum = uniqueKey
        Set oPartProd = dictRef(uniqueKey)
        dictRow.Add strPartNum, r

        xlSheet.Cells(r, 2).Value = dictLevel(uniqueKey)
        xlSheet.Cells(r, 3).Value = strPartNum
        xlSheet.Cells(r, 4).Value = dictDesc(uniqueKey)
        xlSheet.Cells(r, 5).Value = dictQty(uniqueKey)

        propsArray = dictProps(uniqueKey)
        xlSheet.Cells(r, 6).Value = propsArray(0)
        xlSheet.Cells(r, 7).Value = propsArray(1)
        xlSheet.Cells(r, 8).Value = propsArray(2)

        ' Try to read part data without opening any window
        hasPart = TryGetLoadedPart(oPartProd, oPart, oPartDoc)

        ' --- MATERIAL & DENSITY ---
        strMatName = "N/A"
        dDensity = 0
        If hasPart Then
            Call GetMaterialAndDensity(oPart, strMatName, dDensity)
        End If

        xlSheet.Cells(r, 12).Value = strMatName
        If dDensity > 0 Then
            xlSheet.Cells(r, 13).Value = dDensity
        Else
            xlSheet.Cells(r, 13).Value = "N/A"
        End If

        ' --- DIMENSIONS ---
        dims(0) = 0: dims(1) = 0: dims(2) = 0
        If hasPart Then
            Call GetBoundingBoxDims(oPart, dims)
        End If

        If dims(0) < 0.1 Then
            Call GetInertiaDims(oPartProd.ReferenceProduct, dims)
        End If

        If dims(0) > 0 Then xlSheet.Cells(r, 9).Value = Format(dims(0), "0.0") Else xlSheet.Cells(r, 9).Value = "N/A"
        If dims(1) > 0 Then xlSheet.Cells(r, 10).Value = Format(dims(1), "0.0") Else xlSheet.Cells(r, 10).Value = "N/A"
        If dims(2) > 0 Then xlSheet.Cells(r, 11).Value = Format(dims(2), "0.0") Else xlSheet.Cells(r, 11).Value = "N/A"

        r = r + 1
        nDone = nDone + 1
        xlApp.StatusBar = "BOM data: " & nDone & " / " & nTotal
        DoEvents
    Next uniqueKey

    ' ==========================================================================
    ' PHASE 2: SCREENSHOT PASS
    ' ==========================================================================
    Set oViewer = Nothing
    On Error Resume Next
    Set oViewer = CATIA.ActiveWindow.ActiveViewer
    CATIA.StartCommand "Specification Tree Display"  ' Toggle off
    CATIA.StartCommand "Compass"                     ' Toggle off
    Sleep 100
    On Error GoTo 0

    ' Hide ALL leaf parts with ONE batched SetShow call.
    ' Parent nodes stay untouched, so showing one leaf later is enough.
    Call SetShowCollection(sel, colLeaves, 1)
    Sleep 100

    nDone = 0
    For Each uniqueKey In dictQty.Keys
        strPartNum = uniqueKey
        Set oPartProd = dictRef(uniqueKey)
        r = dictRow(strPartNum)

        If Not oPartProd Is Nothing Then
            Call SetShowSingle(sel, oPartProd, 0)   ' 0 = SHOW
            Sleep 50
            Call CaptureViewToExcel(oViewer, xlSheet, r, tempPicPath, fso)
            Call SetShowSingle(sel, oPartProd, 1)   ' 1 = HIDE
        Else
            xlSheet.Cells(r, 1).Value = "No Preview"
        End If

        nDone = nDone + 1
        xlApp.StatusBar = "Screenshots: " & nDone & " / " & nTotal
        DoEvents
    Next uniqueKey

    ' --- RESTORE SCENE (done ONCE) ---
    Call SetShowCollection(sel, colLeaves, 0)   ' show all leaf parts again

    On Error Resume Next
    CATIA.StartCommand "Specification Tree Display"  ' Toggle on
    CATIA.StartCommand "Compass"                     ' Toggle on
    If Not oViewer Is Nothing Then oViewer.Reframe
    On Error GoTo 0
    sel.Clear

    ' --- FINALIZE EXCEL ---
    xlApp.StatusBar = False
    xlApp.ScreenUpdating = True
    xlSheet.Columns("B:M").AutoFit
    xlSheet.Columns("A:A").ColumnWidth = 15

    MsgBox "BOM Exported Successfully!" & vbCrLf & _
           "Parts: " & nTotal & vbCrLf & _
           "Time elapsed: " & Format(Timer - startTime, "0.0") & " seconds", vbInformation

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
    Sleep 50

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
' HELPER: TRAVERSE TREE (CALCULATES MASS + TRACKS LEVEL)
' Level 1 = direct children of root, Level 2 = grandchildren, etc.
' Also fills colLeaves with EVERY leaf instance (for the batched hide)
' ==============================================================================
Sub TraverseTree(oProd As Product, dQty As Object, dRef As Object, dDesc As Object, _
                 dProps As Object, dLevel As Object, colLeaves As Collection, _
                 ByVal currentLevel As Integer)
    Dim childProd As Product, i As Integer, partNum As String
    Dim dMass As Double, dVol As Double, dArea As Double
    Dim oAnalyze As Object

    If oProd.Products.Count > 0 Then
        For i = 1 To oProd.Products.Count
            Set childProd = oProd.Products.Item(i)
            On Error Resume Next
            childProd.ApplyWorkMode 2
            On Error GoTo 0

            If childProd.Products.Count > 0 Then
                Call TraverseTree(childProd, dQty, dRef, dDesc, dProps, dLevel, _
                                  colLeaves, currentLevel + 1)
            Else
                ' Remember every leaf instance so all can be hidden in one call
                colLeaves.Add childProd

                partNum = ""
                On Error Resume Next
                partNum = childProd.PartNumber
                On Error GoTo 0

                If partNum <> "" Then
                    If dQty.Exists(partNum) Then
                        dQty(partNum) = dQty(partNum) + 1
                    Else
                        dQty.Add partNum, 1
                        dRef.Add partNum, childProd
                        dLevel.Add partNum, currentLevel

                        On Error Resume Next
                        dDesc.Add partNum, childProd.DescriptionRef
                        If Err.Number <> 0 Then
                            dDesc(partNum) = ""
                            Err.Clear
                        End If
                        On Error GoTo 0

                        dMass = 0: dVol = 0: dArea = 0
                        On Error Resume Next

                        Set oAnalyze = childProd.Analyze
                        If Not oAnalyze Is Nothing Then
                            dMass = oAnalyze.Mass
                            dVol = oAnalyze.Volume / (1000# ^ 3)
                            dArea = oAnalyze.WetArea / (1000# ^ 2)
                        End If

                        If dMass <= 0.0000001 Then
                            Dim oInertia As Object
                            Set oInertia = childProd.ReferenceProduct.GetTechnologicalObject("Inertia")
                            If Not oInertia Is Nothing Then
                                dMass = oInertia.Mass
                            End If
                            Set oInertia = Nothing
                        End If

                        Set oAnalyze = Nothing
                        On Error GoTo 0

                        dProps.Add partNum, Array(dMass, dVol, dArea)
                    End If
                End If
            End If

            Set childProd = Nothing
        Next i
    End If
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
' HELPER: GET MATERIAL & DENSITY (Using VBExt)
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
' HELPER: BOUNDING BOX (original working version)
' ==============================================================================
Sub GetBoundingBoxDims(oPart As Part, ByRef dDims() As Double)
    On Error Resume Next
    Dim foundValidBox As Boolean: foundValidBox = False

    If Not oPart.MainBody Is Nothing Then
        foundValidBox = MeasureTarget(oPart, oPart.MainBody, dDims)
    End If
    If foundValidBox = False Then
        Dim hb As HybridBody
        For Each hb In oPart.HybridBodies
            foundValidBox = MeasureTarget(oPart, hb, dDims)
            If foundValidBox = True Then Exit For
        Next
        Set hb = Nothing
    End If
    Err.Clear
End Sub

Function MeasureTarget(oPart As Part, oTargetObj As Object, ByRef dDims() As Double) As Boolean
    On Error Resume Next
    MeasureTarget = False
    Dim oHSF As Object: Set oHSF = oPart.HybridShapeFactory
    Dim oRef As Reference: Set oRef = oPart.CreateReferenceFromObject(oTargetObj)
    Dim oBox As Object: Set oBox = oHSF.AddNewBoundingBox(oRef)

    If oBox Is Nothing Then
        Set oRef = Nothing
        Set oHSF = Nothing
        Exit Function
    End If

    oBox.Type = 1
    oPart.UpdateObject oBox

    If Err.Number <> 0 Then
        Dim oSelFail As Selection: Set oSelFail = CATIA.ActiveDocument.Selection
        oSelFail.Clear: oSelFail.Add oBox: oSelFail.Delete
        Set oSelFail = Nothing
        Set oBox = Nothing
        Set oRef = Nothing
        Set oHSF = Nothing
        Err.Clear
        Exit Function
    End If

    Dim d1 As Double, d2 As Double, d3 As Double
    d1 = oBox.GetLength.Value: d2 = oBox.GetWidth.Value: d3 = oBox.GetHeight.Value

    Dim oSel As Selection: Set oSel = CATIA.ActiveDocument.Selection
    oSel.Clear: oSel.Add oBox: oSel.Delete
    Set oSel = Nothing

    If (d1 + d2 + d3) > 0.1 Then
        Call SortThree(d1, d2, d3, dDims)
        MeasureTarget = True
    End If

    Set oBox = Nothing
    Set oRef = Nothing
    Set oHSF = Nothing
End Function

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
