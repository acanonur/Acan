Option Explicit

' ==============================================================================
' MACRO: MASTER BOM (NO WINDOW OPENING + PER-PART THUMBNAILS)
' FIXED: Properly hides all parts, shows only current part for screenshot
' Hides compass and spec tree during screenshot
' ==============================================================================

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub GenerateMasterBOM()

    ' --- 1. SETUP & VARIABLES ---
    Dim catDoc As Document
    Dim rootProd As Product
    Dim sel As Selection

    ' Dictionaries
    Dim dictQty As Object, dictRef As Object, dictDesc As Object, dictProps As Object

    ' Excel
    Dim xlApp As Object, xlBook As Object, xlSheet As Object

    ' Loop Vars
    Dim uniqueKey As Variant, r As Integer, strPartNum As String
    Dim oPartProd As Product
    Dim propsArray As Variant

    ' Data Variables
    Dim dims(2) As Double
    Dim strMatName As String, dDensity As Double

    ' Screenshot
    Dim tempPicPath As String, fso As Object

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

    ' --- 3. TRAVERSE TREE (Count Parts & Calc Mass) ---
    Call TraverseTree(rootProd, dictQty, dictRef, dictDesc, dictProps)

    If dictQty.Count = 0 Then MsgBox "No parts found.", vbInformation: Exit Sub

    ' --- 4. EXCEL SETUP ---
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0

    xlApp.Visible = True
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)

    With xlSheet
        .Cells(1, 1).Value = "Thumbnail"
        .Cells(1, 2).Value = "Part Number"
        .Cells(1, 3).Value = "Description"
        .Cells(1, 4).Value = "Qty"
        .Cells(1, 5).Value = "Mass (kg)"
        .Cells(1, 6).Value = "Volume (m3)"
        .Cells(1, 7).Value = "Area (m2)"
        .Cells(1, 8).Value = "Length (mm)"
        .Cells(1, 9).Value = "Width (mm)"
        .Cells(1, 10).Value = "Height (mm)"
        .Cells(1, 11).Value = "Material"
        .Cells(1, 12).Value = "Density (kg/m3)"

        .Range("A1:L1").Font.Bold = True
        .Range("A1:L1").Interior.Color = RGB(220, 220, 220)
        .Columns("A:A").ColumnWidth = 15
        .Columns("B:C").ColumnWidth = 20
        .Columns("E:G").NumberFormat = "0.000000000"
        .Columns("L:L").NumberFormat = "0.000"
    End With
    r = 2

    ' --- 5. PROCESS UNIQUE PARTS ---
    For Each uniqueKey In dictQty.Keys

        strPartNum = uniqueKey

        ' Write Basic Data & Pre-Calculated Mass
        xlSheet.Cells(r, 2).Value = strPartNum
        xlSheet.Cells(r, 3).Value = dictDesc(uniqueKey)
        xlSheet.Cells(r, 4).Value = dictQty(uniqueKey)

        propsArray = dictProps(uniqueKey)
        xlSheet.Cells(r, 5).Value = propsArray(0)
        xlSheet.Cells(r, 6).Value = propsArray(1)
        xlSheet.Cells(r, 7).Value = propsArray(2)

        Set oPartProd = dictRef(uniqueKey)

        ' --- THUMBNAIL FROM ASSEMBLY VIEW (HIDE ALL, SHOW ONLY TARGET) ---
        If Not oPartProd Is Nothing Then
            Call CaptureThumbnailFromAssembly(oPartProd, rootProd, xlSheet, r, tempPicPath, fso, catDoc)
        End If

        ' Try to read part data without opening any window
        Dim oPart As Part
        Dim oPartDoc As PartDocument
        Dim hasPart As Boolean
        hasPart = TryGetLoadedPart(oPartProd, oPart, oPartDoc)

        ' --- MATERIAL & DENSITY ---
        strMatName = "N/A"
        dDensity = 0
        If hasPart Then
            Call GetMaterialAndDensity(oPart, strMatName, dDensity)
        End If

        xlSheet.Cells(r, 11).Value = strMatName
        If dDensity > 0 Then
            xlSheet.Cells(r, 12).Value = dDensity
        Else
            xlSheet.Cells(r, 12).Value = "N/A"
        End If

        ' --- DIMENSIONS ---
        dims(0) = 0: dims(1) = 0: dims(2) = 0
        If hasPart Then
            Call GetBoundingBoxDims(oPart, dims)
        End If

        If dims(0) < 0.1 Then
            Call GetInertiaDims(oPartProd.ReferenceProduct, dims)
        End If

        If dims(0) > 0 Then xlSheet.Cells(r, 8).Value = Format(dims(0), "0.0") Else xlSheet.Cells(r, 8).Value = "N/A"
        If dims(1) > 0 Then xlSheet.Cells(r, 9).Value = Format(dims(1), "0.0") Else xlSheet.Cells(r, 9).Value = "N/A"
        If dims(2) > 0 Then xlSheet.Cells(r, 10).Value = Format(dims(2), "0.0") Else xlSheet.Cells(r, 10).Value = "N/A"

        r = r + 1
        DoEvents
    Next uniqueKey

    xlSheet.Columns("A:L").AutoFit
    MsgBox "BOM Exported Successfully!", vbInformation

End Sub

' ==============================================================================
' HELPER: THUMBNAIL FROM ASSEMBLY VIEW (HIDE ALL, SHOW ONLY TARGET)
' ==============================================================================
Sub CaptureThumbnailFromAssembly(oProd As Product, rootProd As Product, xlSheet As Object, ByVal row As Long, _
                                 tempPicPath As String, fso As Object, catDoc As Document)
    On Error Resume Next
    If oProd Is Nothing Or rootProd Is Nothing Then Exit Sub

    Dim sel As Selection
    Dim oViewer As Viewer
    Dim shp As Object
    Dim i As Integer
    Dim childProd As Product

    Set sel = catDoc.Selection

    ' 1) Hide compass and specification tree for clean screenshot
    On Error Resume Next
    CATIA.StartCommand "Specification Tree Display"  ' Toggle off
    CATIA.StartCommand "Compass"  ' Toggle off
    Sleep 100
    On Error GoTo 0

    ' 2) HIDE ALL PARTS - Loop through all children and hide each one individually
    Call HideAllProductsRecursive(rootProd, sel)
    Sleep 100

    ' 3) Force target part to load
    On Error Resume Next
    oProd.ApplyWorkMode 2
    If Not oProd.ReferenceProduct Is Nothing Then
        oProd.ReferenceProduct.ApplyWorkMode 2
    End If
    On Error GoTo 0

    ' 4) SHOW ONLY THE TARGET PART (and its parent chain for visibility)
    Call ShowSingleProduct(oProd, sel)
    Sleep 200

    ' 5) Reframe on the visible part
    sel.Clear
    sel.Add oProd

    Set oViewer = CATIA.ActiveWindow.ActiveViewer
    If Not oViewer Is Nothing Then
        oViewer.Reframe
        Sleep 300

        ' 6) Capture screenshot
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
    Else
        xlSheet.Cells(row, 1).Value = "No Viewer"
    End If

    ' 7) SHOW ALL PARTS AGAIN (restore visibility)
    Call ShowAllProductsRecursive(rootProd, sel)
    Sleep 100

    ' 8) Restore compass and specification tree
    On Error Resume Next
    CATIA.StartCommand "Specification Tree Display"  ' Toggle on
    CATIA.StartCommand "Compass"  ' Toggle on
    On Error GoTo 0

    sel.Clear
    Err.Clear
End Sub

' ==============================================================================
' Hide all products recursively (each one individually)
' In CATIA: SetShow(1) = catVisPropertyNoShowAttr = HIDE
' ==============================================================================
Sub HideAllProductsRecursive(oProd As Product, sel As Selection)
    On Error Resume Next
    Dim i As Integer
    Dim childProd As Product

    If oProd Is Nothing Then Exit Sub

    ' Hide children first (bottom-up approach)
    If oProd.Products.Count > 0 Then
        For i = 1 To oProd.Products.Count
            Set childProd = oProd.Products.Item(i)

            ' Recursively hide children of this child
            Call HideAllProductsRecursive(childProd, sel)

            ' Hide this child
            sel.Clear
            sel.Add childProd
            sel.VisProperties.SetShow 1  ' 1 = HIDE (catVisPropertyNoShowAttr)
        Next i
    End If

    Err.Clear
End Sub

' ==============================================================================
' Show all products recursively (restore visibility)
' In CATIA: SetShow(0) = catVisPropertyShowAttr = SHOW
' ==============================================================================
Sub ShowAllProductsRecursive(oProd As Product, sel As Selection)
    On Error Resume Next
    Dim i As Integer
    Dim childProd As Product

    If oProd Is Nothing Then Exit Sub

    ' Show this level's children
    If oProd.Products.Count > 0 Then
        For i = 1 To oProd.Products.Count
            Set childProd = oProd.Products.Item(i)

            ' Show this child
            sel.Clear
            sel.Add childProd
            sel.VisProperties.SetShow 0  ' 0 = SHOW (catVisPropertyShowAttr)

            ' Recursively show children of this child
            Call ShowAllProductsRecursive(childProd, sel)
        Next i
    End If

    Err.Clear
End Sub

' ==============================================================================
' Show only a single product (and its parent chain)
' In CATIA: SetShow(0) = catVisPropertyShowAttr = SHOW
' ==============================================================================
Sub ShowSingleProduct(oProd As Product, sel As Selection)
    On Error Resume Next

    If oProd Is Nothing Then Exit Sub

    ' Show the target product
    sel.Clear
    sel.Add oProd
    sel.VisProperties.SetShow 0  ' 0 = SHOW (catVisPropertyShowAttr)

    ' Also show parent chain so the part is visible in context
    Dim parentProd As Product
    Set parentProd = GetParentProduct(oProd)

    Do While Not parentProd Is Nothing
        sel.Clear
        sel.Add parentProd
        sel.VisProperties.SetShow 0  ' 0 = SHOW
        Set parentProd = GetParentProduct(parentProd)
    Loop

    Err.Clear
End Sub

' ==============================================================================
' Returns parent Product (or Nothing if top-level)
' ==============================================================================
Private Function GetParentProduct(p As Product) As Product
    On Error Resume Next
    Set GetParentProduct = Nothing

    If p Is Nothing Then Exit Function

    If TypeName(p.Parent) = "Products" Then
        Set GetParentProduct = p.Parent.Parent
    End If

    On Error GoTo 0
End Function

' ==============================================================================
' HELPER: TRAVERSE TREE (CALCULATES MASS)
' ==============================================================================
Sub TraverseTree(oProd As Product, dQty As Object, dRef As Object, dDesc As Object, dProps As Object)
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
                Call TraverseTree(childProd, dQty, dRef, dDesc, dProps)
            Else
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
' HELPER: BOUNDING BOX
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
