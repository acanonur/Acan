Option Explicit

' ==============================================================================
' MACRO: MASTER BOM (NO WINDOW VERSION)
' This version does NOT open parts in new windows - avoids save dialog issue
' All data is extracted from the assembly context directly
' ==============================================================================

' Win32 API Sleep for stability
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
    Dim strMatName As String, dDensity As Double
    Dim dLength As Double, dWidth As Double, dHeight As Double

    Dim partCounter As Integer

    ' --- 2. INITIALIZATION ---
    On Error Resume Next
    Set catDoc = CATIA.ActiveDocument
    If Err.Number <> 0 Then MsgBox "No Document.", vbCritical: Exit Sub
    On Error GoTo 0

    If InStr(catDoc.Name, ".CATProduct") = 0 Then MsgBox "Open Assembly.", vbExclamation: Exit Sub

    Set rootProd = catDoc.Product
    Set sel = catDoc.Selection

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
        .Cells(1, 1).Value = "Part Number"
        .Cells(1, 2).Value = "Description"
        .Cells(1, 3).Value = "Qty"
        .Cells(1, 4).Value = "Mass (kg)"
        .Cells(1, 5).Value = "Volume (m3)"
        .Cells(1, 6).Value = "Area (m2)"
        .Cells(1, 7).Value = "Length (mm)"
        .Cells(1, 8).Value = "Width (mm)"
        .Cells(1, 9).Value = "Height (mm)"
        .Cells(1, 10).Value = "Material"
        .Cells(1, 11).Value = "Density (kg/m3)"

        .Range("A1:K1").Font.Bold = True
        .Range("A1:K1").Interior.Color = RGB(220, 220, 220)
        .Columns("A:B").ColumnWidth = 25
        .Columns("D:F").NumberFormat = "0.000000000"
        .Columns("K:K").NumberFormat = "0.000"
    End With
    r = 2
    partCounter = 0

    ' --- 5. PROCESS UNIQUE PARTS (NO WINDOW OPENING) ---
    For Each uniqueKey In dictQty.Keys

        strPartNum = uniqueKey

        ' Write Basic Data
        xlSheet.Cells(r, 1).Value = strPartNum
        xlSheet.Cells(r, 2).Value = dictDesc(uniqueKey)
        xlSheet.Cells(r, 3).Value = dictQty(uniqueKey)

        ' Retrieve Mass/Vol/Area from Dictionary (already calculated)
        propsArray = dictProps(uniqueKey)
        xlSheet.Cells(r, 4).Value = propsArray(0) ' Mass
        xlSheet.Cells(r, 5).Value = propsArray(1) ' Volume
        xlSheet.Cells(r, 6).Value = propsArray(2) ' Area

        ' Get Product reference
        Set oPartProd = dictRef(uniqueKey)

        If Not oPartProd Is Nothing Then
            On Error Resume Next

            ' --- GET DIMENSIONS FROM INERTIA (No window needed) ---
            dLength = 0: dWidth = 0: dHeight = 0
            Call GetDimensionsFromInertia(oPartProd, dLength, dWidth, dHeight)

            xlSheet.Cells(r, 7).Value = Format(dLength, "0.0")
            xlSheet.Cells(r, 8).Value = Format(dWidth, "0.0")
            xlSheet.Cells(r, 9).Value = Format(dHeight, "0.0")

            ' --- GET MATERIAL (No window needed) ---
            strMatName = "No Material"
            dDensity = 0
            Call GetMaterialFromProduct(oPartProd, strMatName, dDensity)

            xlSheet.Cells(r, 10).Value = strMatName
            If dDensity > 0 Then
                xlSheet.Cells(r, 11).Value = dDensity
            Else
                xlSheet.Cells(r, 11).Value = "N/A"
            End If

            On Error GoTo 0
        End If

        partCounter = partCounter + 1

        ' Progress update
        If partCounter Mod 10 = 0 Then
            xlApp.StatusBar = "Processing: " & partCounter & " of " & dictQty.Count & " parts..."
            DoEvents
        End If

        r = r + 1

    Next uniqueKey

    ' Final cleanup
    xlApp.StatusBar = False
    xlSheet.Columns("A:K").AutoFit

    MsgBox "BOM Exported Successfully!" & vbCrLf & _
           "Total unique parts: " & dictQty.Count, vbInformation

End Sub

' ==============================================================================
' HELPER: TRAVERSE TREE (CALCULATES MASS)
' ==============================================================================
Sub TraverseTree(oProd As Product, dQty As Object, dRef As Object, dDesc As Object, dProps As Object)
    Dim childProd As Product, i As Integer, partNum As String
    Dim dMass As Double, dVol As Double, dArea As Double
    Dim oAnalyze As Object

    On Error Resume Next

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

    On Error GoTo 0
End Sub

' ==============================================================================
' HELPER: GET DIMENSIONS FROM INERTIA (No window needed)
' Calculates approximate dimensions from principal moments of inertia
' ==============================================================================
Sub GetDimensionsFromInertia(oProd As Product, ByRef dLength As Double, ByRef dWidth As Double, ByRef dHeight As Double)
    On Error Resume Next

    Dim oInertia As Object
    Set oInertia = oProd.ReferenceProduct.GetTechnologicalObject("Inertia")

    If oInertia Is Nothing Then
        Set oInertia = oProd.GetTechnologicalObject("Inertia")
    End If

    If oInertia Is Nothing Then Exit Sub

    Dim dMass As Double
    dMass = oInertia.Mass
    If dMass <= 0 Then
        Set oInertia = Nothing
        Exit Sub
    End If

    Dim Matrix(8)
    oInertia.GetPrincipalMoments Matrix

    ' Calculate approximate dimensions from moments of inertia
    ' For a rectangular box: Ixx = m(b²+c²)/12, etc.
    Dim A As Double, B As Double, C As Double
    A = (6 * (Matrix(1) + Matrix(2) - Matrix(0)) / dMass)
    B = (6 * (Matrix(0) + Matrix(2) - Matrix(1)) / dMass)
    C = (6 * (Matrix(0) + Matrix(1) - Matrix(2)) / dMass)

    If A < 0 Then A = 0
    If B < 0 Then B = 0
    If C < 0 Then C = 0

    ' Convert to mm and sort (largest first)
    Dim d1 As Double, d2 As Double, d3 As Double
    d1 = Sqr(A) * 1000
    d2 = Sqr(B) * 1000
    d3 = Sqr(C) * 1000

    ' Sort descending
    Call SortThreeDims(d1, d2, d3, dLength, dWidth, dHeight)

    Set oInertia = Nothing
    Err.Clear
End Sub

' ==============================================================================
' HELPER: GET MATERIAL FROM PRODUCT (No window needed)
' Accesses the part document directly from the product reference
' ==============================================================================
Sub GetMaterialFromProduct(oProd As Product, ByRef matName As String, ByRef density As Double)
    On Error Resume Next

    Dim oPartDoc As Document
    Dim oPart As Part

    ' Get the part document from the product
    Set oPartDoc = oProd.ReferenceProduct.Parent
    If oPartDoc Is Nothing Then Exit Sub

    Set oPart = oPartDoc.Part
    If oPart Is Nothing Then Exit Sub

    ' Try to get material using CATMatManagerVBExt
    Dim oManager As Object
    Set oManager = oPart.GetItem("CATMatManagerVBExt")

    If Not oManager Is Nothing Then
        Dim oMat As Object
        Dim oBody As Body

        Set oBody = oPart.MainBody
        If Not oBody Is Nothing Then
            oManager.GetMaterialOnBody oBody, oMat
        End If

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

    ' Fallback: try to get density from part inertia
    If density <= 0 Then
        Dim oInertia As Object
        Set oInertia = oPart.Inertia
        If Not oInertia Is Nothing Then
            density = oInertia.Density
            Set oInertia = Nothing
        End If
    End If

    Set oPart = Nothing
    Set oPartDoc = Nothing
    Err.Clear
End Sub

' ==============================================================================
' HELPER: SORT THREE DIMENSIONS (Descending)
' ==============================================================================
Sub SortThreeDims(d1 As Double, d2 As Double, d3 As Double, ByRef outL As Double, ByRef outW As Double, ByRef outH As Double)
    Dim temp As Double
    Dim vals(2) As Double
    vals(0) = d1: vals(1) = d2: vals(2) = d3

    Dim i As Integer, j As Integer
    For i = 0 To 1
        For j = i + 1 To 2
            If vals(i) < vals(j) Then
                temp = vals(i)
                vals(i) = vals(j)
                vals(j) = temp
            End If
        Next j
    Next i

    outL = vals(0)
    outW = vals(1)
    outH = vals(2)
End Sub
