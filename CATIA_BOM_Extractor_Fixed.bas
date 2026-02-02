Option Explicit

' ==============================================================================
' MACRO: MASTER BOM (FIXED VERSION)
' FIX: Ensures all opened part windows are properly closed after processing
' CHANGES:
' 1. Restructured window handling with guaranteed cleanup
' 2. Added explicit window close after each part processing
' 3. Moved CloseAllExtraWindows outside of nested conditionals
' 4. Added error handling that ensures cleanup happens
' 5. Added periodic cleanup to prevent memory buildup
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
    Dim dims(2) As Double
    Dim strMatName As String, dDensity As Double
    Dim tempPicPath As String, fso As Object

    ' Window Handling
    Dim mainWinCaption As String
    Dim winCountBefore As Integer
    Dim tStart As Single
    Dim winOpened As Boolean
    Dim partCounter As Integer

    ' --- 2. INITIALIZATION ---
    On Error Resume Next
    Set catDoc = CATIA.ActiveDocument
    If Err.Number <> 0 Then MsgBox "No Document.", vbCritical: Exit Sub
    On Error GoTo 0

    If InStr(catDoc.Name, ".CATProduct") = 0 Then MsgBox "Open Assembly.", vbExclamation: Exit Sub

    ' Memorize Main Window
    mainWinCaption = CATIA.ActiveWindow.Caption

    Set rootProd = catDoc.Product
    Set sel = catDoc.Selection
    Set fso = CreateObject("Scripting.FileSystemObject")

    tempPicPath = "C:\Temp\catia_master_shot.jpg"
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
    partCounter = 0

    ' --- 5. PROCESS UNIQUE PARTS ---
    For Each uniqueKey In dictQty.Keys

        strPartNum = uniqueKey
        winOpened = False  ' Reset flag for each iteration

        ' Write Basic Data & Pre-Calculated Mass
        xlSheet.Cells(r, 2).Value = strPartNum
        xlSheet.Cells(r, 3).Value = dictDesc(uniqueKey)
        xlSheet.Cells(r, 4).Value = dictQty(uniqueKey)

        ' Retrieve Mass/Vol/Area from Dictionary (Calculated in Step 3)
        propsArray = dictProps(uniqueKey)
        xlSheet.Cells(r, 5).Value = propsArray(0) ' Mass
        xlSheet.Cells(r, 6).Value = propsArray(1) ' Vol
        xlSheet.Cells(r, 7).Value = propsArray(2) ' Area

        ' Open Window for Dimensions & Material
        Set oPartProd = dictRef(uniqueKey)

        If oPartProd Is Nothing Then
            ' No valid product reference - skip to next
            xlSheet.Cells(r, 1).Value = "No Ref"
            GoTo NextPart
        End If

        On Error Resume Next

        ' A. OPEN WINDOW
        winCountBefore = CATIA.Windows.Count
        sel.Clear
        sel.Add oPartProd
        CATIA.StartCommand "Open in New Window"

        ' Wait loop for window to open
        winOpened = False
        tStart = Timer
        Do While Timer < tStart + 10
            If CATIA.Windows.Count > winCountBefore Then
                winOpened = True
                Exit Do
            End If
            DoEvents
            Sleep 100  ' Small delay to reduce CPU usage
        Loop

        ' B. PROCESS IF OPENED
        If Not winOpened Then
            xlSheet.Cells(r, 1).Value = "Load Timeout"
            sel.Clear
            GoTo CleanupWindow
        End If

        Sleep 300

        Dim activeWin As Window
        Set activeWin = CATIA.ActiveWindow

        If activeWin.Caption = mainWinCaption Then
            ' Failed to switch to new window
            sel.Clear
            GoTo CleanupWindow
        End If

        ' Hide UI for cleaner screenshot
        On Error Resume Next
        CATIA.StartCommand "Compass Display"
        CATIA.StartCommand "Specifications"
        On Error GoTo 0

        Dim oPart As Part
        On Error Resume Next
        Set oPart = CATIA.ActiveDocument.Part

        If oPart Is Nothing Then
            xlSheet.Cells(r, 1).Value = "Not a Part"
            GoTo CleanupWindow
        End If
        On Error GoTo 0

        ' --- 1. GET MATERIAL & DENSITY ---
        strMatName = "No Material"
        dDensity = 0
        Call GetMaterialAndDensity(oPart, strMatName, dDensity)

        xlSheet.Cells(r, 11).Value = strMatName
        If dDensity > 0 Then
            xlSheet.Cells(r, 12).Value = dDensity
        Else
            xlSheet.Cells(r, 12).Value = "N/A"
        End If

        ' --- 2. GET DIMENSIONS ---
        dims(0) = 0: dims(1) = 0: dims(2) = 0
        Call GetBoundingBoxDims(oPart, dims)

        ' Fallback to Inertia Dims
        If dims(0) < 0.1 Then
            On Error Resume Next
            Call GetInertiaDims(CATIA.ActiveDocument.Product, dims)
            On Error GoTo 0
        End If

        xlSheet.Cells(r, 8).Value = Format(dims(0), "0.0")
        xlSheet.Cells(r, 9).Value = Format(dims(1), "0.0")
        xlSheet.Cells(r, 10).Value = Format(dims(2), "0.0")

        ' --- 3. SCREENSHOT ---
        On Error Resume Next
        Dim oViewer As Viewer
        Set oViewer = activeWin.ActiveViewer
        If Not oViewer Is Nothing Then
            oViewer.Reframe
            If fso.FileExists(tempPicPath) Then fso.DeleteFile tempPicPath
            oViewer.CaptureToFile 4, tempPicPath
            If fso.FileExists(tempPicPath) Then
                Dim shp As Object
                Set shp = xlSheet.Shapes.AddPicture(tempPicPath, False, True, _
                    xlSheet.Cells(r, 1).Left + 2, xlSheet.Cells(r, 1).Top + 2, -1, -1)
                If Not shp Is Nothing Then
                    shp.Height = 50
                    If shp.Width > 90 Then shp.Width = 90
                End If
                xlSheet.Rows(r).RowHeight = 60
            End If
        End If
        On Error GoTo 0

CleanupWindow:
        ' *** CRITICAL: ALWAYS CLOSE EXTRA WINDOWS ***
        ' This section runs regardless of success/failure above

        On Error Resume Next
        sel.Clear

        ' Close the specific window we opened (if it exists)
        If winOpened Then
            Call ForceClosePartWindow(mainWinCaption)
        End If

        ' Safety: Close any remaining extra windows
        Call CloseAllExtraWindows(mainWinCaption)

        ' Re-activate main document
        catDoc.Activate
        Sleep 200  ' Allow CATIA to stabilize

        On Error GoTo 0

        ' Increment counter and periodic cleanup
        partCounter = partCounter + 1

        ' Every 10 parts, do a thorough cleanup
        If partCounter Mod 10 = 0 Then
            Call PeriodicCleanup(mainWinCaption, catDoc)
            xlApp.StatusBar = "Processing: " & partCounter & " of " & dictQty.Count & " parts..."
        End If

NextPart:
        r = r + 1
        DoEvents  ' Allow system to process events

    Next uniqueKey

    ' Final cleanup
    xlApp.StatusBar = False
    xlSheet.Columns("A:L").AutoFit

    ' Delete temp file
    On Error Resume Next
    If fso.FileExists(tempPicPath) Then fso.DeleteFile tempPicPath
    On Error GoTo 0

    MsgBox "BOM Exported Successfully!" & vbCrLf & _
           "Total unique parts: " & dictQty.Count, vbInformation

End Sub

' ==============================================================================
' HELPER 1: TRAVERSE TREE (CALCULATES MASS HERE)
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

                        ' --- CALCULATE MASS (Product Context) ---
                        dMass = 0: dVol = 0: dArea = 0
                        On Error Resume Next

                        ' Method A: Analyze
                        Set oAnalyze = childProd.Analyze
                        If Not oAnalyze Is Nothing Then
                            dMass = oAnalyze.Mass
                            dVol = oAnalyze.Volume / (1000# ^ 3)
                            dArea = oAnalyze.WetArea / (1000# ^ 2)
                        End If

                        ' Method B: Inertia Fallback
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
' HELPER 2: GET MATERIAL & DENSITY (Using VBExt)
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
' HELPER 3: FORCE CLOSE PART WINDOW (New - More Aggressive)
' ==============================================================================
Sub ForceClosePartWindow(keepCaption As String)
    Dim i As Integer
    Dim win As Window
    Dim attempts As Integer

    On Error Resume Next

    ' Try multiple times to ensure window closes
    For attempts = 1 To 3
        For i = CATIA.Windows.Count To 1 Step -1
            Set win = CATIA.Windows.Item(i)
            If Not win Is Nothing Then
                If win.Caption <> keepCaption Then
                    win.Close
                    Sleep 100
                End If
            End If
            Set win = Nothing
        Next i

        ' Check if only main window remains
        If CATIA.Windows.Count <= 1 Then Exit For

        Sleep 200
    Next attempts

    Err.Clear
End Sub

' ==============================================================================
' HELPER 4: CLOSE ALL EXTRA WINDOWS (Safety Net)
' ==============================================================================
Sub CloseAllExtraWindows(keepCaption As String)
    Dim i As Integer
    Dim win As Window

    On Error Resume Next

    For i = CATIA.Windows.Count To 1 Step -1
        Set win = CATIA.Windows.Item(i)
        If Not win Is Nothing Then
            If win.Caption <> keepCaption Then
                win.Close
            End If
        End If
        Set win = Nothing
    Next i

    Err.Clear
End Sub

' ==============================================================================
' HELPER 5: PERIODIC CLEANUP (Memory Management)
' ==============================================================================
Sub PeriodicCleanup(mainCaption As String, mainDoc As Document)
    On Error Resume Next

    ' Close any stray windows
    Call CloseAllExtraWindows(mainCaption)

    ' Re-activate main document
    mainDoc.Activate

    ' Small delay for CATIA to stabilize
    Sleep 500

    ' Force garbage collection (VBA doesn't have explicit GC, but this helps)
    DoEvents

    Err.Clear
End Sub

' ==============================================================================
' HELPER 6: BOUNDING BOX
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
        Dim oSelFail As Selection
        Set oSelFail = CATIA.ActiveDocument.Selection
        oSelFail.Clear
        oSelFail.Add oBox
        oSelFail.Delete
        Set oSelFail = Nothing
        Set oBox = Nothing
        Set oRef = Nothing
        Set oHSF = Nothing
        Err.Clear
        Exit Function
    End If

    Dim d1 As Double, d2 As Double, d3 As Double
    d1 = oBox.GetLength.Value
    d2 = oBox.GetWidth.Value
    d3 = oBox.GetHeight.Value

    Dim oSel As Selection
    Set oSel = CATIA.ActiveDocument.Selection
    oSel.Clear
    oSel.Add oBox
    oSel.Delete
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

    Dim oInertia As Object
    Set oInertia = oProd.GetTechnologicalObject("Inertia")
    If oInertia Is Nothing Then Exit Sub

    Dim dMass As Double
    dMass = oInertia.Mass
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
                temp = vals(i)
                vals(i) = vals(j)
                vals(j) = temp
            End If
        Next j
    Next i

    arr(0) = vals(0)
    arr(1) = vals(1)
    arr(2) = vals(2)
End Sub
