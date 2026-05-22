Option Explicit

' ==============================================================================
' MACRO: PART BODY BOM EXTRACTOR (For CATPart with multiple Bodies)
' Extracts info from each Body in a CATPart, takes individual screenshots
' OPTIMIZED FOR SPEED
' ==============================================================================

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Global variables for speed optimization
Private g_ScreenUpdatesDisabled As Boolean
Private g_SpecTreeHidden As Boolean
Private g_CompassHidden As Boolean

Sub ExtractPartBodyBOM()

    ' --- VARIABLES ---
    Dim catDoc As Document
    Dim partDoc As PartDocument
    Dim oPart As Part
    Dim oBody As Body
    Dim sel As Selection
    Dim oViewer As Viewer

    ' Excel
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim r As Integer

    ' Data
    Dim bodyName As String
    Dim dMass As Double, dVol As Double, dArea As Double
    Dim dims(2) As Double
    Dim strMatName As String, dDensity As Double

    ' Screenshot
    Dim tempPicPath As String, fso As Object
    Dim shp As Object

    ' Timing
    Dim startTime As Double
    startTime = Timer

    ' --- INITIALIZATION ---
    On Error Resume Next
    Set catDoc = CATIA.ActiveDocument
    If Err.Number <> 0 Then
        MsgBox "No document open.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Check if it's a CATPart
    If InStr(catDoc.Name, ".CATPart") = 0 Then
        MsgBox "Please open a CATPart file.", vbExclamation
        Exit Sub
    End If

    Set partDoc = catDoc
    Set oPart = partDoc.Part
    Set sel = catDoc.Selection

    ' Check if there are bodies
    If oPart.Bodies.Count = 0 Then
        MsgBox "No bodies found in this part.", vbInformation
        Exit Sub
    End If

    ' Setup temp folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    tempPicPath = "C:\Temp\catia_body_shot.jpg"
    If Not fso.FolderExists("C:\Temp") Then fso.CreateFolder ("C:\Temp")

    ' --- EXCEL SETUP ---
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0

    xlApp.Visible = True
    xlApp.ScreenUpdating = False  ' Speed: disable Excel screen updates
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)

    With xlSheet
        .Cells(1, 1).Value = "Thumbnail"
        .Cells(1, 2).Value = "Body Name"
        .Cells(1, 3).Value = "Mass (kg)"
        .Cells(1, 4).Value = "Volume (mm3)"
        .Cells(1, 5).Value = "Area (mm2)"
        .Cells(1, 6).Value = "Length (mm)"
        .Cells(1, 7).Value = "Width (mm)"
        .Cells(1, 8).Value = "Height (mm)"
        .Cells(1, 9).Value = "Material"
        .Cells(1, 10).Value = "Density (kg/m3)"

        .Range("A1:J1").Font.Bold = True
        .Range("A1:J1").Interior.Color = RGB(220, 220, 220)
        .Columns("A:A").ColumnWidth = 15
        .Columns("B:B").ColumnWidth = 30
    End With
    r = 2

    ' --- PREPARE VIEW (once, not per body) ---
    Set oViewer = CATIA.ActiveWindow.ActiveViewer

    ' Hide spec tree and compass once at start
    On Error Resume Next
    CATIA.StartCommand "Compass"
    Sleep 50
    g_CompassHidden = True
    On Error GoTo 0

    ' --- HIDE ALL BODIES FIRST (faster than hiding/showing each iteration) ---
    Call HideAllBodies(oPart, sel)
    Sleep 50

    ' --- PROCESS EACH BODY ---
    Dim i As Integer
    Dim totalBodies As Integer
    totalBodies = oPart.Bodies.Count

    For i = 1 To totalBodies
        Set oBody = oPart.Bodies.Item(i)
        bodyName = oBody.Name

        ' Skip if body has no shapes (empty body)
        On Error Resume Next
        If oBody.Shapes.Count = 0 Then
            Err.Clear
            GoTo NextBody
        End If
        On Error GoTo 0

        ' --- SHOW ONLY THIS BODY ---
        sel.Clear
        sel.Add oBody
        sel.VisProperties.SetShow 0  ' 0 = SHOW

        ' Quick update
        oViewer.Reframe
        Sleep 100  ' Reduced sleep time

        ' --- TAKE SCREENSHOT ---
        On Error Resume Next
        If fso.FileExists(tempPicPath) Then fso.DeleteFile tempPicPath
        oViewer.CaptureToFile 4, tempPicPath

        If fso.FileExists(tempPicPath) Then
            Set shp = xlSheet.Shapes.AddPicture(tempPicPath, False, True, _
                xlSheet.Cells(r, 1).Left + 2, xlSheet.Cells(r, 1).Top + 2, -1, -1)
            If Not shp Is Nothing Then
                shp.Height = 50
                If shp.Width > 90 Then shp.Width = 90
            End If
            xlSheet.Rows(r).RowHeight = 60
        Else
            xlSheet.Cells(r, 1).Value = "No Preview"
        End If
        On Error GoTo 0

        ' --- GET MEASUREMENTS USING SPAWorkbench (FAST) ---
        dMass = 0: dVol = 0: dArea = 0
        dims(0) = 0: dims(1) = 0: dims(2) = 0

        Call GetBodyMeasurements(oPart, oBody, dMass, dVol, dArea, dims)

        ' --- GET MATERIAL ---
        strMatName = "N/A"
        dDensity = 0
        Call GetBodyMaterial(oPart, oBody, strMatName, dDensity)

        ' --- WRITE TO EXCEL ---
        xlSheet.Cells(r, 2).Value = bodyName

        If dMass > 0 Then
            xlSheet.Cells(r, 3).Value = Format(dMass, "0.000")
        Else
            xlSheet.Cells(r, 3).Value = "N/A"
        End If

        If dVol > 0 Then
            xlSheet.Cells(r, 4).Value = Format(dVol, "0.00")
        Else
            xlSheet.Cells(r, 4).Value = "N/A"
        End If

        If dArea > 0 Then
            xlSheet.Cells(r, 5).Value = Format(dArea, "0.00")
        Else
            xlSheet.Cells(r, 5).Value = "N/A"
        End If

        If dims(0) > 0 Then xlSheet.Cells(r, 6).Value = Format(dims(0), "0.0") Else xlSheet.Cells(r, 6).Value = "N/A"
        If dims(1) > 0 Then xlSheet.Cells(r, 7).Value = Format(dims(1), "0.0") Else xlSheet.Cells(r, 7).Value = "N/A"
        If dims(2) > 0 Then xlSheet.Cells(r, 8).Value = Format(dims(2), "0.0") Else xlSheet.Cells(r, 8).Value = "N/A"

        xlSheet.Cells(r, 9).Value = strMatName
        If dDensity > 0 Then
            xlSheet.Cells(r, 10).Value = Format(dDensity, "0.000")
        Else
            xlSheet.Cells(r, 10).Value = "N/A"
        End If

        ' --- HIDE THIS BODY AGAIN ---
        sel.Clear
        sel.Add oBody
        sel.VisProperties.SetShow 1  ' 1 = HIDE

        r = r + 1

NextBody:
        Set oBody = Nothing

        ' Status update every 10 bodies
        If i Mod 10 = 0 Then
            Application.StatusBar = "Processing body " & i & " of " & totalBodies
            DoEvents
        End If
    Next i

    ' --- RESTORE ALL BODIES ---
    Call ShowAllBodies(oPart, sel)

    ' --- RESTORE VIEW ---
    On Error Resume Next
    If g_CompassHidden Then CATIA.StartCommand "Compass"
    On Error GoTo 0

    ' --- FINALIZE EXCEL ---
    xlApp.ScreenUpdating = True
    xlSheet.Columns("A:J").AutoFit
    xlSheet.Columns("A:A").ColumnWidth = 15  ' Keep thumbnail column width

    ' --- COMPLETION MESSAGE ---
    Dim elapsed As Double
    elapsed = Timer - startTime

    sel.Clear
    MsgBox "BOM Exported Successfully!" & vbCrLf & _
           "Bodies processed: " & (r - 2) & vbCrLf & _
           "Time elapsed: " & Format(elapsed, "0.0") & " seconds", vbInformation

End Sub

' ==============================================================================
' HIDE ALL BODIES
' ==============================================================================
Sub HideAllBodies(oPart As Part, sel As Selection)
    On Error Resume Next
    Dim oBody As Body
    Dim i As Integer

    For i = 1 To oPart.Bodies.Count
        Set oBody = oPart.Bodies.Item(i)
        sel.Clear
        sel.Add oBody
        sel.VisProperties.SetShow 1  ' 1 = HIDE
    Next i

    sel.Clear
    Err.Clear
End Sub

' ==============================================================================
' SHOW ALL BODIES
' ==============================================================================
Sub ShowAllBodies(oPart As Part, sel As Selection)
    On Error Resume Next
    Dim oBody As Body
    Dim i As Integer

    For i = 1 To oPart.Bodies.Count
        Set oBody = oPart.Bodies.Item(i)
        sel.Clear
        sel.Add oBody
        sel.VisProperties.SetShow 0  ' 0 = SHOW
    Next i

    sel.Clear
    Err.Clear
End Sub

' ==============================================================================
' GET BODY MEASUREMENTS (Using SPAWorkbench - FAST)
' ==============================================================================
Sub GetBodyMeasurements(oPart As Part, oBody As Body, ByRef dMass As Double, _
                        ByRef dVol As Double, ByRef dArea As Double, ByRef dims() As Double)
    On Error Resume Next

    ' Method 1: Use SPAWorkbench Measurable (fastest)
    Dim oSPAWB As Object
    Dim oMeasurable As Object
    Dim oRef As Reference

    Set oSPAWB = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
    If Not oSPAWB Is Nothing Then
        Set oRef = oPart.CreateReferenceFromObject(oBody)
        Set oMeasurable = oSPAWB.GetMeasurable(oRef)

        If Not oMeasurable Is Nothing Then
            dVol = oMeasurable.Volume  ' in mm3
            dArea = oMeasurable.Area   ' in mm2

            ' Get bounding box for dimensions
            Dim bbox(5) As Variant
            oMeasurable.GetMinimumBoundingBox bbox

            Dim d1 As Double, d2 As Double, d3 As Double
            d1 = Abs(bbox(3) - bbox(0))  ' X dimension
            d2 = Abs(bbox(4) - bbox(1))  ' Y dimension
            d3 = Abs(bbox(5) - bbox(2))  ' Z dimension

            Call SortThree(d1, d2, d3, dims)

            Set oMeasurable = Nothing
        End If
        Set oRef = Nothing
        Set oSPAWB = Nothing
    End If

    ' Method 2: Use Inertia for mass (if SPAWorkbench didn't work)
    If dMass <= 0 Then
        Dim oInertia As Object
        Set oInertia = oPart.GetTechnologicalObject("Inertia")
        If Not oInertia Is Nothing Then
            ' Need to temporarily show this body for inertia calculation
            Dim sel As Selection
            Set sel = CATIA.ActiveDocument.Selection
            sel.Clear
            sel.Add oBody

            ' Get inertia data
            dMass = oInertia.Mass

            Set oInertia = Nothing
            Set sel = Nothing
        End If
    End If

    ' If still no volume, try calculating from inertia
    If dVol <= 0 And dMass > 0 Then
        ' Estimate volume from mass (assuming density ~7800 kg/m3 for steel)
        ' This is a rough estimate only
        dVol = (dMass / 7800) * 1000000000  ' Convert m3 to mm3
    End If

    Err.Clear
End Sub

' ==============================================================================
' GET BODY MATERIAL
' ==============================================================================
Sub GetBodyMaterial(oPart As Part, oBody As Body, ByRef matName As String, ByRef density As Double)
    On Error Resume Next

    Dim oManager As Object
    Set oManager = oPart.GetItem("CATMatManagerVBExt")

    If Not oManager Is Nothing Then
        Dim oMat As Object

        ' Try to get material from body first
        oManager.GetMaterialOnBody oBody, oMat

        ' If no material on body, try part level
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

    ' Fallback: Get density from part inertia
    If density <= 0 Then
        Dim oInertia As Object
        Set oInertia = oPart.GetTechnologicalObject("Inertia")
        If Not oInertia Is Nothing Then
            density = oInertia.Density
            Set oInertia = Nothing
        End If
    End If

    Err.Clear
End Sub

' ==============================================================================
' SORT THREE VALUES (Largest to smallest)
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
