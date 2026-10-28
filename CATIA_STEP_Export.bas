Option Explicit

' ==============================================================================
' MACRO: STEP EXPORT OF EVERY PART OF THE BOM - AND OF EVERY BODY INSIDE
'
' Usage: open the product (CATProduct) or a part (CATPart) in CATIA,
'        Tools > Macro > Macros... > ExportStepFiles > Run
'
' WHAT YOU GET (in C:\Temp\catia_step\<product>_<date>_<time>\):
'   <part number>.stp                one STEP file per part number of the
'                                    product - every part of the BOM list,
'                                    exported once however often it is used
'   <part number>__<body>.stp        one STEP file per BODY of a part with
'                                    several bodies (PartBody, "Grundplatte",
'                                    "Zapfen", ..., also the bodies inside
'                                    another body: Add, Assemble, Union Trim).
'                                    Tool bodies of Remove / Intersect are cut
'                                    away, not pieces of the part: no file.
'   step_index.csv                   one line per file: part number,
'                                    description, quantity in the product,
'                                    body, file name, status - opens in Excel
'   step_export_log.txt              one line per step of the run
' The file names use the part numbers of column C of the BOM, so every row
' of the BOM finds its STEP file.
'
' WHICH PARTS:
'   - an open CATPart: that part and its bodies
'   - an open CATProduct: every part below it, down to the LOWEST level of
'     the tree. With Teamcenter a sub-assembly whose content is not loaded
'     yet shows no children, just like a part - so a component counts as a
'     part only when something says it is one (its CATPart is loaded, it has
'     a shape, its file is a CATPart / cgr); otherwise its structure is
'     loaded and the search goes on below it. A "part" that still turns out
'     to have components once it is loaded is searched as well.
'     step_index.csv says where in the tree every part was found.
'   - a component selected in the tree when the macro starts: you are asked
'     whether only that component (and everything below it) is exported
'
' HOW:
'   - a part: its CATPart document is exported as it is (Document.ExportData).
'     Positions are those of the part itself (its own axis system), not of
'     its place in the product.
'   - a body: copied and pasted "as result" (no link) into a new, empty
'     CATPart, which is exported and closed again WITHOUT saving. Nothing is
'     changed in your parts.
'   - Teamcenter: like the BOM macros, every part is loaded one at a time
'     (design mode) - a big product takes as long as a BOM run without
'     pictures. A sub-assembly that is not loaded gets its structure loaded
'     first; if that shows nothing, it is loaded completely (design mode, one
'     long step) so that no part below it is missed. Only a component CATIA
'     cannot load at all is listed in step_index.csv ("its parts are missing").
'     A part whose part number shows only once it is loaded gets its files
'     under that part number. A part CATIA has only as a visualization (cgr)
'     cannot be exported: its line in step_index.csv says so.
'   - a body file is named inside like the file (not "Part1"), its body is
'     shown even when it is hidden in your part, and it must hold a solid -
'     otherwise its line in step_index.csv says so.
'
' STEP settings (AP214 or AP203, units, what a part file contains, hidden
' elements) are the ones of Tools > Options > General > Compatibility > STEP.
' The STEP export needs the STEP interface license of CATIA.
' ==============================================================================

' Windows must not offer to close CATIA while one big part loads (one call
' into CATIA that no macro can interrupt) - see the BOM macros
Private Declare PtrSafe Sub DisableProcessWindowsGhosting Lib "user32" ()

' --- SETTINGS -----------------------------------------------------------------
' where the folder of the run is made ("" = the Windows TEMP folder)
Private Const STEP_FOLDER As String = "C:\Temp\catia_step"
' True = one STEP file per body of a part with several bodies
Private Const EXPORT_BODIES As Boolean = True
' True = a body file also for a part with only ONE body (the same geometry as
' the part file, so normally off)
Private Const EXPORT_SINGLE_BODY As Boolean = False
' True = the bodies inside another body (taken in by Add, Assemble or Union
' Trim, e.g. PartBody > Assemble.1 > Body.2) get a file too
Private Const LIST_INNER_BODIES As Boolean = True
' True = a file of an earlier run with the same name is replaced
Private Const OVERWRITE_FILES As Boolean = True
' True = the folder of the run opens in the Explorer at the end
Private Const OPEN_FOLDER_AT_END As Boolean = True
' True = Windows does not offer to close CATIA while a big part loads
Private Const NO_NOT_RESPONDING_GHOST As Boolean = True
' the export format: "stp" (STEP). "igs" would give IGES files instead.
Private Const EXPORT_FORMAT As String = "stp"

' --- the run ------------------------------------------------------------------
Private mFso As Object
Private mFolder As String            ' the folder of this run
Private mLogPath As String
Private mParts As Collection         ' Array(part number, first instance, named provisionally, place in the tree)
Private mQty As Object               ' part number -> instances in the exported scope
Private mDesc As Object              ' part number -> description
Private mPath As Object              ' part number -> its first place in the tree ("root > sub > ...")
Private mRows As Collection          ' index lines: Array(pn, body, file, status) - desc and qty at writing
Private mFileNames As Object         ' file names used (lower case) - kept unique
Private mPump As Long
Private mNoPn As Long                ' components named before their part number could be read
Private mProvOwner As Object         ' provisional name -> the parent it belongs to (kept unique)
Private mExported As Object          ' part numbers already worked on in this run
Private mNotLoaded As Long           ' components CATIA could not load - their parts are missing
Private mNodes As Long               ' components visited by the tree search
Private mMaxDepth As Long            ' deepest level reached
Private mStructLoads As Long         ' components whose structure had to be loaded first
Private mLateAssy As Long            ' "parts" that turned out to be sub-assemblies once loaded
Private mDesignLoads As Long         ' components loaded completely because their structure did not come
Private mFailed As Object            ' part number -> pieces of a component CATIA could not load (tried once)


' ==============================================================================
' ENTRY POINT
' ==============================================================================
Sub ExportStepFiles()
    Dim catDoc As Document, rootProd As Product, rootWin As Object
    Dim scopeProd As Product, isPartDoc As Boolean, ans As Long
    Dim rootPN As String, scopeName As String
    Dim i As Long, it As Variant, pn As String, realPn As String, realDesc As String, nk As Long
    Dim mult As Long
    Dim oInst As Product, oPart As Part, oPartDoc As PartDocument
    Dim nParts As Long, nPartFiles As Long, nBodyFiles As Long, nFailed As Long
    Dim startTime As Double, failMsg As String, statusTxt As String

    startTime = Timer
    mLogPath = ""
    Set mFso = CreateObject("Scripting.FileSystemObject")
    Set mParts = New Collection
    Set mRows = New Collection
    Set mQty = CreateObject("Scripting.Dictionary")
    Set mDesc = CreateObject("Scripting.Dictionary")
    Set mPath = CreateObject("Scripting.Dictionary")
    Set mFileNames = CreateObject("Scripting.Dictionary")
    Set mProvOwner = CreateObject("Scripting.Dictionary")
    Set mExported = CreateObject("Scripting.Dictionary")
    mPump = 0
    mNoPn = 0
    mNotLoaded = 0
    mNodes = 0
    mMaxDepth = 0
    mStructLoads = 0
    mLateAssy = 0
    mDesignLoads = 0
    Set mFailed = CreateObject("Scripting.Dictionary")

    ' --- 1. DOCUMENT -----------------------------------------------------------
    On Error Resume Next
    Set catDoc = CATIA.ActiveDocument
    Set rootWin = CATIA.ActiveWindow
    If Err.Number <> 0 Or catDoc Is Nothing Then
        On Error GoTo 0
        MsgBox "No document open.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    isPartDoc = (InStr(1, catDoc.Name, ".CATPart", vbTextCompare) > 0)
    If Not isPartDoc And InStr(1, catDoc.Name, ".CATProduct", vbTextCompare) = 0 Then
        MsgBox "Please open a product (CATProduct) or a part (CATPart) first.", vbExclamation
        Exit Sub
    End If
    Set rootProd = catDoc.Product
    rootPN = ""
    On Error Resume Next
    rootPN = rootProd.PartNumber
    Err.Clear
    On Error GoTo 0
    If Len(rootPN) = 0 Then rootPN = BaseFileName(catDoc.Name)

    ' --- 2. WHAT IS EXPORTED: the whole product, or the selected component --
    Set scopeProd = rootProd
    If Not isPartDoc Then
        Set oInst = SelectedComponent(catDoc, rootProd)
        If Not oInst Is Nothing Then
            ans = MsgBox("A component is selected in the tree:" & vbCrLf & "   " & InstanceName(oInst) & _
                         vbCrLf & vbCrLf & _
                         "Yes = export only this component (and everything below it)" & vbCrLf & _
                         "No  = export the whole product", _
                         vbYesNoCancel + vbQuestion + vbDefaultButton2, "STEP export")
            If ans = vbCancel Then Exit Sub
            If ans = vbYes Then Set scopeProd = oInst
        End If
    End If
    scopeName = rootPN
    If Not scopeProd Is rootProd Then scopeName = InstanceName(scopeProd)

    If MsgBox("Export STEP files of every part" & IIf(EXPORT_BODIES, " and of every body of a multi-body part", "") & _
              vbCrLf & "of:  " & scopeName & vbCrLf & vbCrLf & _
              "Every part is loaded once; a big product takes a while." & vbCrLf & _
              "The files go to a new folder below " & OutputRoot() & ".", _
              vbOKCancel + vbInformation, "STEP export") = vbCancel Then Exit Sub

    ' --- 3. SESSION, FOLDER, LOG -----------------------------------------------
    On Error Resume Next
    If NO_NOT_RESPONDING_GHOST Then DisableProcessWindowsGhosting
    CATIA.DisplayFileAlerts = False
    Err.Clear
    On Error GoTo Fail

    Call StartRunFolder(scopeName)
    Call LogLine("STEP export of " & scopeName & "  (" & catDoc.Name & ")")
    Call LogLine("folder " & mFolder)

    ' --- 4. THE PARTS ----------------------------------------------------------
    Call SetStatus("STEP export: reading the product structure ...")
    If isPartDoc Then
        Call AddPart(rootProd, "", 1)
    Else
        ' down to the lowest level of the tree (a selected part is one part)
        Call CollectNode(scopeProd, "", 0, 1)
    End If
    nParts = mParts.Count
    Call LogLine("tree searched: " & mNodes & " components, " & mMaxDepth & " levels deep, " & _
                 mStructLoads & " structures loaded first, " & mDesignLoads & " components loaded completely")
    Call LogLine(nParts & " part numbers to export (" & mNoPn & " named before their part number could be " & _
                 "read - renamed once loaded), " & mNotLoaded & " components not loaded")
    If nParts = 0 Then
        If mNotLoaded > 0 Then
            failMsg = "CATIA could not load " & scopeName & ", or it is empty - nothing to export." & vbCrLf & _
                      "Load it in CATIA (Teamcenter) and run the export again."
        Else
            failMsg = "No part was found below " & scopeName & "."
        End If
        GoTo Finish
    End If

    ' --- 5. EXPORT -------------------------------------------------------------
    ' (a "part" that turns out to be a sub-assembly adds its parts to the list
    '  while the loop runs - so the list is walked until its end, not a count)
    i = 0
    Do While i < mParts.Count
        i = i + 1
        nParts = mParts.Count
        it = mParts(i)
        pn = CStr(it(0))
        Set oInst = it(1)
        Call SetStatus("STEP export: part " & i & " of " & nParts & " - " & pn & " (loading ...)")
        Call LogLine("part " & i & "/" & nParts & "  " & pn & "   in " & CStr(it(3)))

        nk = 0
        Set oPart = Nothing
        Set oPartDoc = Nothing
        If isPartDoc Then
            Set oPartDoc = catDoc
            Set oPart = oPartDoc.Part
        Else
            If Not LoadPart(oInst, oPart, oPartDoc) Then Set oPart = Nothing
            ' once loaded it has components of its own: a sub-assembly whose
            ' content was not loaded when the tree was searched - also when its
            ' shape representation is a CATPart (LoadPart then found that)
            On Error Resume Next
            nk = oInst.Products.Count
            Err.Clear
            On Error GoTo Fail
        End If

        ' a sub-assembly: the search goes on below it, for every piece of it
        ' that was counted, and it is not exported as a part. Its entry is
        ' cleared first, so a CATPart inside it with the same part number (a
        ' Teamcenter item product and its CATPart) gets an entry of its own.
        If nk > 0 Then
            mLateAssy = mLateAssy + 1
            Call LogLine("  " & pn & " is a sub-assembly (" & nk & " components) - its parts are searched too")
            mult = CLng(mQty(pn))
            mQty.Remove pn
            If mDesc.Exists(pn) Then mDesc.Remove pn
            If mPath.Exists(pn) Then mPath.Remove pn
            Call CollectNode(oInst, CStr(it(3)), 1, mult)
            Call LogLine("  the list has " & mParts.Count & " part numbers now")
            GoTo NextPart
        End If

        ' named before its part number could be read: now that it is loaded,
        ' its real part number and description - the names of the BOM
        If it(2) And Not oPart Is Nothing Then
            realPn = ""
            realDesc = ""
            On Error Resume Next
            realPn = oInst.PartNumber
            If Len(realPn) = 0 Then realPn = oPartDoc.Product.PartNumber
            realDesc = oInst.DescriptionRef
            If Len(realDesc) = 0 Then realDesc = oPartDoc.Product.DescriptionRef
            Err.Clear
            On Error GoTo Fail
            If Len(realPn) > 0 And realPn <> pn Then
                Call LogLine("  part number of " & pn & ": " & realPn)
                ' its pieces count for the real part number
                If mQty.Exists(realPn) Then
                    mQty(realPn) = mQty(realPn) + mQty(pn)
                Else
                    mQty.Add realPn, mQty(pn)
                    mDesc.Add realPn, ""
                End If
                If Not mPath.Exists(realPn) Then mPath.Add realPn, CStr(it(3))
                pn = realPn
            End If
            ' (also when the provisional name already was the part number)
            If Len(realDesc) > 0 And mDesc.Exists(pn) Then
                If Len(CStr(mDesc(pn))) = 0 Then mDesc(pn) = realDesc
            End If
        End If

        If mExported.Exists(pn) Then
            ' the same part as an earlier entry (found under a provisional name)
            Call LogLine("  " & pn & " is exported already")
        ElseIf oPart Is Nothing Then
            mExported.Add pn, True
            statusTxt = "not exported: CATIA has no CATPart of it (not loaded, or only a visualization)"
            Call AddRow(pn, "", "", statusTxt)
            Call LogLine("  " & statusTxt)
            nFailed = nFailed + 1
        Else
            mExported.Add pn, True
            Call ExportPart(i, nParts, pn, oPart, oPartDoc, nPartFiles, nBodyFiles, nFailed)
        End If
NextPart:
        Call SetStatus("")
        DoEvents
    Loop
    nParts = mParts.Count
    Call LogLine(mLateAssy & " components found to be sub-assemblies only once loaded")

Finish:
    ' --- 6. INDEX, RESTORE, REPORT ---------------------------------------------
    On Error Resume Next
    Call WriteIndex
    If Not rootWin Is Nothing Then rootWin.Activate
    CATIA.DisplayFileAlerts = True
    CATIA.StatusBar = ""
    Err.Clear
    On Error GoTo 0
    Call LogLine("finished after " & Format(Elapsed(startTime), "0") & " s: " & nPartFiles & " part files, " & _
                 nBodyFiles & " body files, " & nFailed & " not exported")

    If Len(failMsg) > 0 Then
        MsgBox failMsg, vbExclamation, "STEP export"
        Exit Sub
    End If
    If OPEN_FOLDER_AT_END Then
        On Error Resume Next
        Shell "explorer.exe """ & mFolder & """", vbNormalFocus
        Err.Clear
        On Error GoTo 0
    End If
    MsgBox "STEP export finished" & vbCrLf & vbCrLf & _
           "Part files:   " & nPartFiles & "  (of " & mExported.Count & " part numbers)" & vbCrLf & _
           "Body files:   " & nBodyFiles & vbCrLf & _
           "Not exported: " & nFailed & IIf(nFailed > 0, "  - see step_index.csv", "") & vbCrLf & _
           IIf(mNotLoaded > 0, "Components CATIA could not load: " & mNotLoaded & _
               "  - their parts are missing (see step_index.csv)" & vbCrLf, "") & _
           "Time: " & Format(Elapsed(startTime) / 60#, "0.0") & " min" & vbCrLf & vbCrLf & _
           "Folder:" & vbCrLf & mFolder, IIf(nFailed > 0 Or mNotLoaded > 0, vbExclamation, vbInformation), _
           "STEP export"
    Exit Sub

Fail:
    failMsg = "The STEP export stopped with error " & Err.Number & " - " & Err.Description
    Resume FailCleanup

FailCleanup:
    On Error Resume Next
    Call LogLine(failMsg)
    Call WriteIndex
    If Not rootWin Is Nothing Then rootWin.Activate
    CATIA.DisplayFileAlerts = True
    CATIA.StatusBar = ""
    Err.Clear
    MsgBox failMsg & vbCrLf & vbCrLf & _
           "The files exported so far are in:" & vbCrLf & mFolder, vbCritical, "STEP export"
End Sub


' ==============================================================================
' ONE PART: its STEP file, then one file per body
' ==============================================================================
Private Sub ExportPart(ByVal no As Long, ByVal nParts As Long, ByVal pn As String, _
                       oPart As Part, oPartDoc As PartDocument, ByRef nPartFiles As Long, _
                       ByRef nBodyFiles As Long, ByRef nFailed As Long)
    Dim fileName As String, statusTxt As String
    Dim names() As String, objs As Object, n As Long, k As Long
    Dim partWin As Object, t As Double

    ' --- the part ---
    t = Timer
    Call SetStatus("STEP export: part " & no & " of " & nParts & " - " & pn)
    fileName = UniqueFileName(SafeName(pn))
    statusTxt = ExportDocument(oPartDoc, fileName)
    Call AddRow(pn, "", fileName, statusTxt)
    Call LogLine("  " & fileName & ": " & statusTxt & "  (" & Format(Elapsed(t), "0.0") & " s)")
    If statusTxt = "OK" Then nPartFiles = nPartFiles + 1 Else nFailed = nFailed + 1

    ' --- its bodies ---
    If Not EXPORT_BODIES Then Exit Sub
    Set objs = CreateObject("Scripting.Dictionary")
    n = ListBodies(oPart, names, objs)
    If n = 0 Then Exit Sub
    If n = 1 And Not EXPORT_SINGLE_BODY Then Exit Sub
    Call LogLine("  " & n & " bodies")

    Set partWin = Nothing
    For k = 0 To n - 1
        t = Timer
        Call SetStatus("STEP export: part " & no & " of " & nParts & " - " & pn & _
                       " - body " & (k + 1) & " of " & n & " (" & names(k) & ")")
        fileName = UniqueFileName(SafeName(pn & "__" & names(k)))
        statusTxt = ExportBody(oPartDoc, objs(names(k)), fileName, BaseFileName(fileName), partWin)
        Call AddRow(pn, names(k), fileName, statusTxt)
        Call LogLine("  " & fileName & ": " & statusTxt & "  (" & Format(Elapsed(t), "0.0") & " s)")
        If statusTxt = "OK" Then nBodyFiles = nBodyFiles + 1 Else nFailed = nFailed + 1
        DoEvents
    Next k

    ' a window the body copies needed is closed again
    If Not partWin Is Nothing Then
        On Error Resume Next
        partWin.Close
        Err.Clear
        On Error GoTo 0
    End If
End Sub

' a whole document into one STEP file; "OK" or why not
Private Function ExportDocument(oDoc As Object, ByVal fileName As String) As String
    Dim pth As String, e As Long, d As String

    pth = mFolder & "\" & fileName
    If Not FreeTarget(pth) Then
        ExportDocument = "not exported: the file exists and cannot be replaced"
        Exit Function
    End If
    On Error Resume Next
    Err.Clear
    oDoc.ExportData pth, EXPORT_FORMAT
    e = Err.Number
    d = Err.Description
    Err.Clear
    On Error GoTo 0
    If mFso.FileExists(pth) Then
        ExportDocument = "OK"
    ElseIf e <> 0 Then
        ExportDocument = "not exported: error " & e & " - " & d
    Else
        ExportDocument = "not exported: CATIA wrote no file"
    End If
End Function

' ==============================================================================
' ONE BODY: copied "as result" into a new, empty CATPart, exported, and that
' CATPart closed again without saving. partWin: a window of the source part,
' opened when copying needs one (a part loaded only inside the product has
' none) - the caller closes it after the last body.
' ==============================================================================
Private Function ExportBody(oPartDoc As PartDocument, ByVal oBody As Object, ByVal fileName As String, _
                            ByVal docName As String, ByRef partWin As Object) As String
    Dim srcSel As Object, newDoc As Object, newSel As Object, newPart As Object
    Dim newBody As Object, nBefore As Long, nAfter As Long, nShapes As Long
    Dim e As Long, d As String, pth As String, st As String

    ExportBody = "not exported"
    pth = mFolder & "\" & fileName
    If Not FreeTarget(pth) Then
        ExportBody = "not exported: the file exists and cannot be replaced"
        Exit Function
    End If

    On Error Resume Next
    ' --- copy the body ---
    Set srcSel = oPartDoc.Selection
    e = CopyToClipboard(srcSel, oBody)
    If e <> 0 And partWin Is Nothing Then
        ' copying needs the part in a window of its own
        Err.Clear
        Set partWin = oPartDoc.NewWindow
        If Not partWin Is Nothing Then partWin.Activate
        Err.Clear
        Call Pause(300)
        Set srcSel = oPartDoc.Selection
        e = CopyToClipboard(srcSel, oBody)
    End If
    If e <> 0 Then
        ExportBody = "not exported: CATIA could not copy the body (error " & e & ")"
        Err.Clear
        Exit Function
    End If

    ' --- paste it as result into a new, empty part ---
    Err.Clear
    Set newDoc = Nothing
    Set newDoc = CATIA.Documents.Add("Part")
    If newDoc Is Nothing Then
        ExportBody = "not exported: CATIA could not make a new part (error " & Err.Number & ")"
        Err.Clear
        Exit Function
    End If
    Set newPart = newDoc.Part
    ' inside the STEP file the part is named like the file, not "Part1"
    newDoc.Product.PartNumber = docName
    Err.Clear
    nBefore = 0
    nBefore = newPart.Bodies.Count
    Set newSel = newDoc.Selection
    newSel.Clear
    newSel.Add newPart
    Err.Clear
    newSel.PasteSpecial "CATPrtResultWithOutLink"
    e = Err.Number
    d = Err.Description
    newSel.Clear
    Err.Clear
    newPart.Update
    Err.Clear

    ' the pasted body: there, with a solid in it, and shown - a hidden body
    ' (it keeps the look of the source) may not be written to the file
    nAfter = 0
    nAfter = newPart.Bodies.Count
    Set newBody = Nothing
    nShapes = 0
    If nAfter > nBefore Then
        Set newBody = newPart.Bodies.Item(nAfter)
    ElseIf nAfter > 0 Then
        Set newBody = newPart.Bodies.Item(1)       ' pasted into the PartBody itself
    End If
    If Not newBody Is Nothing Then
        nShapes = newBody.Shapes.Count
        newSel.Clear
        newSel.Add newBody
        newSel.VisProperties.SetShow 0
        newSel.Clear
    End If
    Err.Clear

    ' --- export it, close it without saving ---
    If e <> 0 Then
        ExportBody = "not exported: CATIA could not paste the body (error " & e & " - " & d & ")"
    ElseIf newBody Is Nothing Or nShapes = 0 Then
        ExportBody = "not exported: nothing was pasted (the body has no solid of its own)"
    Else
        st = ExportDocument(newDoc, fileName)
        ' (the solid check reads STEP text; an IGES file is not checked)
        If st = "OK" And LCase$(EXPORT_FORMAT) = "stp" Then
            If Not HasSolid(pth) Then st = "exported, but the file holds no solid - check it"
        End If
        ExportBody = st
    End If
    newDoc.Close
    Err.Clear
    On Error GoTo 0
End Function

' the body alone in the selection, copied; 0 = done, else the error
Private Function CopyToClipboard(srcSel As Object, ByVal oBody As Object) As Long
    Dim e As Long
    On Error Resume Next
    Err.Clear
    srcSel.Clear
    srcSel.Add oBody
    If Err.Number = 0 Then srcSel.Copy
    e = Err.Number
    srcSel.Clear
    Err.Clear
    CopyToClipboard = e
End Function

' the STEP file holds a solid (a body file must) - read in pieces, a big
' file is never held in memory at once
Private Function HasSolid(ByVal pth As String) As Boolean
    Const CHUNK As Long = 1048576
    Dim f As Integer, n As Long, pos As Long, take As Long
    Dim buf As String, tail As String, txt As String

    HasSolid = False
    On Error Resume Next
    f = FreeFile
    Err.Clear
    Open pth For Binary Access Read As #f
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    n = LOF(f)
    pos = 1
    tail = ""
    Do While pos <= n
        take = CHUNK
        If pos + take - 1 > n Then take = n - pos + 1
        buf = Space$(take)
        Get #f, pos, buf
        txt = UCase$(tail & buf)
        If InStr(txt, "MANIFOLD_SOLID_BREP") > 0 Or InStr(txt, "BREP_WITH_VOIDS") > 0 Or _
           InStr(txt, "FACETED_BREP") > 0 Then
            HasSolid = True
            Exit Do
        End If
        ' (a word cut at the end of a piece is found with the next one)
        If Len(txt) > 40 Then tail = Right$(txt, 40) Else tail = txt
        pos = pos + take
    Loop
    Close #f
    Err.Clear
End Function

' an older file of the same name out of the way; False = it cannot be replaced
Private Function FreeTarget(ByVal pth As String) As Boolean
    FreeTarget = True
    If Not mFso.FileExists(pth) Then Exit Function
    If Not OVERWRITE_FILES Then
        FreeTarget = False
        Exit Function
    End If
    On Error Resume Next
    mFso.DeleteFile pth, True
    Err.Clear
    On Error GoTo 0
    FreeTarget = Not mFso.FileExists(pth)
End Function


' ==============================================================================
' THE PARTS BELOW A COMPONENT - every part number once, in the order of the
' tree. A component Teamcenter has not loaded shows no children: its
' structure is loaded first (visualization mode, like the BOM macros), so a
' sub-assembly is never taken for a part and loaded completely in one go. One
' that still cannot be loaded is listed as not loaded (its parts are missing).
'   ownerKey  the path of the parent - keeps provisional names apart
' ==============================================================================
Private Sub CollectNode(oProd As Product, ByVal ownerKey As String, ByVal depth As Long, _
                        ByVal mult As Long)
    Dim i As Long, nk As Long, nm As String, here As String
    Dim child As Product

    If depth > 60 Then Exit Sub
    mNodes = mNodes + 1
    If depth > mMaxDepth Then mMaxDepth = depth
    nk = ChildCount(oProd, ownerKey)
    If nk = 0 Then
        Call AddPart(oProd, ownerKey, mult)
        Exit Sub
    ElseIf nk < 0 Then
        Call NoteNotLoaded(oProd, ownerKey, mult)
        Exit Sub
    End If

    ' the place of the children in the tree: "root > sub-assembly > ..."
    nm = InstanceName(oProd)
    If Len(ownerKey) > 0 Then here = ownerKey & " > " & nm Else here = nm
    For i = 1 To nk
        Set child = Nothing
        On Error Resume Next
        Set child = oProd.Products.Item(i)
        Err.Clear
        On Error GoTo 0
        If Not child Is Nothing Then Call CollectNode(child, here, depth + 1, mult)
        mPump = mPump + 1
        If mPump Mod 50 = 0 Then DoEvents
        If mPump Mod 500 = 0 Then Call SetStatus("STEP export: searching the product tree ... " & _
                                                 mPump & " components, level " & mMaxDepth)
    Next i
End Sub

' the children of a component. Teamcenter loads on demand: a sub-assembly
' whose content is not loaded yet shows NO children, exactly like a part. So a
' childless component counts as a part only when something says it is one
' (its CATPart is in the session, it has a shape, its file is a CATPart / cgr);
' otherwise its structure is loaded first (visualization mode) and the search
' goes on below it - down to the lowest level of the tree.
' A sub-assembly whose structure does not come in visualization mode is
' loaded completely (design mode) as a last resort, so no part below it is
' missed. -1 = a component CATIA cannot load even then.
Private Function ChildCount(oProd As Product, ByVal ownerKey As String) As Long
    Dim nk As Long, pn As String

    nk = 0
    pn = ""
    On Error Resume Next
    nk = oProd.Products.Count
    pn = oProd.PartNumber
    Err.Clear
    ' another piece of a component that could not be loaded: not tried again
    If nk = 0 Then
        If mFailed.Exists(FailKey(oProd, ownerKey)) Then
            On Error GoTo 0
            ChildCount = -1
            Exit Function
        End If
    End If
    If nk = 0 Then
        If Len(pn) = 0 Or IsUnloadedAssembly(oProd) Or Not HasPartEvidence(oProd) Then
            mStructLoads = mStructLoads + 1
            Call SetStatus("STEP export: loading the structure of " & InstanceName(oProd) & " ...")
            Err.Clear
            oProd.ApplyWorkMode 1                 ' visualization: the structure only
            Err.Clear
            nk = 0
            nk = oProd.Products.Count
            pn = ""
            pn = oProd.PartNumber
            Err.Clear
            If nk = 0 Then
                If IsUnloadedAssembly(oProd) Or Not HasPartEvidence(oProd) Then
                    ' still nothing: as a last resort design mode - for a
                    ' sub-assembly that loads everything below it in one go
                    ' (each of those parts is loaded for its STEP file anyway),
                    ' so no part below it is missed
                    mDesignLoads = mDesignLoads + 1
                    Call SetStatus("STEP export: loading " & InstanceName(oProd) & " completely (design mode) ...")
                    Call LogLine("loading " & InstanceName(oProd) & " completely (design mode) - its structure did not come")
                    Err.Clear
                    oProd.ApplyWorkMode 2
                    Err.Clear
                    nk = 0
                    nk = oProd.Products.Count
                    pn = ""
                    pn = oProd.PartNumber
                    Err.Clear
                    ' still no component and nothing that says it is a part
                    If nk = 0 Then
                        If IsUnloadedAssembly(oProd) Or Not HasPartEvidence(oProd) Then nk = -1
                    End If
                End If
            End If
        End If
    End If
    On Error GoTo 0
    ChildCount = nk
End Function

' a component CATIA cannot load: one line in the index, nothing exported
Private Sub NoteNotLoaded(oProd As Product, ByVal ownerKey As String, ByVal mult As Long)
    Dim nm As String, pn As String

    nm = InstanceName(oProd)
    pn = FailKey(oProd, ownerKey)
    ' one line per component; its Qty = the pieces that are missing
    If mFailed.Exists(pn) Then
        mFailed(pn) = mFailed(pn) + mult
        Exit Sub
    End If
    mFailed.Add pn, mult
    If Not mPath.Exists(pn) Then mPath.Add pn, ownerKey
    mNotLoaded = mNotLoaded + 1
    Call AddRow(pn, "", "", "not exported: CATIA could not load " & nm & _
                " (a sub-assembly that is not loaded or empty, or no part number) - its parts are missing")
    Call LogLine("NOT LOADED: " & nm)
End Sub

' the name a component CATIA cannot load is noted under: its part number, else
' the name of its file (the same for every piece of it), else its instance
' name with its place in the tree (an instance name repeats below other
' parents)
Private Function FailKey(oProd As Product, ByVal ownerKey As String) As String
    Dim k As String, pth As String

    k = ""
    On Error Resume Next
    k = oProd.PartNumber
    Err.Clear
    If Len(k) = 0 Then
        pth = ""
        pth = CStr(oProd.GetMasterShapeRepresentationPathName)
        Err.Clear
        If Len(pth) > 0 Then k = BaseFileName(Mid$(pth, InStrRev(Replace(pth, "/", "\"), "\") + 1))
    End If
    If Len(k) = 0 Then
        k = InstanceName(oProd)
        If Len(ownerKey) > 0 Then k = k & "  (in " & ownerKey & ")"
    End If
    Err.Clear
    On Error GoTo 0
    FailKey = k
End Function

' notes one part instance: a new part number, or one more piece of a known one.
' it(2) of the list = True: named before its part number could be read
Private Sub AddPart(oProd As Product, ByVal ownerKey As String, ByVal mult As Long)
    Dim pn As String, desc As String, prov As Boolean

    pn = ""
    desc = ""
    On Error Resume Next
    pn = oProd.PartNumber
    desc = oProd.DescriptionRef
    Err.Clear
    On Error GoTo 0
    prov = (Len(pn) = 0)
    If prov Then
        ' not loaded yet: named by its file or instance until it is loaded
        pn = ProvisionalName(oProd, ownerKey)
        mNoPn = mNoPn + 1
    End If
    If mQty.Exists(pn) Then
        mQty(pn) = mQty(pn) + mult
    Else
        mQty.Add pn, mult
        mDesc.Add pn, desc
        mParts.Add Array(pn, oProd, prov, ownerKey)
        If Not mPath.Exists(pn) Then mPath.Add pn, ownerKey
    End If
End Sub

' the name of a component whose part number CATIA cannot read yet: its file,
' else its instance name - an instance name is unique only below its own
' parent, so the same name below another parent gets " ~2", " ~3" ...
Private Function ProvisionalName(oProd As Product, ByVal ownerKey As String) As String
    Dim nm As String, pth As String, base As String, k As Long

    nm = ""
    On Error Resume Next
    pth = ""
    pth = CStr(oProd.GetMasterShapeRepresentationPathName)
    Err.Clear
    If Len(pth) > 0 Then nm = BaseFileName(Mid$(pth, InStrRev(Replace(pth, "/", "\"), "\") + 1))
    If Len(nm) = 0 Then
        nm = oProd.Name
        Err.Clear
        If Len(nm) = 0 Then nm = "unnamed"
        base = nm
        k = 1
        Do While mProvOwner.Exists(nm)
            If CStr(mProvOwner(nm)) = ownerKey Then Exit Do
            k = k + 1
            nm = base & " ~" & k
        Loop
        If Not mProvOwner.Exists(nm) Then mProvOwner.Add nm, ownerKey
    End If
    Err.Clear
    On Error GoTo 0
    If Len(nm) = 0 Then nm = "unnamed"
    ProvisionalName = nm
End Function

' something says this component IS a part: its CATPart is in the session, it
' has a shape, or its file is a CATPart / cgr - nothing is loaded to find out
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

' the CATPart of a part instance - promoted to design mode, only this part
Private Function LoadPart(oInst As Product, ByRef oPart As Part, ByRef oPartDoc As PartDocument) As Boolean
    Dim d As Object, ok As Boolean, t As Double

    Set oPart = Nothing
    Set oPartDoc = Nothing
    ok = False
    t = Timer
    On Error Resume Next
    ' already loaded as a CATPart?
    If TypeName(oInst.ReferenceProduct.Parent) = "PartDocument" Then
        Set oPartDoc = oInst.ReferenceProduct.Parent
        Set oPart = oPartDoc.Part
        ok = Not (oPart Is Nothing)
    End If
    Err.Clear
    If Not ok Then
        ' design mode loads the CATPart - a part has no children, so this is
        ' one file, never a whole sub-assembly
        oInst.ApplyWorkMode 2
        Err.Clear
        If TypeName(oInst.ReferenceProduct.Parent) = "PartDocument" Then
            Set oPartDoc = oInst.ReferenceProduct.Parent
            Set oPart = oPartDoc.Part
            ok = Not (oPart Is Nothing)
        End If
        Err.Clear
    End If
    If Not ok Then
        ' a CATProduct whose geometry is a CATPart (shape representation)
        Set d = Nothing
        Set d = oInst.ReferenceProduct.GetMasterShapeRepresentation(True)
        If d Is Nothing Then
            Err.Clear
            Set d = oInst.GetMasterShapeRepresentation(True)
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
    End If
    On Error GoTo 0
    If Elapsed(t) >= 5 Then Call LogLine("  loading took " & Format(Elapsed(t), "0.0") & " s")
    LoadPart = ok
End Function

' a childless node that is a sub-assembly not loaded yet - decided WITHOUT
' loading anything (same rules as in the BOM macros)
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

' the component of the first selected element (a product, or the part a
' selected body / face belongs to) - Nothing when nothing below the root is
' selected
Private Function SelectedComponent(catDoc As Document, rootProd As Product) As Product
    Dim oSel As Object, it As Object, v As Object, p As Product
    Dim n As Long, nm As String, rootNm As String

    Set p = Nothing
    On Error Resume Next
    Set oSel = catDoc.Selection
    n = 0
    n = oSel.Count
    If n > 0 Then
        Set it = oSel.Item(1)
        Set v = it.Value
        If TypeName(v) = "Product" Then Set p = v
        If p Is Nothing Then Set p = it.LeafProduct
    End If
    If Not p Is Nothing Then
        ' the root product itself is "the whole product"
        nm = ""
        nm = p.Name
        rootNm = ""
        rootNm = rootProd.Name
        If p Is rootProd Or (Len(nm) > 0 And nm = rootNm) Then Set p = Nothing
    End If
    Err.Clear
    On Error GoTo 0
    Set SelectedComponent = p
End Function


' ==============================================================================
' THE BODIES OF A PART - the same bodies the BOM lists, in the order of the tree
'   names(0..n-1)  body names;  objs: name -> Body
' The bodies directly below the part, and with LIST_INNER_BODIES the bodies
' taken in by Add, Assemble or Union Trim (one level deeper in the tree).
' Tool bodies of Remove / Intersect and bodies without geometry are left out.
' ==============================================================================
Private Function ListBodies(oPart As Part, ByRef names() As String, objs As Object) As Long
    Dim dObj As Object, dKids As Object, dGeo As Object, dIsOperand As Object, dDone As Object
    Dim i As Long, n As Long, nBodies As Long, oBody As Object, k As Variant

    ListBodies = 0
    Set dObj = CreateObject("Scripting.Dictionary")
    Set dKids = CreateObject("Scripting.Dictionary")
    Set dGeo = CreateObject("Scripting.Dictionary")
    Set dIsOperand = CreateObject("Scripting.Dictionary")
    Set dDone = CreateObject("Scripting.Dictionary")
    If oPart Is Nothing Then Exit Function

    On Error Resume Next
    nBodies = 0
    nBodies = oPart.Bodies.Count
    Err.Clear
    For i = 1 To nBodies
        Set oBody = Nothing
        Set oBody = oPart.Bodies.Item(i)
        Err.Clear
        If Not oBody Is Nothing Then Call ScanBody(oBody, dObj, dKids, dGeo, dIsOperand, 0)
    Next i
    Err.Clear
    On Error GoTo 0
    If dObj.Count = 0 Then Exit Function

    ReDim names(0 To dObj.Count - 1)
    n = 0
    For Each k In dObj.Keys
        If Not dIsOperand.Exists(CStr(k)) Then Call EmitBody(CStr(k), 1, dKids, dGeo, dDone, names, n)
    Next k
    For Each k In dObj.Keys
        If Not dDone.Exists(CStr(k)) Then Call EmitBody(CStr(k), 1, dKids, dGeo, dDone, names, n)
    Next k
    If n = 0 Then Exit Function
    ReDim Preserve names(0 To n - 1)
    For i = 0 To n - 1
        If Not objs.Exists(names(i)) Then objs.Add names(i), dObj(names(i))
    Next i
    ListBodies = n
End Function

' one body and - through the boolean operations among its shapes - every body
' it takes in (only boolean operations have a .Body: the body they take in)
Private Sub ScanBody(ByVal oBody As Object, dObj As Object, dKids As Object, dGeo As Object, _
                     dIsOperand As Object, ByVal depth As Long)
    Dim nm As String, nm2 As String, tn As String
    Dim j As Long, nShapes As Long, isTool As Boolean
    Dim shp As Object, ob As Object, kids As Collection

    On Error Resume Next
    nm = ""
    nm = oBody.Name
    Err.Clear
    If Len(nm) = 0 Then Exit Sub
    If dObj.Exists(nm) Then Exit Sub
    dObj.Add nm, oBody
    Set kids = New Collection
    dKids.Add nm, kids
    nShapes = 0
    nShapes = oBody.Shapes.Count
    Err.Clear
    dGeo.Add nm, (nShapes > 0)
    If depth >= 30 Then Exit Sub

    For j = 1 To nShapes
        Set shp = Nothing
        Set shp = oBody.Shapes.Item(j)
        Set ob = Nothing
        If Not shp Is Nothing Then Set ob = shp.Body
        Err.Clear
        If Not ob Is Nothing Then
            If TypeName(ob) = "Body" Then
                nm2 = ""
                nm2 = ob.Name
                tn = TypeName(shp)
                Err.Clear
                ' Remove and Intersect only use the body as a tool
                isTool = (InStr(1, tn, "Remove", vbTextCompare) > 0 Or _
                          InStr(1, tn, "Intersect", vbTextCompare) > 0)
                If Len(nm2) > 0 And nm2 <> nm Then
                    kids.Add Array(nm2, isTool)
                    If Not dIsOperand.Exists(nm2) Then dIsOperand.Add nm2, True
                    Call ScanBody(ob, dObj, dKids, dGeo, dIsOperand, depth + 1)
                End If
            End If
        End If
        If j Mod 50 = 0 Then DoEvents
    Next j
    Err.Clear
End Sub

' one body into the list, then the bodies it takes in
Private Sub EmitBody(ByVal nm As String, ByVal depth As Long, dKids As Object, dGeo As Object, _
                     dDone As Object, names() As String, ByRef n As Long)
    Dim it As Variant

    If dDone.Exists(nm) Then Exit Sub
    dDone.Add nm, True
    If dGeo.Exists(nm) Then
        If dGeo(nm) Then
            names(n) = nm
            n = n + 1
        End If
    End If
    If depth >= 30 Or Not dKids.Exists(nm) Then Exit Sub
    For Each it In dKids(nm)
        If it(1) Or Not LIST_INNER_BODIES Then
            Call SkipBody(CStr(it(0)), dKids, dDone, 0)
        Else
            Call EmitBody(CStr(it(0)), depth + 1, dKids, dGeo, dDone, names, n)
        End If
    Next it
End Sub

' a body that gets no file, and what is inside it
Private Sub SkipBody(ByVal nm As String, dKids As Object, dDone As Object, ByVal depth As Long)
    Dim it As Variant

    If dDone.Exists(nm) Then Exit Sub
    dDone.Add nm, True
    If depth >= 30 Or Not dKids.Exists(nm) Then Exit Sub
    For Each it In dKids(nm)
        Call SkipBody(CStr(it(0)), dKids, dDone, depth + 1)
    Next it
End Sub


' ==============================================================================
' HELPERS: FOLDER, FILE NAMES, INDEX, LOG
' ==============================================================================
Private Function OutputRoot() As String
    If Len(STEP_FOLDER) > 0 Then
        OutputRoot = STEP_FOLDER
    Else
        OutputRoot = CStr(Environ$("TEMP"))
    End If
End Function

' the folder of this run: <root>\<scope>_<date>_<time> - the Windows TEMP
' folder when the root cannot be written
Private Sub StartRunFolder(ByVal scopeName As String)
    Dim root As String, f As Integer, e As Long

    root = OutputRoot()
    Call EnsureFolder(root)
    mFolder = root & "\" & Left$(SafeName(scopeName), 60) & "_" & Format(Now, "yyyymmdd_hhnnss")
    Call EnsureFolder(mFolder)
    On Error Resume Next
    Err.Clear
    f = FreeFile
    Open mFolder & "\step_export_log.txt" For Append As #f
    e = Err.Number
    Close #f
    Err.Clear
    On Error GoTo 0
    If e <> 0 Then
        mFolder = CStr(Environ$("TEMP")) & "\" & Left$(SafeName(scopeName), 60) & "_" & Format(Now, "yyyymmdd_hhnnss")
        Call EnsureFolder(mFolder)
    End If
    mLogPath = mFolder & "\step_export_log.txt"
End Sub

Private Sub EnsureFolder(ByVal p As String)
    Dim parentPath As String
    If Len(p) = 0 Then Exit Sub
    On Error Resume Next
    If mFso.FolderExists(p) Then Exit Sub
    parentPath = mFso.GetParentFolderName(p)
    If Len(parentPath) > 0 And parentPath <> p Then Call EnsureFolder(parentPath)
    mFso.CreateFolder p
    Err.Clear
End Sub

' a file name Windows accepts (and that the log can write)
Private Function SafeName(ByVal s As String) As String
    Dim bad As Variant, k As Long, ch As String

    For Each bad In Array("\", "/", ":", "*", "?", Chr(34), "<", ">", "|")
        s = Replace(s, CStr(bad), "_")
    Next bad
    For k = 0 To 31
        s = Replace(s, Chr(k), "_")
    Next k
    For k = 1 To Len(s)
        ch = Mid$(s, k, 1)
        If StrConv(StrConv(ch, vbFromUnicode), vbUnicode) <> ch Then Mid$(s, k, 1) = "_"
    Next k
    s = Trim$(s)
    Do While Right$(s, 1) = "."
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(s) > 120 Then s = Left$(s, 120)
    If Len(s) = 0 Then s = "unnamed"
    SafeName = s
End Function

' <base>.stp, or <base>~2.stp ... when the name is taken in this run
Private Function UniqueFileName(ByVal base As String) As String
    Dim cand As String, n As Long

    cand = base & "." & EXPORT_FORMAT
    n = 1
    Do While mFileNames.Exists(LCase$(cand))
        n = n + 1
        cand = base & "~" & n & "." & EXPORT_FORMAT
    Loop
    mFileNames.Add LCase$(cand), True
    UniqueFileName = cand
End Function

' one line of the index - description and quantity are looked up when the
' index is written, so pieces found later under a provisional name count too
Private Sub AddRow(ByVal pn As String, ByVal bodyName As String, ByVal fileName As String, _
                   ByVal statusTxt As String)
    mRows.Add Array(pn, bodyName, fileName, statusTxt)
End Sub

' step_index.csv - ";" separated, as German Excel opens it with a double click
Private Sub WriteIndex()
    Dim f As Integer, it As Variant, desc As String, qty As String, place As String

    If Len(mFolder) = 0 Or mRows Is Nothing Then Exit Sub
    On Error Resume Next
    f = FreeFile
    Open mFolder & "\step_index.csv" For Output As #f
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    Print #f, "Part Number;Description;Qty;Body;STEP file;Status;Found in (first place in the tree)"
    For Each it In mRows
        desc = ""
        qty = ""
        place = ""
        If mDesc.Exists(CStr(it(0))) Then desc = CStr(mDesc(CStr(it(0))))
        If mQty.Exists(CStr(it(0))) Then
            qty = CStr(mQty(CStr(it(0))))
        ElseIf mFailed.Exists(CStr(it(0))) Then
            qty = CStr(mFailed(CStr(it(0))))
        End If
        If mPath.Exists(CStr(it(0))) Then place = CStr(mPath(CStr(it(0))))
        Print #f, CsvField(it(0)) & ";" & CsvField(desc) & ";" & qty & ";" & CsvField(it(1)) & ";" & _
                  CsvField(it(2)) & ";" & CsvField(it(3)) & ";" & CsvField(place)
    Next it
    Close #f
    Err.Clear
End Sub

Private Function CsvField(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    If InStr(s, ";") > 0 Or InStr(s, Chr(34)) > 0 Or InStr(s, vbLf) > 0 Then
        s = Chr(34) & Replace(s, Chr(34), Chr(34) & Chr(34)) & Chr(34)
    End If
    CsvField = s
End Function

Private Sub LogLine(ByVal txt As String)
    Dim f As Integer
    If Len(mLogPath) = 0 Then Exit Sub
    On Error Resume Next
    f = FreeFile
    Err.Clear
    Open mLogPath For Append As #f
    If Err.Number <> 0 Then
        mLogPath = ""
    Else
        Print #f, Format(Now, "hh:nn:ss") & "  " & txt
        Close #f
    End If
    Err.Clear
End Sub

Private Sub SetStatus(ByVal txt As String)
    On Error Resume Next
    CATIA.StatusBar = txt
    Err.Clear
End Sub

Private Function InstanceName(oProd As Product) As String
    Dim s As String
    s = ""
    On Error Resume Next
    s = oProd.Name
    Err.Clear
    On Error GoTo 0
    If Len(s) = 0 Then s = "the component"
    InstanceName = s
End Function

Private Function BaseFileName(ByVal s As String) As String
    Dim p As Long
    p = InStrRev(s, ".")
    If p > 1 Then
        BaseFileName = Left$(s, p - 1)
    Else
        BaseFileName = s
    End If
End Function

' seconds since t0 - also across midnight
Private Function Elapsed(ByVal t0 As Double) As Double
    Dim d As Double
    d = Timer - t0
    If d < 0 Then d = d + 86400#
    Elapsed = d
End Function

' a short wait that keeps CATIA alive
Private Sub Pause(ByVal ms As Long)
    Dim t As Double
    t = Timer
    Do
        DoEvents
    Loop While Elapsed(t) * 1000# < ms
End Sub
