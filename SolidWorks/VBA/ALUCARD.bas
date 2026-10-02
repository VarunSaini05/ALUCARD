Option Explicit

'===========================================================
' SOLIDWORKS CUT LIST -> DXF (bulk)
'
' PASS 1 : extract body > convert to sheet metal > flatten > DXF
' PASS 2 : (on confirmation) re-extract a CLEAN body for each failed
'          part and try Insert Bends using candidate edges chosen
'          automatically (works for curved / rolled parts)
' PASS 3 : largest-face DXF, ONLY if that face is planar.
'          Curved parts that still fail are listed for manual work.
'
' FIX APPLIED:
'   - Added AlignViewToTop helper
'   - ExportToDWG2 Alignment argument flipped False -> True
'     (this is the main tilt fix)
'===========================================================

Dim swApp As Object
Dim swMaster As Object
Dim swPart As Object

Dim masterFolder As String
Dim partFolder As String
Dim dxfFolder As String

Dim partCounter As Long
Dim successCounter As Long
Dim failedCounter As Long

Const MAX_FAILED_PARTS As Long = 1000

Dim fPartPath(1 To MAX_FAILED_PARTS) As String
Dim fDxfPath(1 To MAX_FAILED_PARTS) As String
Dim fBaseName(1 To MAX_FAILED_PARTS) As String
Dim fBodyName(1 To MAX_FAILED_PARTS) As String
Dim fResolved(1 To MAX_FAILED_PARTS) As Long   ' 0 = unresolved, 1 = bends, 2 = face fallback

Dim failedCount As Long
Dim retryWasRun As Boolean

'--- Insert Bends settings (same as your recorded macro) ---
Const BEND_RADIUS As Double = 0.002
Const BEND_KFACTOR As Double = 0.4
Const BEND_ALLOWANCE As Double = -1
Const BEND_AUTO_RELIEF As Boolean = True
Const BEND_OFFSET_RATIO As Double = 0.5
Const BEND_DO_FLATTEN As Boolean = True

'--- How many candidate edges to try per failed part ---
Const MAX_EDGE_TRIES As Long = 12


'===========================================================
' MAIN
'===========================================================
Sub main()

    Set swApp = Application.SldWorks
    Set swMaster = swApp.ActiveDoc

    If swMaster Is Nothing Then
        MsgBox "Open the master PART containing the Cut List.", vbCritical, "Cut List Automation"
        Exit Sub
    End If

    If swMaster.GetType <> 1 Then
        MsgBox "The active document must be a PART.", vbCritical, "Cut List Automation"
        Exit Sub
    End If

    Dim masterPath As String
    masterPath = swMaster.GetPathName

    If Len(masterPath) = 0 Then
        MsgBox "Save the master part before running the macro.", vbCritical, "Cut List Automation"
        Exit Sub
    End If

    masterFolder = Left(masterPath, InStrRev(masterPath, "\"))
    partFolder = masterFolder & "EXTRACTED PARTS\"
    dxfFolder = partFolder & "DXF\"

    CreateFolder partFolder
    CreateFolder dxfFolder

    partCounter = 0
    successCounter = 0
    failedCounter = 0
    failedCount = 0
    retryWasRun = False

    UpdateCutList

    Debug.Print ""
    Debug.Print "================ STARTING FIRST PASS ================"

    ProcessAllCutListFeatures
    ActivateMaster

    MsgBox "FIRST PASS COMPLETE" & vbCrLf & vbCrLf & _
           "Total bodies: " & partCounter & vbCrLf & _
           "Successful: " & successCounter & vbCrLf & _
           "Failed: " & failedCounter & vbCrLf & vbCrLf & _
           "DXF files: " & dxfFolder, _
           vbInformation, "FIRST PASS REPORT"

    If failedCount > 0 Then

        Dim ans As VbMsgBoxResult
        ans = MsgBox("Failed parts: " & failedCount & vbCrLf & vbCrLf & _
                     "Run the Insert Bends recovery pass (handles curved / rolled parts)?", _
                     vbYesNo + vbQuestion, "RUN FAILED-PART RECOVERY?")

        If ans = vbYes Then
            retryWasRun = True
            RetryFailedParts
        End If

    End If

    ActivateMaster

    '--- final report ---
    Dim recovered As Long, faceFallback As Long, stillFailed As Long
    Dim failList As String
    Dim i As Long

    For i = 1 To failedCount
        Select Case fResolved(i)
            Case 1: recovered = recovered + 1
            Case 2: faceFallback = faceFallback + 1
            Case Else
                stillFailed = stillFailed + 1
                If stillFailed <= 25 Then failList = failList & "  - " & fBaseName(i) & vbCrLf
        End Select
    Next i

    Dim report As String
    report = "PROCESS COMPLETE" & vbCrLf & vbCrLf & _
             "TOTAL BODIES: " & partCounter & vbCrLf & _
             "FIRST PASS SUCCESS: " & successCounter & vbCrLf & _
             "FIRST PASS FAILED: " & failedCounter & vbCrLf & _
             "RECOVERED WITH INSERT BENDS: " & recovered & vbCrLf & _
             "LARGEST PLANAR FACE FALLBACK: " & faceFallback & vbCrLf & _
             "STILL FAILED: " & IIf(retryWasRun, stillFailed, failedCount) & vbCrLf

    If retryWasRun And stillFailed > 0 Then
        report = report & vbCrLf & "NEEDS MANUAL WORK:" & vbCrLf & failList
    End If

    report = report & vbCrLf & "DXF FILES:" & vbCrLf & dxfFolder

    MsgBox report, vbInformation, "CUT LIST AUTOMATION COMPLETE"

End Sub


Sub ActivateMaster()
    Dim e As Long
    On Error Resume Next
    swMaster.ClearSelection2 True
    swApp.ActivateDoc2 swMaster.GetTitle, False, e
    On Error GoTo 0
End Sub


'===========================================================
' VIEW ALIGNMENT HELPER (added - tilt fix support)
'===========================================================
Sub AlignViewToTop(ByVal part As Object)
    On Error Resume Next
    part.ShowNamedView2 "*Top", 8       ' 8 = swEndCustomView
    part.ViewZoomtofit2
    On Error GoTo 0
    DoEvents
End Sub


'===========================================================
' CUT LIST HANDLING
'===========================================================
Sub UpdateCutList()

    Dim feat As Object
    Dim folder As Object

    Set feat = swMaster.FirstFeature

    Do While Not feat Is Nothing

        On Error Resume Next
        If LCase(feat.GetTypeName2) = "solidbodyfolder" Then
            Set folder = feat.GetSpecificFeature2
            If Not folder Is Nothing Then
                folder.SetAutomaticCutList True
                folder.UpdateCutList
            End If
            On Error GoTo 0
            Exit Do
        End If
        On Error GoTo 0

        Set feat = feat.GetNextFeature
    Loop

    DoEvents

End Sub


Sub ProcessAllCutListFeatures()

    Dim feat As Object
    Set feat = swMaster.FirstFeature

    Do While Not feat Is Nothing
        If IsCutListFeature(feat) Then ProcessCutListFeature feat
        Set feat = feat.GetNextFeature
    Loop

End Sub


Function IsCutListFeature(ByVal feat As Object) As Boolean

    Dim typeName As String

    On Error Resume Next
    typeName = LCase(feat.GetTypeName2)
    On Error GoTo 0

    Select Case typeName
        Case "cutlistfolder", "subweldfolder", "weldmentcutlistfolder"
            IsCutListFeature = True
        Case Else
            IsCutListFeature = False
    End Select

End Function


Sub ProcessCutListFeature(ByVal cutListFeat As Object)

    Dim bodyFolder As Object
    Dim bodies As Variant

    Set bodyFolder = Nothing

    On Error Resume Next
    Set bodyFolder = cutListFeat.GetSpecificFeature2
    On Error GoTo 0

    If bodyFolder Is Nothing Then Exit Sub

    On Error Resume Next
    bodies = bodyFolder.GetBodies
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    If IsEmpty(bodies) Then Exit Sub

    Dim i As Long
    For i = LBound(bodies) To UBound(bodies)
        If Not IsEmpty(bodies(i)) Then ProcessOneBody bodies(i), cutListFeat
    Next i

End Sub


'===========================================================
' FIRST PASS - ONE BODY
'===========================================================
Sub ProcessOneBody(ByVal body As Variant, ByVal cutListFeat As Object)

    partCounter = partCounter + 1

    Dim bodyName As String
    Dim cutListName As String

    On Error Resume Next
    bodyName = body.Name
    cutListName = cutListFeat.Name
    On Error GoTo 0

    If Len(bodyName) = 0 Then bodyName = "BODY_" & partCounter
    If Len(cutListName) = 0 Then cutListName = bodyName

    Dim cleanName As String
    cleanName = CleanFileName(cutListName)
    If Len(cleanName) = 0 Then cleanName = CleanFileName(bodyName)
    If Len(cleanName) = 0 Then cleanName = "PART"

    '--- naming convention ---
    ' Master part name becomes the prefix.
    ' Example master: 26672-1
    ' Extracted parts: 26672-1-001, 26672-1-002, ...
    ' DXF: 26672-1-001-Flat Pattern, 26672-1-002-Flat Pattern, ...
    Dim masterBaseName As String
    Dim baseName As String

    masterBaseName = CleanFileName(GetFileBaseName(swMaster.GetPathName))
    If Len(masterBaseName) = 0 Then masterBaseName = "MASTER"

    baseName = masterBaseName & "-" & Format(partCounter, "000")

    Dim partPath As String
    Dim dxfPath As String
    partPath = partFolder & baseName & ".SLDPRT"
    dxfPath = dxfFolder & baseName & ".DXF"

    Debug.Print ""
    Debug.Print "FIRST PASS BODY " & partCounter & " : " & cutListName & " / " & bodyName

    If Not ExtractBody(bodyName, partPath) Then
        Debug.Print "  EXTRACTION FAILED."
        RegisterFailedPart partPath, dxfPath, baseName, bodyName
        failedCounter = failedCounter + 1
        Exit Sub
    End If

    If FirstPassConvert(partPath, dxfPath) Then
        successCounter = successCounter + 1
        Debug.Print "  OK"
    Else
        RegisterFailedPart partPath, dxfPath, baseName, bodyName
        failedCounter = failedCounter + 1
        Debug.Print "  FAILED - queued for recovery"
    End If

End Sub


Function FirstPassConvert(ByVal partPath As String, ByVal dxfPath As String) As Boolean

    FirstPassConvert = False

    Dim doc As Object
    Set doc = OpenPart(partPath)
    If doc Is Nothing Then Exit Function

    Dim isPlanar As Boolean
    Dim face As Object
    Set face = GetLargestFace(doc, isPlanar)

    If face Is Nothing Then
        ClosePart doc
        Exit Function
    End If

    doc.ClearSelection2 True
    face.Select2 False, 0

    If Not ConvertSelectedFaceToSheetMetal(doc) Then
        ClosePart doc
        Exit Function
    End If

    If Not FlattenPart(doc) Then
        ClosePart doc
        Exit Function
    End If

    Dim se As Long, sw As Long
    doc.EditRebuild3
    doc.Save3 1, se, sw
    DoEvents

    If Not ExportFlatDXF(doc, dxfPath) Then
        ClosePart doc
        Exit Function
    End If

    ClosePart doc
    FirstPassConvert = True

End Function


'===========================================================
' OPEN / CLOSE HELPERS
'===========================================================
Function OpenPart(ByVal path As String) As Object

    Set OpenPart = Nothing

    Dim e As Long, w As Long, ae As Long
    Dim doc As Object

    Set doc = swApp.OpenDoc6(path, 1, 0, "", e, w)
    If doc Is Nothing Then Exit Function

    swApp.ActivateDoc2 doc.GetTitle, False, ae
    Set doc = swApp.ActiveDoc

    doc.EditRebuild3
    DoEvents

    Set OpenPart = doc

End Function


Sub ClosePart(ByVal doc As Object)
    On Error Resume Next
    swApp.CloseDoc doc.GetTitle
    On Error GoTo 0
End Sub


'===========================================================
' EXTRACT BODY
'===========================================================
Function ExtractBody(ByVal bodyName As String, ByVal savePath As String) As Boolean

    ExtractBody = False

    swMaster.ClearSelection2 True

    Dim selected As Boolean

    On Error Resume Next
    selected = swMaster.Extension.SelectByID2(bodyName, "SOLIDBODY", 0#, 0#, 0#, False, 0, Nothing, 0)
    If Err.Number <> 0 Then
        Err.Clear
        selected = False
    End If
    On Error GoTo 0

    If Not selected Then
        Debug.Print "  BODY COULD NOT BE SELECTED: " & bodyName
        Exit Function
    End If

    Dim se As Long, sw As Long
    Dim result As Boolean

    On Error Resume Next
    result = swMaster.SaveToFile3(savePath, 0, 2, False, "", se, sw)
    If Err.Number <> 0 Then
        Err.Clear
        result = False
    End If
    On Error GoTo 0

    swMaster.ClearSelection2 True
    ExtractBody = result

End Function


Sub RegisterFailedPart(ByVal partPath As String, ByVal dxfPath As String, _
                       ByVal baseName As String, ByVal bodyName As String)

    Dim i As Long
    For i = 1 To failedCount
        If StrComp(fPartPath(i), partPath, vbTextCompare) = 0 Then Exit Sub
    Next i

    If failedCount >= MAX_FAILED_PARTS Then
        Debug.Print "FAILED PART STORAGE IS FULL."
        Exit Sub
    End If

    failedCount = failedCount + 1
    fPartPath(failedCount) = partPath
    fDxfPath(failedCount) = dxfPath
    fBaseName(failedCount) = baseName
    fBodyName(failedCount) = bodyName
    fResolved(failedCount) = 0

End Sub


'===========================================================
' FACE / FLAT PATTERN HELPERS
'===========================================================
Function GetLargestFace(ByVal part As Object, ByRef isPlanar As Boolean) As Object

    Set GetLargestFace = Nothing
    isPlanar = False

    Dim bodies As Variant
    On Error Resume Next
    bodies = part.GetBodies2(0, False)
    On Error GoTo 0

    If IsEmpty(bodies) Then Exit Function

    Dim body As Object, face As Object, bestFace As Object, srf As Object
    Dim bestArea As Double, a As Double
    Dim i As Long

    For i = LBound(bodies) To UBound(bodies)

        Set body = bodies(i)

        If Not body Is Nothing Then

            Set face = body.GetFirstFace

            Do While Not face Is Nothing

                a = 0
                On Error Resume Next
                a = face.GetArea
                On Error GoTo 0

                If a > bestArea Then
                    bestArea = a
                    Set bestFace = face
                End If

                Set face = face.GetNextFace
            Loop

        End If

    Next i

    If bestFace Is Nothing Then Exit Function

    On Error Resume Next
    Set srf = bestFace.GetSurface
    If Not srf Is Nothing Then isPlanar = srf.IsPlane
    On Error GoTo 0

    Set GetLargestFace = bestFace

End Function


Function GetFlatPatternFeature(ByVal part As Object) As Object

    Dim feat As Object, lastFp As Object
    Dim t As String, n As String

    Set feat = part.FirstFeature

    Do While Not feat Is Nothing

        t = ""
        n = ""

        On Error Resume Next
        t = LCase(feat.GetTypeName2)
        n = LCase(feat.Name)
        On Error GoTo 0

        If t = "flatpattern" Or InStr(1, n, "flat-pattern") > 0 Then
            Set lastFp = feat
        End If

        Set feat = feat.GetNextFeature
    Loop

    Set GetFlatPatternFeature = lastFp

End Function


Function FindFlatPattern(ByVal part As Object) As Boolean
    FindFlatPattern = Not (GetFlatPatternFeature(part) Is Nothing)
End Function


Function ConvertSelectedFaceToSheetMetal(ByVal part As Object) As Boolean

    ConvertSelectedFaceToSheetMetal = False

    Dim r As Boolean

    On Error Resume Next
    r = part.FeatureManager.InsertConvertToSheetMetal2(0.002, False, True, 0.002, 0.002, 0, 0.5, 1, 0.5, False)
    If Err.Number <> 0 Then
        Debug.Print "  CONVERT ERROR: " & Err.Number & " - " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    part.ClearSelection2 True
    DoEvents
    part.EditRebuild3
    DoEvents

    ConvertSelectedFaceToSheetMetal = FindFlatPattern(part)

End Function


Function FlattenPart(ByVal part As Object) As Boolean

    FlattenPart = False

    Dim fp As Object
    Set fp = GetFlatPatternFeature(part)
    If fp Is Nothing Then Exit Function

    part.ClearSelection2 True
    fp.Select2 False, 0
    part.ClearSelection2 True

    Dim st As Long

    On Error Resume Next
    st = part.SetBendState(2)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    DoEvents
    part.EditRebuild3
    DoEvents

    FlattenPart = True

End Function


'===========================================================
' EXPORT FLAT PATTERN DXF (three ways, first that works wins)
'   TILT FIX: AlignViewToTop + Alignment = True on ExportToDWG2
'===========================================================
Function ExportFlatDXF(ByVal part As Object, ByVal dxfPath As String) As Boolean

    ExportFlatDXF = False

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim modelPath As String
    modelPath = part.GetPathName

    If Len(modelPath) = 0 Then Exit Function

    If fso.FileExists(dxfPath) Then
        On Error Resume Next
        fso.DeleteFile dxfPath, True
        On Error GoTo 0
    End If

    '--- force Top view so the flat pattern is not exported at an angle ---
    AlignViewToTop part

    part.EditRebuild3
    DoEvents

    Dim r As Variant

    '--- method 1: ExportToDWG2, flat pattern action, ALIGNED ---
    On Error Resume Next
    r = part.ExportToDWG2(dxfPath, modelPath, 1, True, Empty, True, False, 1, Empty)
    Err.Clear
    On Error GoTo 0
    DoEvents

    If fso.FileExists(dxfPath) Then
        ExportFlatDXF = True
        Exit Function
    End If

    '--- method 2: select flat pattern + SaveAs3 (your recorded method) ---
    Dim fp As Object
    Set fp = GetFlatPatternFeature(part)

    If Not fp Is Nothing Then

        part.ClearSelection2 True
        fp.Select2 False, 0

        On Error Resume Next
        r = part.SaveAs3(dxfPath, 0, 2)
        Err.Clear
        On Error GoTo 0
        DoEvents

        part.ClearSelection2 True

        If fso.FileExists(dxfPath) Then
            ExportFlatDXF = True
            Exit Function
        End If

    End If

    '--- method 3: ExportFlatPatternView ---
    On Error Resume Next
    r = part.ExportFlatPatternView(dxfPath, 0)
    Err.Clear
    On Error GoTo 0
    DoEvents

    If fso.FileExists(dxfPath) Then ExportFlatDXF = True

End Function


'===========================================================
' SECOND PASS - RETRY FAILED PARTS
'===========================================================
Sub RetryFailedParts()

    Debug.Print ""
    Debug.Print "================ RECOVERY PASS : " & failedCount & " parts ================"

    Dim i As Long
    For i = 1 To failedCount
        RetryOneFailedPart i
        DoEvents
    Next i

End Sub


Sub RetryOneFailedPart(ByVal idx As Long)

    Dim partPath As String, dxfPath As String
    partPath = fPartPath(idx)
    dxfPath = fDxfPath(idx)

    Debug.Print ""
    Debug.Print "RECOVERY " & idx & " : " & fBaseName(idx)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    '--- re-extract a CLEAN body so no half-built sheet-metal chain remains ---
    ActivateMaster

    If fso.FileExists(partPath) Then
        On Error Resume Next
        fso.DeleteFile partPath, True
        On Error GoTo 0
    End If

    If Not ExtractBody(fBodyName(idx), partPath) Then
        Debug.Print "  RE-EXTRACTION FAILED."
        Exit Sub
    End If

    '--- method 2: Insert Bends with automatically chosen edges ---
    Dim k As Long
    Dim rc As Long

    For k = 1 To MAX_EDGE_TRIES

        Set swPart = OpenPart(partPath)
        If swPart Is Nothing Then Exit For

        rc = TryInsertBendsWithEdge(swPart, k, dxfPath)

        If rc = 0 Then

            Dim se As Long, sw As Long
            swPart.ClearSelection2 True
            swPart.EditRebuild3
            swPart.Save3 1, se, sw

            ClosePart swPart
            Set swPart = Nothing

            fResolved(idx) = 1
            Debug.Print "  RECOVERED WITH INSERT BENDS (edge try " & k & ")"
            Exit Sub

        End If

        ClosePart swPart       ' discard changes, next attempt starts clean
        Set swPart = Nothing

        If rc = 1 Then Exit For    ' no more candidate edges

    Next k

    '--- method 3: largest face, only when it is planar ---
    If TryLargestFaceFallback(partPath, dxfPath, fBaseName(idx)) Then
        fResolved(idx) = 2
        Debug.Print "  RECOVERED WITH LARGEST-FACE DXF"
    Else
        Debug.Print "  STILL FAILED - needs manual work"
    End If

End Sub


' Returns 0 = success, 1 = no more candidate edges, 2 = this edge failed
Function TryInsertBendsWithEdge(ByVal part As Object, ByVal k As Long, ByVal dxfPath As String) As Long

    TryInsertBendsWithEdge = 2

    Dim edge As Object
    Set edge = GetCandidateEdge(part, k)

    If edge Is Nothing Then
        TryInsertBendsWithEdge = 1
        Exit Function
    End If

    Debug.Print "  EDGE TRY " & k

    part.ClearSelection2 True

    Dim sel As Boolean
    On Error Resume Next
    sel = edge.Select4(False, Nothing)
    If Err.Number <> 0 Then
        Err.Clear
        sel = False
    End If
    On Error GoTo 0

    If Not sel Then Exit Function

    Dim bendResult As Boolean

    On Error Resume Next
    bendResult = part.InsertBends2(BEND_RADIUS, "", BEND_KFACTOR, BEND_ALLOWANCE, _
                                   BEND_AUTO_RELIEF, BEND_OFFSET_RATIO, BEND_DO_FLATTEN)
    If Err.Number <> 0 Then
        Debug.Print "    InsertBends2 ERROR: " & Err.Number & " - " & Err.Description
        Err.Clear
        bendResult = False
    End If
    On Error GoTo 0

    part.ClearSelection2 True
    DoEvents
    part.EditRebuild3
    DoEvents

    Dim fp As Object
    Set fp = GetFlatPatternFeature(part)
    If fp Is Nothing Then Exit Function

    ' reject flat patterns that rebuilt with an error
    Dim errCode As Long
    Dim isWarn As Boolean
    On Error Resume Next
    errCode = fp.GetErrorCode2(isWarn)
    Err.Clear
    On Error GoTo 0
    If errCode <> 0 And Not isWarn Then Exit Function

    If Not FlattenPart(part) Then Exit Function
    If Not ExportFlatDXF(part, dxfPath) Then Exit Function

    TryInsertBendsWithEdge = 0

End Function


'===========================================================
' CANDIDATE EDGES: straight edges, longest first. Returns the k-th.
'===========================================================
Function GetCandidateEdge(ByVal part As Object, ByVal k As Long) As Object

    Set GetCandidateEdge = Nothing

    Dim bodies As Variant
    On Error Resume Next
    bodies = part.GetBodies2(0, False)
    On Error GoTo 0

    If IsEmpty(bodies) Then Exit Function

    Dim body As Object
    Set body = bodies(LBound(bodies))
    If body Is Nothing Then Exit Function

    Dim edges As Variant
    On Error Resume Next
    edges = body.GetEdges
    On Error GoTo 0

    If IsEmpty(edges) Then Exit Function

    Dim n As Long
    n = UBound(edges) - LBound(edges) + 1
    If n <= 0 Then Exit Function

    Dim lens() As Double
    Dim idx() As Long
    Dim cnt As Long
    ReDim lens(1 To n)
    ReDim idx(1 To n)

    Dim i As Long, L As Double
    Dim edge As Object, crv As Object
    Dim isLine As Boolean

    For i = LBound(edges) To UBound(edges)

        Set edge = Nothing
        Set crv = Nothing
        isLine = False
        L = 0

        On Error Resume Next
        Set edge = edges(i)
        Set crv = edge.GetCurve
        If Not crv Is Nothing Then isLine = crv.isLine
        On Error GoTo 0

        If Not edge Is Nothing And isLine Then
            L = EdgeLength(edge)
            If L > 0 Then
                cnt = cnt + 1
                lens(cnt) = L
                idx(cnt) = i
            End If
        End If

    Next i

    If cnt = 0 Or k > cnt Then Exit Function

    ' selection sort, descending by length, only as far as k
    Dim a As Long, b As Long, best As Long
    Dim tmpL As Double, tmpI As Long

    For a = 1 To k
        best = a
        For b = a + 1 To cnt
            If lens(b) > lens(best) Then best = b
        Next b
        If best <> a Then
            tmpL = lens(a): lens(a) = lens(best): lens(best) = tmpL
            tmpI = idx(a): idx(a) = idx(best): idx(best) = tmpI
        End If
    Next a

    Set GetCandidateEdge = edges(idx(k))

End Function


Function EdgeLength(ByVal edge As Object) As Double

    EdgeLength = 0

    Dim v1 As Object, v2 As Object
    Dim p1 As Variant, p2 As Variant

    On Error GoTo EdgeErr

    Set v1 = edge.GetStartVertex
    Set v2 = edge.GetEndVertex
    If v1 Is Nothing Or v2 Is Nothing Then Exit Function

    p1 = v1.GetPoint
    p2 = v2.GetPoint

    EdgeLength = Sqr((CDbl(p2(0)) - CDbl(p1(0))) ^ 2 + _
                     (CDbl(p2(1)) - CDbl(p1(1))) ^ 2 + _
                     (CDbl(p2(2)) - CDbl(p1(2))) ^ 2)
    Exit Function

EdgeErr:
    Err.Clear
    EdgeLength = 0

End Function


'===========================================================
' THIRD METHOD: LARGEST FACE DXF (planar only)
'   TILT FIX: AlignViewToTop + Alignment = True on ExportToDWG2
'===========================================================
Function TryLargestFaceFallback(ByVal partPath As String, ByVal dxfPath As String, _
                                ByVal baseName As String) As Boolean

    TryLargestFaceFallback = False

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim doc As Object
    Set doc = OpenPart(partPath)
    If doc Is Nothing Then Exit Function

    Dim isPlanar As Boolean
    Dim face As Object
    Set face = GetLargestFace(doc, isPlanar)

    If face Is Nothing Then
        ClosePart doc
        Exit Function
    End If

    If Not isPlanar Then
        ' Largest face is curved: exporting a planar strip would give a WRONG DXF.
        Debug.Print "  LARGEST FACE IS CURVED - fallback skipped (manual work needed)."
        ClosePart doc
        Exit Function
    End If

    If fso.FileExists(dxfPath) Then
        On Error Resume Next
        fso.DeleteFile dxfPath, True
        On Error GoTo 0
    End If

    '--- force Top view before exporting ---
    AlignViewToTop doc

    doc.ClearSelection2 True
    face.Select2 False, 0

    Dim r As Variant
    On Error Resume Next
    r = doc.ExportToDWG2(dxfPath, doc.GetPathName, 0, True, Empty, True, False, 1, Empty)
    Err.Clear
    On Error GoTo 0
    DoEvents

    TryLargestFaceFallback = fso.FileExists(dxfPath)

    ClosePart doc

End Function


'===========================================================
' UTILITIES
'===========================================================
Sub CreateFolder(ByVal folderPath As String)

    On Error Resume Next
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
    On Error GoTo 0

End Sub


Function GetFileBaseName(ByVal filePath As String) As String

    GetFileBaseName = ""

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    GetFileBaseName = fso.GetBaseName(filePath)
    On Error GoTo 0

End Function


Function CleanFileName(ByVal fileName As String) As String

    Dim bad As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        fileName = Replace(fileName, bad(i), "_")
    Next i

    fileName = Replace(fileName, vbCr, "")
    fileName = Replace(fileName, vbLf, "")

    CleanFileName = Trim(fileName)

End Function



