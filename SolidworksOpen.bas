Option Explicit

' SolidworksOpen.bas - SolidWorks VBA macro
'
' Prompts for a job number, locates the SolidWorks job folder (same
' folder hierarchy Pack-n-Go uses), and opens the job's drawing:
' <jobNum>-01.SLDDRW, then <jobNum>-02.SLDDRW. If neither drawing
' exists it falls back to the assembly: <jobNum>-01.SLDASM, then
' <jobNum>-02.SLDASM.

Private Const SW_ROOT As String = "Z:\Solidworks\Current\JOBS\"

' Range folder name based on first 3 digits, in groups of 5.
' Special case: 401-405 is rolled into "400-405".
Private Function ComputeRangeFolder(jobNum As String) As String
    Dim prefix As Long: prefix = CLng(Left$(jobNum, 3))
    Dim n As Long:      n = -Int(-prefix / 5)               ' ceil(prefix / 5)
    Dim startN As Long: startN = 5 * (n - 1) + 1
    Dim endN As Long:   endN = 5 * n
    If startN = 401 And endN = 405 Then
        ComputeRangeFolder = "400-405"
    Else
        ComputeRangeFolder = startN & "-" & endN
    End If
End Function

' SolidWorks intermediate: HD-PFD lives in a "NNXXXX" bucket keyed on the
' first two digits of the job number (40XXXX, 41XXXX, ... 49XXXX);
' AXIAL has no intermediate at all (jobs sit directly under AXIAL\);
' everyone else uses a range folder on the first 3 digits.
Private Function ComputeSwIntermediate(swType As String, jobNum As String) As String
    Select Case UCase$(swType)
        Case "HD-PFD": ComputeSwIntermediate = Left$(jobNum, 2) & "XXXX"
        Case "AXIAL":  ComputeSwIntermediate = ""
        Case Else:     ComputeSwIntermediate = ComputeRangeFolder(jobNum)
    End Select
End Function

Private Function FolderExists(p As String) As Boolean
    On Error Resume Next
    FolderExists = (Len(Dir$(p, vbDirectory)) > 0)
    On Error GoTo 0
End Function

Private Function FileExists(p As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(p)) > 0)
    On Error GoTo 0
End Function

' Probes every SolidWorks job-type folder; returns the type that contains
' <jobNum> and writes the matching SW job folder path to swJobFolder.
Private Function FindSwJobFolder(jobNum As String, ByRef swJobFolder As String) As String
    Dim swTypes As Variant
    swTypes = Array("GENERAL LINE", "HD-PFD", "HDX", "AXIAL")
    Dim i As Long, candidate As String, intermediate As String
    For i = LBound(swTypes) To UBound(swTypes)
        intermediate = ComputeSwIntermediate(CStr(swTypes(i)), jobNum)
        If Len(intermediate) > 0 Then intermediate = intermediate & "\"
        candidate = SW_ROOT & swTypes(i) & "\" & intermediate & jobNum & "\"
        If FolderExists(candidate) Then
            FindSwJobFolder = CStr(swTypes(i))
            swJobFolder = candidate
            Exit Function
        End If
    Next i
    FindSwJobFolder = ""
End Function

' Looks for <jobNum>-01 / <jobNum>-02 in the job folder: every drawing
' first, then assemblies only if no drawing exists. Writes the matching
' document type to docType and returns the full path ("" if none).
Private Function FindJobFile(jobFolder As String, jobNum As String, _
                             ByRef docType As swDocumentTypes_e) As String
    Dim suffixes As Variant: suffixes = Array("-01", "-02")
    Dim i As Long, cand As String
    For i = LBound(suffixes) To UBound(suffixes)
        cand = jobFolder & jobNum & suffixes(i) & ".SLDDRW"
        If FileExists(cand) Then
            docType = swDocDRAWING
            FindJobFile = cand
            Exit Function
        End If
    Next i
    For i = LBound(suffixes) To UBound(suffixes)
        cand = jobFolder & jobNum & suffixes(i) & ".SLDASM"
        If FileExists(cand) Then
            docType = swDocASSEMBLY
            FindJobFile = cand
            Exit Function
        End If
    Next i
    FindJobFile = ""
End Function

Public Sub main()
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    Dim jobNum As String
    jobNum = Trim$(InputBox("Enter job number:", "Solidworks-Open"))
    If Len(jobNum) = 0 Then Exit Sub
    If Not IsNumeric(jobNum) Or Len(jobNum) < 3 Then
        MsgBox "Job number must be numeric and at least 3 digits.", vbExclamation
        Exit Sub
    End If

    Dim swJobFolder As String, swType As String
    swType = FindSwJobFolder(jobNum, swJobFolder)
    If Len(swType) = 0 Then
        MsgBox "No SolidWorks job folder found for job " & jobNum & "." & vbCrLf & _
               "Searched all type folders under " & SW_ROOT, vbExclamation
        Exit Sub
    End If

    Dim docType As swDocumentTypes_e
    Dim filePath As String
    filePath = FindJobFile(swJobFolder, jobNum, docType)
    If Len(filePath) = 0 Then
        MsgBox "No drawing or assembly named " & jobNum & "-01 or " & jobNum & "-02 found in:" & vbCrLf & _
               swJobFolder, vbExclamation
        Exit Sub
    End If

    Dim errors As Long, warnings As Long
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.OpenDoc6(filePath, docType, swOpenDocOptions_Silent, "", errors, warnings)

    ' OpenDoc6 can come back empty when the document is already open in
    ' this session - fall back to the loaded copy before giving up.
    If swModel Is Nothing Then Set swModel = swApp.GetOpenDocumentByName(filePath)
    If swModel Is Nothing Then
        MsgBox "Failed to open:" & vbCrLf & filePath, vbExclamation
        Exit Sub
    End If

    ' Bring it to the front (covers the already-open case, where OpenDoc6
    ' does not activate the document).
    Dim actErrors As Long
    swApp.ActivateDoc3 swModel.GetPathName, False, swDontRebuildActiveDoc, actErrors
End Sub
