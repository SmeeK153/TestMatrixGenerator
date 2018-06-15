VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmControlPanel 
   Caption         =   "Test Matrix Generator Control Panel"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6555
   OleObjectBlob   =   "frmControlPanel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''POSSIBLE ENHANCEMENTS TO COMMENTS: http://www.contextures.com/xlcomments03.html


Const TEST_SUITE_HEADER_ROW = 3

Const SHARED_TEST_MATRIX_RUN_REPOSITORY_ROOT = "\\tapcas049@SSL\DavWWWRoot\te\Shared Documents\Project Documents\Current Programs\ICODES\ICODES Applications and Services\"
Const ICODES_APPLICATIONS_AND_SERVICES_PATH = "Project Documents\Current Programs\ICODES\ICODES Applications and Services\"
Dim fso As New FileSystemObject

Private Function GetSharepointStorageFile() As Scripting.File
    Dim exportsFolder As Scripting.Folder: Set exportsFolder = GetSharepointStorageLocation()
    Dim fileName As String: fileName = Format(Now(), "MM.DD.YYYY hh.mm.ss") & ".txt"
    
    If exportsFolder Is Nothing Then
        Debug.Print "Error: Could not capture storage location."
        Set GetSharepointStorageFile = Nothing
        Debug.Print "Could not create a TextStream for: " & fileName
        Exit Function
    Else
        exportsFolder.CreateTextFile fileName
        Set GetSharepointStorageFile = fso.GetFile(exportsFolder.Path & "\" & fileName)
        Debug.Print "Created export stream: " & exportsFolder.Path & Format(Now(), "MM.DD.YYYY hh.mm.ss") & ".txt"
    End If
End Function

Private Sub btnAddException_Click()
    Dim excludedProject As Range
    Set excludedProject = Sheet2.Range("A" & 1)
    
    While Len(excludedProject.Text) > 0
        Set excludedProject = Sheet2.Range("A" & excludedProject.Row + 1)
        If excludedProject.Value = Me.cmbProjectSelect.Column(2) Then
            Exit Sub
        End If
    Wend
    
    excludedProject.Value = Me.cmbProjectSelect.Column(2)
    Me.btnClose.SetFocus
    Me.cmbProjectSelect.Text = ""
    Me.btnGenerateMatrix.Visible = False
    Me.btnAddException.Visible = False
    Me.cmbProjectSelect.List = GetExceptionFreeProjectList(Me.cmbProjectSelect.List)
End Sub

Private Function GetExceptionFreeProjectList(originalProjectList As Variant)
    Dim populatedRows As Integer
    ReDim newProjectList(1 To UBound(originalProjectList), 1 To 3) As String
    
    Dim originalIndex As Integer
    populatedRows = 0
    For originalIndex = LBound(originalProjectList) To UBound(originalProjectList)
        If Len(CStr(originalProjectList(originalIndex, 2))) > 0 And Not IsOnExceptionList(CStr(originalProjectList(originalIndex, 2))) Then
            populatedRows = populatedRows + 1
        End If
    Next
    
'    For Each projectKey In originalProjectList
'        If Len(projectKey) > 0 And Not IsOnExceptionList(projectKey) Then
'            populatedRows = populatedRows + 1
'        End If
'    Next
    
    ReDim revisedProjectList(1 To populatedRows, 1 To 3) As String
    Dim revisedIndex As Integer
    revisedIndex = 1
    
    For originalIndex = LBound(originalProjectList) To UBound(originalProjectList)
        If Len(CStr(originalProjectList(originalIndex, 2))) > 0 And Not IsOnExceptionList(CStr(originalProjectList(originalIndex, 2))) Then
            revisedProjectList(revisedIndex, 1) = CStr(originalProjectList(originalIndex, 0))
            revisedProjectList(revisedIndex, 2) = CStr(originalProjectList(originalIndex, 1))
            revisedProjectList(revisedIndex, 3) = CStr(originalProjectList(originalIndex, 2))
            revisedIndex = revisedIndex + 1
        End If
    Next
    GetExceptionFreeProjectList = revisedProjectList
End Function

Private Sub btnGenerateReport_Click()
    If Len(Me.cmbPreviousRun1) = 0 Then
        Exit Sub
    End If
    
    If Len(Me.cmbPreviousRun2) > 0 Then
        ComparedSelectedMatrices
    Else
        btnGenerateMatrix_Click
    End If
    
End Sub

Private Sub DocumentStatusChange(ByVal FromStatus As String, ByVal ToStatus As String, Optional useTotals As Boolean = False)
    If Len(ToStatus) = 0 Then
        ToStatus = "Deleted"
    End If

    Dim toColumn As Integer: toColumn = Utilities.GetColumnByHeader(Sheet1, ToStatus, 5)
    Dim fromRow As Integer: fromRow = Utilities.SearchColumnForValue(Sheet1, FromStatus, 1)
    
    If toColumn > 13 Then
        Stop
    End If
    
    If useTotals Then
        Dim totalsRow As Integer: totalsRow = Utilities.SearchColumnForValue(Sheet1, "Totals", 1)
        Sheet1.Cells(totalsRow, toColumn) = Sheet1.Cells(totalsRow, toColumn) + 1
    End If
    
    If Len(FromStatus) > 0 Then
        Sheet1.Cells(fromRow, toColumn) = Sheet1.Cells(fromRow, toColumn) + 1
    End If
End Sub

Private Sub PrintComparisonReport(FromMatrix As Collection, FromMatrixTitle As Date, ToMatrix As Collection, ToMatrixTitle As Date)
    Sheet1.Cells(3, 1).Value = Format(FromMatrix, "dddd, mmmm dd, yyyy H:nn AM/PM") & "      ->      " & Format(ToMatrixTitle + 1, "dddd, mmmm dd, yyyy H:nn AM/PM")
    
    Dim currentFromIssue As Issue
    Dim currentToIssue As Issue
    
    Dim toIndex As Integer
    Dim fromIndex As Integer
    
    For toIndex = 1 To ToMatrix.Count
        Set currentToIssue = ToMatrix.Item(toIndex)
        Dim ToStatus As String: ToStatus = currentToIssue.GetProperty("Status")("name")
        Dim toID As String: toID = currentToIssue.GetProperty("key")
        Dim FromStatus As String
        
        For fromIndex = 1 To FromMatrix.Count
            Set currentFromIssue = FromMatrix.Item(fromIndex)
            If currentFromIssue.GetProperty("key") = toID Then
                FromStatus = currentFromIssue.GetProperty("Status")("name")
                
                '' Print Status change
                DocumentStatusChange FromStatus, ToStatus, True
                
                '' Print Checklist status change
'                DocumentStatusChange 'fromReview', 'toReview'
                
                '' Print Status vs. Checklist change breakdowns
'                DocumentStatusChange FromStatus, 'review'
'                DocumentStatusChange 'review', ToStatus
                
                FromMatrix.Remove (fromIndex)
                Exit For
            End If
        Next fromIndex
    Next toIndex
    
    '' Print out all of the 'from' test cases that no longer exist in 'to' (they were deleted)
    For fromIndex = 1 To FromMatrix.Count
        Set currentFromIssue = FromMatrix.Item(fromIndex)
        FromStatus = currentFromIssue.GetProperty("Status")("name")
        
        DocumentStatusChange FromStatus, "", True
        
    Next fromIndex
End Sub

Private Sub CompareSelectedMatrixTo(currentMatrix As Collection, currentMatrixDate As Date)
    PrintComparisonReport currentMatrix, currentMatrixDate, GetPreviousExport(Me.cmbPreviousRun1), PrepareMatrixTitle(Replace(Me.cmbPreviousRun1, ".txt", ""))
End Sub

Private Sub ComparedSelectedMatrices()
    CompareSelectedMatrixTo GetPreviousExport(Me.cmbPreviousRun2), PrepareMatrixTitle(Replace(Me.cmbPreviousRun2, ".txt", ""))
End Sub

Private Function PrepareMatrixTitle(MatrixTitle As String) As Date

    Dim dateTimeList As Variant
    dateTimeList = Split("09.01.2017 14.45.51", " ")
    
    Dim dateList As Variant
    dateList = Split(dateTimeList(0), ".")
    
    Dim timeList As Variant
    timeList = Split(dateTimeList(1), ".")
    
    PrepareMatrixTitle = DateSerial(dateList(2), dateList(0), dateList(1))
    PrepareMatrixTitle = DateAdd("h", timeList(0), PrepareMatrixTitle)
    PrepareMatrixTitle = DateAdd("n", timeList(1), PrepareMatrixTitle)
    PrepareMatrixTitle = DateAdd("s", timeList(2), PrepareMatrixTitle)

End Function

Public Function GetPreviousExport(ExportSelector As ComboBox) As Collection
    Set GetPreviousExport = Nothing
    
    
    Dim exportReference As String
    exportReference = ExportSelector
    For Each fileExportIndex In GetSharepointStorageLocation.Files
        Dim fileExport As Scripting.File
        Set fileExport = fileExportIndex
        Debug.Print fileExport.Path
        If fileExport.Name = (exportReference) Then
            Set GetPreviousExport = Factory.CreateNewIssueList(JSON.ParseJson(Utilities.GetContentFromFile(fileExport))("issues"))("Test Case")
            Exit Function
        End If
    Next
End Function

Private Function LoadPreviousMatrixExports() As Variant
    
    Exit Function
    
    Utilities.SetStatusBar "Looking for previous runs..."
    Me.cmbPreviousRun1.Clear
    Me.cmbPreviousRun2.Clear
    
    Dim fileList As Variant
    Set fileList = GetSharepointStorageLocation.Files

    For Each folderFile In GetSharepointStorageLocation.Files
        If folderFile.Type = "Text Document" Then
            Utilities.SetStatusBar "Found: " & folderFile.Name
            Me.cmbPreviousRun1.AddItem folderFile.Name
            Me.cmbPreviousRun2.AddItem folderFile.Name
        End If
    Next
    
    Utilities.ClearStatusBar
End Function

Private Function GetSharepointStorageFolder() As Scripting.Folder
    Set GetSharepointStorageFolder = GetTGSharepointFolder(ICODES_APPLICATIONS_AND_SERVICES_PATH & GetSelectedProjectName)
End Function

Private Function GetSharepointStorageLocation() As Scripting.Folder
    If Len(Me.cmbPreviousRun1.Value) > 0 And Len(Me.cmbPreviousRun2.Value) > 0 Then
        Set GetSharepointStorageLocation = Nothing
        Exit Function
    End If

    Set GetSharepointStorageLocation = GetTGSharepointFolder(ICODES_APPLICATIONS_AND_SERVICES_PATH & GetSelectedProjectName & "\Test Matrix Exports")
'    GetSharepointStorageLocation = SHARED_TEST_MATRIX_RUN_REPOSITORY_ROOT & GetSelectedProjectName & "\Test Matrix Exports"
End Function

Private Function GetSelectedProjectName() As String
'    Select Case Me.cmbProjectSelect.Column(1)
'    Case "TARGET"
'
'
'    Case Default
'        GetSelectedProjectName = ""
'    End Select
Debug.Print "Trying to select project: " & Me.cmbProjectSelect.Column(3)
    GetSelectedProjectName = Me.cmbProjectSelect.Column(2)
End Function

Private Sub RefreshCanGenerateMatrix()
    btnGenerateMatrix.Visible = Me.cmbProjectSelect.ListIndex > -1
End Sub

Private Sub RefreshCanGenerateReport()
    If Len(Me.cmbProjectSelect.Text) = 0 Then
        Me.btnGenerateReport.Visible = False
        Exit Sub
    End If

    Me.btnGenerateReport.Visible = (Me.cmbProjectSelect.ListIndex > -1 And Me.cmbPreviousRun1.ListIndex > 0) Or (Me.cmbPreviousRun1.ListIndex > 0 And Me.cmbPreviousRun2.ListIndex > 0)
End Sub

Private Function CanAddToExceptionList() As Boolean
    
End Function

Private Function CanRemoveFromExceptionList() As Boolean

End Function

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub ClearCmbCheck(cmb As ComboBox)

End Sub

Public Sub ClearPreviousMatrix()
    
    Sheet1.Hyperlinks.Delete
    
    Dim cmt As Comment
    For Each cmt In Sheet1.Comments
        cmt.Delete
    Next cmt
    
    Sheet1.Range(TEST_SUITE_HEADER_ROW & ":" & Sheet1.Rows.Count).ClearContents
    Sheet1.Range(TEST_SUITE_HEADER_ROW & ":" & Sheet1.Rows.Count).VerticalAlignment = xlTop
    Sheet1.Range(TEST_SUITE_HEADER_ROW & ":" & Sheet1.Rows.Count).HorizontalAlignment = xlLeft
    Sheet1.Range(TEST_SUITE_HEADER_ROW & ":" & Sheet1.Rows.Count).RowHeight = 45
    Sheet1.Range(TEST_SUITE_HEADER_ROW & ":" & Sheet1.Rows.Count).WrapText = True
    Sheet1.Cells.Interior.Color = RGB(255, 255, 255)
    Sheet1.Cells.Borders.LineStyle = xlLineStyleNone
    Sheet1.Range(TEST_SUITE_HEADER_ROW & ":" & TEST_SUITE_HEADER_ROW).Interior.Color = RGB(0, 0, 0)
    
    
End Sub

Private Sub btnGenerateMatrix_Click()
'On Error GoTo GenerateMatrixErrorHandler
    Dim startTime As Date: startTime = Now()
    
    ClearPreviousMatrix
    
    Dim selectedProjectKey As String
    selectedProjectKey = Me.cmbProjectSelect.Column(2)
    
    Dim casesSearchStarted As Boolean: casesSearchStarted = False
    Dim cases As Dictionary
    Do While cases Is Nothing
        If Not casesSearchStarted Then
'            Dim storageFileReference As Scripting.File
'            Set storageFileReference = GetSharepointStorageFile
            Set cases = JIRAClient.GetJQLIssueList("ISSUETYPE%20IN%20(" & Chr(34) & "Test%20Case" & Chr(34) & "%2C" & Chr(34) & "Test%20Suite" & Chr(34) & ")%20AND%20PROJECT%20%3D%20" & selectedProjectKey) ', storageFileReference)
            casesSearchStarted = True
        End If
        DoEvents
    Loop
    
    Dim currentColumn As Integer: currentColumn = 1
    Dim currentRow As Integer: currentRow = TEST_SUITE_HEADER_ROW
    
    Dim currentIssue As Issue
    For Each testSuite In cases("Test Suite")
        Set currentIssue = testSuite
        Sheet1.Cells.Hyperlinks.Add Sheet1.Cells(currentRow, currentColumn), "https://jira/jira/browse/" & currentIssue.GetProperty("key"), , currentIssue.GetProperty("Status")("name"), currentIssue.GetProperty("key") & ":" & vbNewLine & currentIssue.GetProperty("Summary")
        currentColumn = currentColumn + 1
    Next
    
    Dim testSuiteCount As Integer: testSuiteCount = currentColumn - 1
    
    Dim testCasesList As Variant
    Set testCasesList = cases("Test Case")
    Utilities.SetStatusBar "Building matrix" & "   ...estimated time to complete < " & Round(cases("Test Case").Count / 600, 0) + 1 & " Seconds"
    
    Dim parentIssue As Issue
    Dim testSuiteIndex As Integer
    Dim JIRAFieldTypes As Dictionary: Set JIRAFieldTypes = JIRAClient.GetIssueFields()
    
    Application.ScreenUpdating = False
    
    Dim testCaseList As Variant
    Set testCaseList = cases("Test Case")
    
    For Each testCase In cases("Test Case")
    DoEvents
        Set currentIssue = testCase
        Set parentIssue = Factory.CreateNewIssue(currentIssue.GetProperty("parent"), JIRAFieldTypes)
        For testSuiteIndex = 1 To testSuiteCount
            currentRow = TEST_SUITE_HEADER_ROW
            currentColumn = 1
            If InStr(1, Sheet1.Cells(TEST_SUITE_HEADER_ROW, testSuiteIndex), parentIssue.GetProperty("key"), vbBinaryCompare) > 0 Then
                Do While Len(Sheet1.Cells(currentRow, testSuiteIndex).Value) > 0
                    currentRow = currentRow + 1
                Loop

                Dim testCaseSummarySplit As Variant
                testCaseSummarySplit = Split(currentIssue.GetProperty("Summary"), ": ")
                
                Sheet1.Cells.Hyperlinks.Add Sheet1.Cells(currentRow, testSuiteIndex), "https://jira/jira/browse/" & currentIssue.GetProperty("key"), , currentIssue.GetProperty("key"), testCaseSummarySplit(UBound(testCaseSummarySplit))
                Sheet1.Range(Sheet1.Cells(currentRow, testSuiteIndex), Sheet1.Cells(currentRow, testSuiteIndex)).AddComment currentIssue.GetProperty("Description of Test Case/Suite")
                
'                Sheet1.Cells(currentRow + 1, testSuiteIndex).RowHeight = 60
'                Sheet1.Cells(currentRow + 1, testSuiteIndex).Value = currentIssue.GetProperty("Description of Test Case/Suite")
                
                With Sheet1.Range(Sheet1.Cells(currentRow, testSuiteIndex), Sheet1.Cells(currentRow, testSuiteIndex))
                    Dim colorRow As Integer
                    colorRow = Utilities.SearchColumnForValue(Sheet2, CStr(currentIssue.GetProperty("Status")("name")), 2)
                    
                    If colorRow > 0 Then
                        .Interior.Color = Sheet2.Range("B" & colorRow).Interior.Color
                    Else
                        .Interior.Color = RGB(128, 128, 128)
                    End If
                    
'                    Select Case currentIssue.GetProperty("Status")("name")
'                    Case "Passed", "Failed", "Blocked", "Ready"
'                        .Interior.Color = RGB(0, 102, 0)
'                    Case "Open"
'                        .Interior.Color = RGB(0, 255, 255)
'                    Case "Pending Review"
'                        .Interior.Color = RGB(102, 0, 102)
'                    Case "Suspended"
'                        .Interior.Color = RGB(255, 0, 0)
'                    Case "Incomplete"
'                        .Interior.Color = RGB(255, 255, 0)
'                    Case Default
'                        .Interior.Color = RGB(128, 128, 128)
'                    End Select
                    With .Borders
                        .LineStyle = xlContinuous
                        .Color = vbBlack
                        .Weight = xlThin
                    End With
                End With
                
                Exit For
            End If
        Next
    Next
    
    Dim shp As Shape
    For Each shp In Sheet1.Shapes
        If Not shp.TopLeftCell.Comment Is Nothing Then
          If shp.AutoShapeType = msoShapeRightTriangle Then
            shp.Delete
          End If
        End If
    Next shp
    
    
    Dim xComment As Comment
    For Each xComment In Sheet1.Comments
        xComment.Shape.TextFrame.AutoSize = True
    Next
    
'    If Len(Me.cmbPreviousRun1) > 0 Then
'        CompareSelectedMatrixTo cases("Test Case"), PrepareMatrixTitle(Replace(storageFileReference.Name, ".txt", ""))
'    End If
    
GenerateMatrixErrorHandler:
    If Err.Number <> 0 Then
        MsgBox "There was an internal error: " & vbNewLine & vbNewLine & Err.Number & ": " & Err.Description, vbOKOnly, "Internal Error"
    End If
    Application.ScreenUpdating = True
    Debug.Print "Total execution took " & Format(DateDiff("s", startTime, Now()) / 60, "###,###,###.00") & " Minute(s)"
    Debug.Print "------------------------------------------------------------" & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine
    Unload Me
End Sub

Private Sub cmbPreviousRun1_Change()
    Dim previous2Option As Boolean
    previous2Option = Len(cmbPreviousRun1) > 0
    lblPreviousRun2.Visible = previous2Option
    cmbPreviousRun2.Visible = previous2Option
    RefreshCanGenerateMatrix
    RefreshCanGenerateReport
    
    Me.cmbPreviousRun2.Clear
    For Each listItem In Me.cmbPreviousRun1.List
        If IsNull(listItem) Then
            Exit For
        End If
    
        If listItem <> Me.cmbPreviousRun1 Then
            Me.cmbPreviousRun2.AddItem listItem
        End If
    Next
    ClearCmbCheck cmbPreviousRun1
End Sub

Private Sub cmbPreviousRun2_Change()
    RefreshCanGenerateMatrix
    RefreshCanGenerateReport
    ClearCmbCheck cmbPreviousRun2
End Sub

Private Sub cmbProjectSelect_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    LoadPreviousMatrixExports
End Sub

Private Sub cmbProjectSelect_Change()
    RefreshCanGenerateMatrix

    If Len(Me.cmbProjectSelect.Text) = 0 Then
        Exit Sub
    End If
    
    Me.btnAddException.Visible = Me.cmbProjectSelect.ListIndex > -1
'    RefreshCanGenerateReport
'    ClearCmbCheck cmbProjectSelect
End Sub

Private Function IsOnExceptionList(ByVal projectKey As String) As Boolean
    Dim excludedProject As Range
    Set excludedProject = Sheet2.Range("A" & 1)
    
    While Len(excludedProject.Text) > 0
        If excludedProject.Text = projectKey Then
            IsOnExceptionList = True
            Exit Function
        End If
        Set excludedProject = Sheet2.Range("A" & excludedProject.Row + 1)
    Wend
    IsOnExceptionList = False
End Function


Public Sub UpdateProjectsList()
    Dim projects As Dictionary
    Set projects = JIRAClient.GetProjects()
    cmbProjectSelect.Clear
    ReDim projectList(0 To projects.Count - 1, 0 To 2) As String
    Dim projectIndex As Integer: projectIndex = 0
    
'    Me.cmbProjectSelect.RowSource = "'Configuration'!$C$2:$C$3"
    For Each projectKey In projects
'        Debug.Print projects(projectKey) & ": " & InStr(1, projects(projectKey), " - ", vbBinaryCompare)
        projectNameStartIndex = InStr(1, projects(projectKey), " - ", vbBinaryCompare)
        If projectNameStartIndex = 0 Then
            projectNameStartIndex = 1
        Else
            projectNameStartIndex = projectNameStartIndex + 3
        End If
        'cmbProjectSelect.AddItem Replace(Mid(projects(projectKey), projectNameStartIndex), " _", "") & " (" & projectKey & ")"
        
        projectList(projectIndex, 0) = Replace(Mid(projects(projectKey), projectNameStartIndex), "_", "")
        
        If InStr(1, projects(projectKey), " - ", vbBinaryCompare) > 1 Then
            projectList(projectIndex, 1) = "(" & Mid(projects(projectKey), 1, InStr(1, projects(projectKey), " - ", vbBinaryCompare) - 1) & ")"
        Else
            projectList(projectIndex, 1) = ""
        End If
        
        projectList(projectIndex, 2) = projectKey
        projectIndex = projectIndex + 1
    Next
    
    Me.cmbProjectSelect.List = GetExceptionFreeProjectList(projectList)
End Sub

Private Sub UserForm_Initialize()
    UpdateProjectsList
    Me.StartUpPosition = 0
'    Me.Top = 0
    
        
'    '' Disable the selective updates until the JSON data payload can be saved, may be running into a string length cap
'    Dim projects As Dictionary
''    If Not Utilities.GetDocumentProperty("Project List Initialized", False) Or _
''        Len(Utilities.GetDocumentProperty("Project List", "")) = 0 Then
'        Set projects = JIRAClient.GetProjects()
'        'Utilities.SetDocumentProperty "Project List", JSON.ConvertToJson(projects)
''        Utilities.SetDocumentProperty "Project List Initialized", True
''    Else
''        Set projects = JSON.ParseJson(Utilities.GetDocumentProperty("Project List", ""))
''    End If
'
'    cmbProjectSelect.Clear
'    Dim projectNameStartIndex As Integer
'
'    ReDim ProjectList(1 To projects.Count + 1, 1 To 2) As String
'    Dim projectIndex As Integer: projectIndex = 1
'
'    For Each projectKey In projects
''        Debug.Print projects(projectKey) & ": " & InStr(1, projects(projectKey), " - ", vbBinaryCompare)
'        projectNameStartIndex = InStr(1, projects(projectKey), " - ", vbBinaryCompare)
'        If projectNameStartIndex = 0 Then
'            projectNameStartIndex = 1
'        Else
'            projectNameStartIndex = projectNameStartIndex + 3
'        End If
'        'cmbProjectSelect.AddItem Replace(Mid(projects(projectKey), projectNameStartIndex), " _", "") & " (" & projectKey & ")"
'
'        ProjectList(projectIndex, 1) = Replace(Mid(projects(projectKey), projectNameStartIndex), " _", "")
'        ProjectList(projectIndex, 2) = projectKey
'        projectIndex = projectIndex + 1
'    Next
'
'    Me.cmbProjectSelect.List = ProjectList
    
'    cmbPreviousRun1.Clear
'
'
'    cmbPreviousRun2.Clear
    
    RefreshCanGenerateMatrix
End Sub
