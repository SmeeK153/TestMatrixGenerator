Attribute VB_Name = "Utilities"
Dim fso As New FileSystemObject
Dim nso As New WshNetwork
Const TEST_GROUP_SHAREPOINT_PATH = "\\tapcas049@SSL\DavWWWRoot\te\Shared Documents"

Const COLLECTION_START = "["
Const COLLECTION_END = "]"
Const DICTIONARY_START = "{"
Const DICTIONARY_END = "}"
Const INDEX_SEPARATOR = ","
Const KEY_VALUE_SEPARATOR = ":"

Public Function DebugMode() As Boolean
    DebugMode = True
End Function

Private Function SpliceObjectFromString(StringToSplice As String, SpliceStart As String, SpliceFinish As String) As Variant
    Dim startIndex As Integer: startIndex = InStr(1, StringToSplice, SpliceStart, vbBinaryCompare)
    Dim finishIndex As Integer: finishIndex = InStr(1, StringToSplice, SpliceFinish, vbBinaryCompare)
    Dim splicedString As Variant
    splicedString = Split(Mid(StringToSplice, startIndex, finishIndex - startIndex), ",", , vbBinaryCompare)
    SpliceObjectFromString = Array(splicedString, Mid(StringToSplice, finishIndex))
End Function

Public Function ConvertFromJSONString(JSON_String As String) As Variant
    Dim stringIndex As Integer
    Dim collectionSplit As Variant
    Dim startIndex As Integer
    Dim finishIndex As Integer
    
    '' Check for indicators of a Collection or Dictionary existing
    If InStr(1, JSON_String, COLLECTION_START, vbBinaryCompare) > 0 Then
        collectionSplit = SpliceObjectFromString(JSON_String, COLLECTION_START, COLLECTION_END)
        
        
    ElseIf InStr(1, JSON_String, DICTIONARY_START, vbBinaryCompare) > 0 Then
        collectionSplit = SpliceObjectFromString(JSON_String, DICTIONARY_START, DICTIONARY_END)
        
        
        Dim JSONCollection As New Collection
        'JSONCollection.Add
        
    Else
        If IsNumeric(JSON_String) Then
            ConvertFromJSONString = CDec(JSON_String)
        ElseIf IsDate(JSON_String) Then
            ConvertFromJSONString = CDate(JSON_String)
        ElseIf Trim(JSON_String) = "True" Or Trim(JSON_String) = "False" Then
            ConvertFromJSONString = CBool(JSON_String)
        Else
            ConvertFromJSONString = CStr(JSON_String)
        End If
    End If
End Function

Public Function ConvertToJSONString(JSON_Object As Variant, Optional existingString As String) As String
    Dim unmarshalledString As String
    
    Select Case TypeName(JSON_Object)
    Case "Dictionary"
        
    Case "Collection"
        Dim JSON_Collection As Collection: Set JSON_Collection = JSON_Object
        Dim collectionIndex As Integer: collectionIndex = 1
        unmarshalledString = unmarshalledString & "["
        For Each index In JSON_Collection
            unmarshalledString = unmarshalledString & ConvertToJSONString(collectionIndex) & ":" & ConvertToJSONString(JSON_Collection(collectionIndex)) & ","
            collectionIndex = collectionIndex + 1
        Next
        unmarshalledString = Left(unmarshalledString, Len(unmarshalledString) - 1)
        unmarshalledString = unmarshalledString & "]"
        ConvertToJSONString = unmarshalledString
    Case "Integer", "Long Integer", "Single", "Double", "Currency", "Date", "Fixed String", "Variable String", "Boolean", "Decimal", "Byte", "String"
        ConvertToJSONString = CStr(JSON_Object)
    Case Default, "Object", "Variant"
        Debug.Print "Could not unmarshall: " & TypeName(JSON_Object)
    End Select
End Function

Public Sub SetDocumentProperty(PropertyName As String, PropertyValue As Variant)
    DocumentRegistry.SetProperty PropertyName, PropertyLocationCustom, PropertyValue
End Sub

Public Function GetDocumentProperty(PropertyName As String, Optional defaultValue As Variant) As Variant
    If DocumentRegistry.PropertyExists(PropertyName, PropertyLocationCustom) Then
        If IsObject(DocumentRegistry.PropertyExists(PropertyName, PropertyLocationCustom)) Then
            Set GetDocumentProperty = DocumentRegistry.GetProperty(PropertyName, PropertyLocationCustom)
        Else
            GetDocumentProperty = DocumentRegistry.GetProperty(PropertyName, PropertyLocationCustom)
        End If
    Else
        If IsMissing(defaultValue) Then
            Set GetDocumentProperty = Nothing
        Else
            If IsObject(defaultValue) Then
                Set GetDocumentProperty = defaultValue
            Else
                GetDocumentProperty = defaultValue
            End If
        End If
    End If
End Function

Public Sub AppendContentToFile(SaveContent As String, ToFile As Scripting.File)
    Open ToFile.Path For Append As #1
    Write #1, SaveContent
    Close #1
End Sub

Public Sub SaveContentToFile(SaveContent As String, ToFile As Scripting.File)
    Open ToFile.Path For Output As #1
    Write #1, SaveContent
    Close #1
End Sub

Public Function GetContentFromFile(FromFile As Scripting.File) As String
    Dim lineContents As String
    Dim fileContents As String
    
    Open FromFile.Path For Input As #2
    Do While Not EOF(2)
        Line Input #2, lineContents
        fileContents = fileContents & lineContents
    Loop
    Close #2
    GetContentFromFile = fileContents
End Function

Public Function GetTGSharepointFolder(folderPath As String) As Scripting.Folder
    ReDim pathArray(1 To 1) As String
    pathArray = Split(folderPath, "\")
    Dim targetFolder As Scripting.Folder
    Set targetFolder = GetTGSharepoint()
    Dim folderIndex As Integer
    
    '' Exit if the first folder isn't 'Shared Documents'
    Dim firstFolder As String: firstFolder = Trim(pathArray(LBound(pathArray)))
    If firstFolder <> "Admin" And _
        firstFolder <> "Company_Shared" And _
        firstFolder <> "Group Programs" And _
        firstFolder <> "ICODES PMO" And _
        firstFolder <> "Performance" And _
        firstFolder <> "Project Documents" And _
        firstFolder <> "Project Tools" And _
        firstFolder <> "SLO Installations and VMs" And _
        firstFolder <> "Training" Then
        Exit Function
    End If
    
    For folderIndex = LBound(pathArray) To UBound(pathArray)
        Dim pathStr As String
        
        '' Exit if the folder name is empty
        If Len(pathArray(folderIndex)) = 0 Then
            Exit Function
        End If
        
        pathStr = targetFolder.Path & "\" & Trim(pathArray(folderIndex))
        If Not fso.FolderExists(pathStr) Then
            Set targetFolder = fso.CreateFolder(pathStr)
        Else
            Set targetFolder = fso.GetFolder(pathStr)
        End If
    Next
    Set GetTGSharepointFolder = targetFolder
End Function

Public Function GetTGSharepoint() As Scripting.Folder
    Set GetTGSharepoint = MapTGSharepoint.RootFolder
End Function

Public Function MapTGSharepoint() As Drive
    Set MapTGSharepoint = MapLocationToDrive(TEST_GROUP_SHAREPOINT_PATH, True)
End Function

Public Sub UnMapTGSharepoint()
    UnMapLocation TEST_GROUP_SHAREPOINT_PATH
End Sub

Public Function MapLocationToDrive(Location As String, Optional Authenticated As Boolean = False) As Drive
    Dim existingDrive As Drive: Set existingDrive = GetDriveForMappedLocation(Location)
    If Not existingDrive Is Nothing Then
        Set MapLocationToDrive = existingDrive
        Exit Function
    End If

    Dim i As Integer
    For i = Asc("Z") To Asc("A") Step -1
        If Not fso.DriveExists(Chr(i)) Then
            If Authenticated Then
                frmLoginForm.CreateAuthenticatedMappedDrive Chr(i), Location
            Else
                nso.MapNetworkDrive Chr(i) & ":", Location
            End If
            Debug.Print "Mapped " & Chr(i) & " to " & Location
            Set MapLocationToDrive = fso.GetDrive(Chr(i))
            Exit Function
        End If
    Next i
End Function

Public Function GetDriveForMappedLocation(ByVal Location As String) As Drive
    
    '' Remove extraneous part of sharepoint path that gets scrubbed after creating the drive path
    Location = Replace(Location, "DavWWWRoot\", "")

    Dim i As Integer
    For i = 1 To nso.EnumNetworkDrives.Count - 1 Step 2
        If nso.EnumNetworkDrives.Item(i) = Location Then
            Debug.Print "Found a match for: " & Location
            Set GetDriveForMappedLocation = fso.GetDrive(Left(nso.EnumNetworkDrives.Item(i - 1), 1))
        Else
            Debug.Print "Failed to match: " & nso.EnumNetworkDrives.Item(i) & " to: " & Location
            Set GetDriveForMappedLocation = Nothing
        End If
    Next
End Function

Public Function LocationIsMappedToDrive(Location As String) As Boolean
    LocationIsMappedToDrive = Not (GetDriveForMappedLocation(Location) Is Nothing)
End Function

Public Sub UnMapLocation(Location As String)
    Dim existingDrive As Drive
    Set existingDrive = GetDriveForMappedLocation(Location)
    
    If Not existingDrive Is Nothing Then
        UnMapDrive existingDrive.DriveLetter
    End If
End Sub

Public Sub UnMapDrive(DriveLetter As String)
    If fso.DriveExists(DriveLetter) Then
        nso.RemoveNetworkDrive DriveLetter, True
    End If
End Sub

Public Function GetUserAppFolderApplicationData(fileName As String) As String
    Dim filePath As String: filePath = GetUserAppDataFolderChild("TestMatrix").Path & "/" & fileName & ".txt"
    
    If fso.FileExists(filePath) Then
        GetUserAppFolderApplicationData = Replace(Utilities.GetContentFromFile(fso.GetFile(filePath)), Chr(34) & Chr(34), Chr(34))
    End If
End Function

Public Sub UpdateUserAppFolderApplicationData(fileName As String, fileContent As String)
    Dim filePath As String: filePath = GetUserAppDataFolderChild("TestMatrix").Path & "\" & fileName & ".txt"
    If Not fso.FileExists(filePath) Then
        fso.CreateTextFile(filePath, True, False).Close
    End If
    Utilities.SaveContentToFile fileContent, fso.GetFile(filePath)
End Sub

Public Function GetUserAppDataFolderChild(folderName As String) As Scripting.Folder
    Dim folderPath As String: folderPath = Environ("AppData") & "\" & folderName

    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    
    Set GetUserAppDataFolderChild = fso.GetFolder(folderPath)
End Function

Public Function GetUserAppDataFolder() As Scripting.Folder
    Set GetUserAppDataFolder = fso.GetFolder(Environ("AppData"))
End Function

Sub SetStatusBar(Content As String)
    Application.DisplayStatusBar = True
    Application.StatusBar = Content
End Sub

Sub ClearStatusBar()
    Application.StatusBar = False
End Sub

Function SystemUsername() As String
    SystemUsername = LCase(Left(Application.username, 1)) & LCase(Split(Application.username, " ")(1))
End Function

Function EncodeBase64(username As String, password As String) As String
Dim Text As String
Text = username & ":" & password
  Dim arrData() As Byte
  arrData = StrConv(Text, vbFromUnicode)

  Dim objXML As New MSXML2.DOMDocument60
  Dim objNode As MSXML2.IXMLDOMElement
  Set objNode = objXML.createElement("b64")

  objNode.DataType = "bin.base64"
  objNode.nodeTypedValue = arrData
  EncodeBase64 = objNode.Text

  Set objNode = Nothing
  Set objXML = Nothing
End Function

Public Sub OpenLocationInBrowser(URL As String)
    ThisWorkbook.FollowHyperlink URL
End Sub

Function GetMaxColumn(WorksheetToSearch As Worksheet, Optional RowToScan As Integer)
    If RowToScan < 1 Then
        RowToScan = 1
    End If
    GetMaxColumn = WorksheetToSearch.Cells(RowToScan, "XFD").End(xlToLeft).Column
End Function

Function GetColumnByHeader(WorksheetToSearch As Worksheet, HeaderName As String, Optional HeaderRow As Integer) As Integer
    If HeaderRow < 1 Then
        HeaderRow = 1
    End If
    
    Dim colIndex As Integer
    For colIndex = 1 To (WorksheetToSearch.Columns.Count)
        If WorksheetToSearch.Cells(HeaderRow, colIndex).Value = HeaderName Then
            GetColumnByHeader = colIndex
            Exit For
        ElseIf Len(WorksheetToSearch.Cells(HeaderRow, colIndex).Value) = 0 Then
            Exit For
        End If
    Next
End Function

Function SearchColumnForValue(WorksheetToSearch As Worksheet, valueToFind As String, Optional columnToSearch As Integer) As Integer
    If columnToSearch < 1 Then
        columnToSearch = 1
    End If
    
    Dim initializedSearch As Boolean
    
    Dim rowIndex As Integer
    For rowIndex = 1 To WorksheetToSearch.Cells(1048576, columnToSearch).End(xlUp).Row
        initializedSearch = initializedSearch Or Len(WorksheetToSearch.Cells(rowIndex, columnToSearch).Value) > 0

        If WorksheetToSearch.Cells(rowIndex, columnToSearch).Value = valueToFind Then
            SearchColumnForValue = rowIndex
            Exit For
        ElseIf Len(WorksheetToSearch.Cells(rowIndex, columnToSearch).Value) = 0 And initializedSearch Then
            Exit For
        End If
    Next
End Function

Function GetFileName(ByVal fileName As String, ByVal FileType As String, ByVal filePath As String) As String
    Dim strSaveDirectory As String: strSaveDirectory = filePath
    Dim strFileName As String: strFileName = ""
    Dim strTestPath As String: strTestPath = ""
    Dim strFileBaseName As String: strFileBaseName = Trim(fileName)
    Dim strFilePath As String: strFilePath = ""
    Dim intFileCounterIndex As Integer: intFileCounterIndex = 1

    ' Check if desired directory exists
    If Dir(strSaveDirectory, vbDirectory) = "" Then
        Debug.Print "Directory: " & strSaveDirectory & "... doesn't current exist, attempting to create it now..."
        MkDir strSaveDirectory
    End If
    
    Debug.Print "Saving to: " & strSaveDirectory

    ' Base file name
    Debug.Print "File Name will contain: " & strFileBaseName
    
    ' Loop until we find a free file number
    Do
        If intFileCounterIndex > 1 Then
            ' Build test path base on current counter exists.
            strTestPath = strSaveDirectory & strFileBaseName & " (" & Trim(Str(intFileCounterIndex)) & ").pdf"
        Else
            ' Build test path base just on base name to see if it exists.
            strTestPath = strSaveDirectory & strFileBaseName & "." & FileType
        End If

        If (Dir(strTestPath) = "") Then
            ' This file path does not currently exist. Use that.
            strFileName = strTestPath
        Else
            ' Increase the counter as we have not found a free file yet.
            intFileCounterIndex = intFileCounterIndex + 1
        End If

    Loop Until strFileName <> ""

    ' Found useable filename
    Debug.Print "Free file name: " & strFileName
    GetFileName = strFileName

End Function



