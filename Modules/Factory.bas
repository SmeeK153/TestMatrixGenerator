Attribute VB_Name = "Factory"
''
''  Provides an initialization method to create an Issue instance
''      (Since Excel doesn't support parameterized constructors, or custom constructors)
''
''  Parameters:
''      String - JSON Payload of this issue (Expected to be in JIRA format)
''
''  Return: Issue instance
''
Public Function CreateNewIssue(ByVal JSON_Payload As Dictionary, ByVal JIRAFields As Dictionary) As Issue
    Set CreateNewIssue = New Issue
    
    '' Save the JIRA ID
    CreateNewIssue.SetProperty "key", JSON_Payload("key")
    
    '' Load the remaining fields (replace custom id's with real names)
    Dim fieldList As Dictionary: Set fieldList = JSON_Payload("fields")
    For Each fieldKey In fieldList.Keys()
        
        '' Some fields may not have a mapped name, so replacing the field name with the mapped name may wipe-out the value
        If Len(Trim(JIRAFields(fieldKey))) = 0 Then
            JIRAFields(fieldKey) = fieldKey
        End If

        CreateNewIssue.SetProperty JIRAFields(fieldKey), fieldList(fieldKey)
    Next
End Function

''
''  Provides a means to create a Dictionary of Issue instances
''      (Since Excel doesn't support parameterized constructors, or custom constructors)
''
''  Parameters:
''      String - JSON Payload of an issue list (Expected to be in JIRA format)
''
''  Return: Dictionary<String(JIRA ID),Issue> instance
''
Public Function CreateNewIssueList(JSON_Payload As Collection) As Dictionary
    Dim fieldList As Dictionary: Set fieldList = JIRAClient.GetIssueFields()
    Dim newIssueList As New Dictionary
    Dim jsonIndex As Integer
    
    Utilities.SetStatusBar "Processing content..."
    
    '' Add sorting by issue type
    For Each jsonIssue In JSON_Payload
    DoEvents
        Dim newIssue As Issue: Set newIssue = Factory.CreateNewIssue(jsonIssue, fieldList)
        
        '' Create new list for storing different issuetypes
        Dim IssueType As Variant
        IssueType = newIssue.GetProperty("Issue Type")("name")

        If Not newIssueList.Exists(IssueType) Then
            newIssueList.Add IssueType, New Collection
        End If
        
        newIssueList(IssueType).Add newIssue
'        Debug.Print "Logging issue: " & newIssue.GetProperty("key")
        Utilities.SetStatusBar "Found issue: " & newIssue.GetProperty("key")
    Next
    Set CreateNewIssueList = newIssueList
    Utilities.SetStatusBar "Finished processing content."
End Function


