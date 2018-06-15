Attribute VB_Name = "JIRAClient"
Const MAX_JIRA_RESULTS = 2000
Const JIRA_REST_SEARCH_URL = "https://jira/jira/rest/api/latest/search?jql="

Public Sub OpenIssueInBrowser(issueToOpen As Issue)
    Utilities.OpenLocationInBrowser ("https://jira/jira/browse/" & issueToOpen.GetProperty("key"))
End Sub

Public Function GetProjects() As Dictionary
    Set GetProjects = New Dictionary
    
'    If Len(Utilities.GetUserAppFolderApplicationData("Projects")) = 0 Then
'        LoadProjects
'    End If
    
    For Each project In JSON.ParseJson(JIRAServerResource("GET", "https://jira/jira/rest/api/latest/project", Nothing))
        GetProjects.Add project("key"), project("name")
    Next
End Function

'Public Sub LoadProjects()
''    Utilities.UpdateUserAppFolderApplicationData "Projects", JIRAServerResource("GET", "https://jira/jira/rest/api/latest/project")
'    Utilities.SetDocumentProperty "Projects", JIRAServerResource("GET", "https://jira/jira/rest/api/latest/project")
'End Sub

Public Function GetIssueFields() As Dictionary
    Set GetIssueFields = New Dictionary
    Dim issueFieldCollection As Collection
    
'    If Len(Utilities.GetDocumentProperty("Issue Fields", "")) = 0 Then
'        LoadIssueFields
'    End If
    
    Set issueFieldCollection = JSON.ParseJson(JIRAServerResource("GET", "https://jira/jira/rest/api/latest/field", Nothing))
    
    Dim fieldIndex As Integer
    For fieldIndex = 1 To issueFieldCollection.Count
        GetIssueFields.Add issueFieldCollection(fieldIndex)("id"), issueFieldCollection(fieldIndex)("name")
    Next
End Function

'Public Sub LoadIssueFields()
'    Utilities.SetDocumentProperty "Issue Fields", JIRAServerResource("GET", "https://jira/jira/rest/api/latest/field", Nothing)
'End Sub

''
''  Provides the means to access a given JIRA Filter by ID to capture the results and use them in Excel.
''
''  Parameters:
''      Long - Filter ID to be searched
''
''  Return: Dictionary<Issues> of all the issues returned by the filter
''
Public Function GetFilterIssueList(FilterID As Long, Optional SaveQuery As Scripting.File) As Dictionary
    Set GetFilterIssueList = GetJQLIssueList("filter=" & FilterID, SaveQuery)
End Function

''
''  Provides the means to retrieve JIRA issues by string search to capture the results and use them in Excel.
''
''  Parameters:
''      String - Search string to be executed
''
''  Return: Dictionary<Issues> of all the issues returned by the filter
''
Public Function GetJQLIssueList(JQL As String, Optional SaveQuery As Scripting.File) As Dictionary
    Set GetJQLIssueList = Factory.CreateNewIssueList( _
        JSON.ParseJson( _
            SearchResults(JQL, SaveQuery) _
        )("issues") _
    )
End Function

''
''  Provides the means to retrieve JIRA search result.
''
''  Parameters:
''      String - Search string to be executed
''
''  Return: String of the raw JSON returned by the server
''
Private Function SearchResults(JQL As String, Optional SaveQuery As Scripting.File) As String
    SearchResults = JIRAServerResource("GET", JIRA_REST_SEARCH_URL & JQL & "&expand=all&maxResults=" & MAX_JIRA_RESULTS, SaveQuery)
End Function


''
''  Provides the means to access the JIRA Rest API with the 'Basic' Authentication protocol.
''
''  Parameters:
''      String - Request method to be used
''          Supported RFC 7231: Section 4 Method(s): GET, HEAD, POST, PUT, DELETE, CONNECT, OPTIONS, TRACE
''          Supported RFC 5789: Section 2 Method(s): PATCH
''      String - JIRA Rest Resource
''
''  Please Note: This library does NOT currently support data payloads for the JIRA REST API which is
''               required for some REST API calls.
''
''  Return: String of the raw JSON returned by the server
''
Public Function JIRAServerResource(RequestMethod As String, JIRAResource As String, Optional SaveQuery As Scripting.File) As String
On Error GoTo JIRAErrorHandler

    '' Remove any spaces in the request resource string
    JIRAResource = Replace(JIRAResource, " ", "")
    
    Dim xml As New MSXML2.ServerXMLHTTP60
    
    xml.Open RequestMethod, JIRAResource, False
    '' Load the credentials and then clear the variable reference(s)
    Dim credentials As Dictionary
    Utilities.SetStatusBar "Requesting JIRA credentials..."
    xml.setRequestHeader "Authorization", frmLoginForm.GetCredentials()
    Debug.Print RequestMethod & " " & JIRAResource
    
    Utilities.SetStatusBar "Connecting to JIRA..."
    Do While xml.ReadyState <> 4
        If xml.ReadyState < 2 Then
            xml.Send
        End If
        DoEvents
    Loop
    Utilities.SetStatusBar "Retrieving JIRA response" & "   ...estimated time to complete < " & 2 & " Minutes"
    Dim responseContent As String
    responseContent = xml.responseText
    
    JIRAServerResource = responseContent
    Utilities.SetStatusBar "Processing response..."
    
    Dim responseObject As Dictionary
    Select Case Int(xml.Status)
    Case 401
        MsgBox "Either submitted username and/or password was incorrect." & vbNewLine _
        & vbNewLine & "Please try logging in with your browser to avoid getting locked out.", vbOKOnly, "Login Failed"
        frmLoginForm.ClearCredentials
        Utilities.OpenLocationInBrowser "https://jira/jira"
        Utilities.ClearStatusBar
        End
    Case 400
        Set responseObject = JSON.ParseJson(xml.responseText)
        If responseObject.Exists("errorMessages") Then

            MsgBox "There was a problem processing the request: " & vbNewLine & vbNewLine & responseObject("errorMessages")(1) & vbNewLine _
            & vbNewLine & "Please correct the request or speak with an appropriate POC to resolve this issue.", vbOKOnly, "Request Failed"
        End If
        Utilities.ClearStatusBar
        End
    End Select
    
    Set credentials = Nothing
    Debug.Print xml.Status
    Debug.Print "Server connection closed."
'    Utilities.SetStatusBar "JIRA issues downloaded."
    If Not (SaveQuery Is Nothing) Then
        Utilities.SetStatusBar "Saving response to: " & SaveQuery.Path & "   ...estimated time to complete < " & Round(Len(responseContent) / (100000 * 60), 0) + 1 & " Minutes"
        Open SaveQuery.Path For Output As #1
        Print #1, responseContent
        Close #1
        Utilities.SetStatusBar "Response saved."
    End If
    Utilities.ClearStatusBar
JIRAErrorHandler:
    Select Case Err.Number
    Case -2147012889
        MsgBox "Could not connect to JIRA." & vbNewLine & vbNewLine & "Please check that you are on the network and JIRA is available.", vbOKOnly, "Connection Failed"
        End
    Case Else
        If Err.Number <> 0 Then
            MsgBox "There was an internal error: " & vbNewLine & vbNewLine & Err.Number & ": " & Err.Description, vbOKOnly, "Internal Error"
        End If
    End Select
End Function

