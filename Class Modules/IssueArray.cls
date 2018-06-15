VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IssueArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private xIssueLists As Dictionary

Public Sub SetList(ByVal IssueInstance As Issue)
    If xIssueLists.Exists(IssueInstance.GetProperty("Issue Type")("name")) Then
        Set xFields(IssueType) = IssueInstance
    Else
        xIssueLists.Add IssueType, IssueInstance
    End If
End Sub

Public Function GetList(ByVal IssueType As String) As Issue
    If xIssueLists.Exists(IssueType) Then
        Set GetList = xIssueLists(IssueType)
    Else
        Set GetList = Nothing
    End If
End Function

Public Function Exists(ByVal IssueType As Variant) As Boolean
    Dim searchForList As String
    
    Select Case TypeName(IssueType)
    Case "String"
        searchForList = IssueType
    Case "Issue"
        searchForList = IssueType.GetProperty("Issue Type")("name")
    End Select
    Exists = xIssueLists.Exists(searchForList)
End Function

