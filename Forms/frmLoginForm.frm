VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLoginForm 
   Caption         =   "JIRA Login Form"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "frmLoginForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim credentials As Dictionary

Public Sub PreLoadCredentials()
    ClearCredentials
    LoadCredentials
    Debug.Print "Loaded: " & GetUsername() & " with '" & GetPassword() & "'"
End Sub

Public Sub CreateAuthenticatedMappedDrive(DriveLetter As String, Location As String)
    Dim nso As New WshNetwork
    
    If Len(GetUsername) = 0 Or Len(GetPassword) = 0 Then
        LoadCredentials
    End If
    
    nso.MapNetworkDrive DriveLetter & ":", Location, , GetUsername(), GetPassword()

    Debug.Print "Created authenticated drive mapping at: " & DriveLetter
    Debug.Print "          For: " & Location
End Sub

Private Sub SetUsername(JIRAUsername)
    DocumentRegistry.SetProperty "JIRAUsername", PropertyLocationCustom, JIRAUsername
End Sub

Private Function GetUsername() As String
    If DocumentRegistry.PropertyExists("JIRAUsername", PropertyLocationCustom) Then
        GetUsername = DocumentRegistry.GetProperty("JIRAUsername", PropertyLocationCustom)
    Else
        GetUsername = ""
    End If
End Function

Private Sub SetPassword(JIRAPassword As String)
    DocumentRegistry.SetProperty "JIRAPassword", PropertyLocationCustom, JIRAPassword
End Sub

Private Function GetPassword() As String
    If DocumentRegistry.PropertyExists("JIRAPassword", PropertyLocationCustom) Then
        GetPassword = DocumentRegistry.GetProperty("JIRAPassword", PropertyLocationCustom)
    Else
        GetPassword = ""
    End If
End Function

Private Sub LoadCredentials()
    frmLoginForm.txtUsername.Text = Utilities.SystemUsername
    frmLoginForm.Show
End Sub

Public Function ClearCredentials()
    SetUsername ""
    SetPassword ""
End Function

Public Function GetCredentials() As String
    If Len(GetUsername) = 0 Or Len(GetPassword) = 0 Then
        LoadCredentials
    End If
    GetCredentials = "Basic " & Utilities.EncodeBase64(GetUsername(), GetPassword())
    Debug.Print "Created credentials for: " & GetUsername()
End Function

Private Sub btnOK_Click()
    SetUsername (Me.txtUsername.Text)
    SetPassword (Me.txtPassword.Text)
    Me.Hide
End Sub

