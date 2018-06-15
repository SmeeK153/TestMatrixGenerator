VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Issue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private xFields As New Dictionary
''
'' FOR ANY NEW VARIABLES THAT ARE DESIRED (OR THOSE TO BE REMOVED) THE BELOW PROPERTY
'' SETTERS AND GETTERS SHOULD BE ADDED OR REMOVED. KEEP IN MIND THAT EXCEL VBA IS NOT
'' CASE-SENSITIVE SO 'variable', 'Variable' AND 'VARIABLE' ARE ALL THE SAME THING. ALL
'' INTERNAL VARIABLES SHOULD START WITH A LOWERCASE 'X' AND PREFER TO STORE THEM AS A
'' STRING, LONG, INTEGER, DOUBLE, BOOLEAN IF AT ALL POSSIBLE; CUSTOM TYPES ARE NOT
'' DESIRED. SINCE EXCEL IS RATHER LIMITED IN ITS ABILITY TO HANDLE CLASSES, DO NOT
'' EXTEND CLASSES, BUT INSTEAD ADD MORE ATTRIBUTES TO THE 'ISSUE' TO STORE VALUES THAT
'' MAY BE UNIQUE TO CERTAIN ISSUETYPES (AND ENSURE THAT THE FACTORY HAS A CHECK IF THE
'' VALUE EXISTS BEFORE IT TRIES TO SET THE VALUE)
''

Public Sub SetProperty(ByVal Key As String, ByVal Value As Variant)
    If xFields.Exists(Key) Then
        If IsObject(Value) Then
            Set xFields(Key) = Value
        Else
            xFields(Key) = Value
        End If
    Else
        xFields.Add Key, Value
    End If
End Sub

Public Function GetProperty(ByVal Key As String) As Variant
    If xFields.Exists(Key) Then
        If IsObject(xFields(Key)) Then
            Set GetProperty = xFields(Key)
        Else
            GetProperty = xFields(Key)
        End If
    Else
        Set GetProperty = Nothing
    End If
End Function

Public Function Exists(ByVal Key As String) As Boolean
    Exists = xFields.Exists(Key)
End Function
