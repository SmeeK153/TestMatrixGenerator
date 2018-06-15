Attribute VB_Name = "Test_Bed"
Private Sub Test1()
'    Dim testFilePath As String
'    testFilePath = Utilities.GetFileName("TEST", ".txt", Utilities.GetUserFolder("TMG"))
'    Debug.Print testFilePath
    
'    Dim filePath As String: filePath = "this\is\a\test\string.txt"
'    Dim fileTypeIndex As Integer: fileTypeIndex = InStrRev(filePath, ".", , vbBinaryCompare) + 1
'    Dim fileNameIndex As Integer: fileNameIndex = InStrRev(filePath, "\", , vbBinaryCompare) + 1
'    Dim fileType As String: fileType = Mid(filePath, fileTypeIndex)
'    Dim fileName As String: fileName = Mid(filePath, fileNameIndex, fileTypeIndex - fileNameIndex - 1)
'    Dim directoryPath As String: directoryPath = Left(filePath, fileNameIndex - 1)
'

'    Utilities.SaveToFileName "C:\Users\efreeman\TMG\TEST.txt", "content1", False

'    Dim sharepointLoc As String
'    sharepointLoc = "\\tapcas049@SSL\DavWWWRoot\te\Shared Documents\Project Documents\Current Programs\ICODES\ICODES Applications and Services\CE\Conveyance Estimator Test Matrix.xlsm"
'
'    Dim fso As New FileSystemObject
'
'
'    Debug.Print fso.FileExists(sharepointLoc)

'    Dim cases As Dictionary
'    Set cases = JIRAClient.GetJQLIssueList("ISSUETYPE%20IN%20(" & Chr(34) & "Test%20Case" & Chr(34) & "%2C" & Chr(34) & "Test%20Suite" & Chr(34) & ")%20AND%20PROJECT%20%3D%20" & "TARGET")
'
'    Debug.Print JSON.ConvertToJson(cases)
    
    Dim testColl2 As New Collection
    testColl2.Add "Test Col 2"
    
    Dim testColl As New Collection
    testColl.Add "Test 1"
    testColl.Add 1
    testColl.Add testColl2
    
    Dim serializedJSON As String
    serializedJSON = Utilities.ConvertToJSONString(testColl)
    Debug.Print serializedJSON
    
    Dim memJSON As Variant
    Set memJSON = Utilities.ConvertFromJSONString(serializedJSON)
    
End Sub

Sub Test2()
    Dim dateTimeList As Variant
    dateTimeList = Split("09.01.2017 14.45.51", " ")
    
    Dim dateList As Variant
    dateList = Split(dateTimeList(0), ".")
    
    Dim timeList As Variant
    timeList = Split(dateTimeList(1), ".")
    
    Dim dateTime As Date
    dateTime = DateSerial(dateList(2), dateList(0), dateList(1))
    dateTime = DateAdd("h", timeList(0), dateTime)
    dateTime = DateAdd("n", timeList(1), dateTime)
    dateTime = DateAdd("s", timeList(2), dateTime)
    
    Debug.Print Format(dateTime, "dddd, mmm dd, yyyy hh:nn")
End Sub

Sub TestX()
    frmControlPanel.ClearPreviousMatrix
End Sub
