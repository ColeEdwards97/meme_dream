Attribute VB_Name = "MemeDreamV4"
Private oGlobals As New Globals

'
'
'
Public Sub MemeDream()

    Dim sw As New StopWatch
    sw.StartTimer

    'Create meme dream message from template
    Dim template As Outlook.mailItem
    Set template = Application.CreateItemFromTemplate(oGlobals.TemplateDir & oGlobals.TemplateName)
    
    Dim vInspector As Outlook.Inspector
    Set vInspector = template.GetInspector

    Dim wEditor As Object
    Set wEditor = vInspector.WordEditor
    
    'back up memedream
    MemeDreamBackup
    
    'update statistics report
    'MemeDreamStatistics

    'set up the template
    template.Subject = "[Cole's Mid-Day Meme Dream V" & oGlobals.Version & "] #" & CStr(GetCount()) & " " & CStr(Date)
    
    'update chart
    UpdateChart
    
    wEditor.Paragraphs(40).Range.Paste
    
    'display the template
    template.Display
    
    'cleanup
    Set template = Nothing
    Set vInspector = Nothing
    Set wEditor = Nothing
    
    Debug.Print "MemeDream took: " & sw.EndTimer & " milliseconds"

End Sub


'
'
' update the score chart
Sub UpdateChart()

    Dim sw As New StopWatch
    sw.StartTimer


    '
    '
    'Create Excel Application
    Dim oExcelApp As Excel.Application
    Dim oWorkbook As Excel.Workbook
    Dim oWorksheet As Excel.Worksheet
    
    Set oExcelApp = CreateObject("Excel.Application")
    Set oWorkbook = oExcelApp.Workbooks.Open(oGlobals.TrackerDir & oGlobals.TrackerName)
    Set oWorksheet = oWorkbook.Sheets(1)
    
    oWorkbook.Application.Visible = False
    oWorksheet.Application.ScreenUpdating = False
    'oWorkbook.Application.Visible = True
    'oWorksheet.Application.ScreenUpdating = True
    
    'Create dictionary and get voting results
    Dim dictScores As Scripting.Dictionary
    Set dictScores = CalculateScores()
    
    'Get reference to tables and columns
    Dim history As Excel.ListObject
    Dim group As Excel.Shape

    Set history = oWorksheet.ListObjects("history")
    Set group = oWorksheet.Shapes("chart")
     
    'Clear filters & sorts
    history.AutoFilter.ShowAllData
    With history.Sort
        .Header = xlYes
        .SortFields.Clear
    End With
    
    
    'Get the columns we need
    Dim histDB As Range
    Dim rngDateCol As Range
    Dim rngDateDB As Range
    Dim rngMFICol As Range
    Dim rngMFIDB As Range
    Dim rngRAvgDB As Range
    Dim rngFound As Range
    
    Set histDB = history.DataBodyRange
    Set rngDateCol = history.ListColumns("Date").Range
    Set rngDateDB = history.ListColumns("Date").DataBodyRange
    Set rngMFICol = history.ListColumns("MFI").Range
    Set rngMFIDB = history.ListColumns("MFI").DataBodyRange
    Set rngRAvgDB = history.ListColumns("rolling avg").DataBodyRange
    
    'For each date, get the weighted score
    For Each keyDate In dictScores.Keys
        
        'if the entry exists, replace it. otherwise add it to the end
        Set rngFound = rngDateCol.Find(what:=keyDate)
        
        If rngFound Is Nothing Then
            history.ListRows.Add
            Set histDB = history.DataBodyRange
            Set rngDateCol = history.ListColumns("Date").Range
            Set rngDateDB = history.ListColumns("Date").DataBodyRange
            Set rngMFICol = history.ListColumns("MFI").Range
            Set rngMFIDB = history.ListColumns("MFI").DataBodyRange
            Set rngRAvgDB = history.ListColumns("rolling avg").DataBodyRange
            rngDateDB(histDB.Rows.Count, 1).Value = keyDate
            rngMFIDB(histDB.Rows.Count).Value = dictScores.Item(keyDate)
        Else
            rngMFICol(rngFound.Row, 1).Value = dictScores.Item(keyDate)
        End If
        
    Next keyDate
    
    'Sort data by ascending date
    history.DataBodyRange.Sort Key1:=rngDateDB, Order1:=xlAscending, Header:=xlYes
    
    'Calculate rolling average
    rowIdx = 1
    Dim dSum As Double
    Dim dAvg As Double
    dSum = 0
    dAvg = 0
    For Each Row In histDB.Rows
        dSum = dSum + rngMFIDB(rowIdx, 1).Value
        dAvg = dSum / rowIdx
        rngRAvgDB(rowIdx, 1).Value = dAvg
        rowIdx = rowIdx + 1
    Next Row
        
    'Filter data for last 30 days
    histDB.AutoFilter Field:=history.ListColumns("Date").Index, Criteria1:=">" & CLng(Date - 30)
    
    'Copy the chart
    group.Copy
    
    'Close the workbook
    oWorkbook.Close SaveChanges:=True
    oExcelApp.Quit
    
    Set oExcelApp = Nothing
    Set oWorkbook = Nothing
    Set oWorksheet = Nothing
    Set history = Nothing
    Set group = Nothing
    Set dictScores = Nothing
    
    Debug.Print "UpdateChart took: " & sw.EndTimer & " milliseconds"

End Sub

'
'
' Calculate the MFI for every meme dream
Function CalculateScores() As Scripting.Dictionary

    Dim sw As New StopWatch
    sw.StartTimer

    Dim dictScores As Scripting.Dictionary
    Set dictScores = CreateObject("Scripting.Dictionary")
    
    Dim sArchiveDir As String
    Dim sFileType As String
    Dim sSentOn As String
    Dim oFile As Variant
    Dim oMail As Outlook.mailItem
    
    Dim sVotingOptions As String
    
    sArchiveDir = oGlobals.ArchiveDir
    sFileType = "*msg"
    oFile = dir(sArchiveDir & sFileType)
    
    'new (voting responses per date)
    While (oFile <> "")
    
        Set oMail = Application.CreateItemFromTemplate(sArchiveDir & oFile)
        If TypeOf oMail Is Outlook.mailItem Then
        
            'DO NOT COUNT V1 or V2 VOTES!!!
            sVotingOptions = oMail.VotingOptions
            If sVotingOptions = oGlobals.VotingOptions Then
            
                'maybe filter for if it was sent in the last month
                'would need a way to replace it in the excel sheet
            
                sSentOn = VBA.Split(oMail.SentOn, " ")(0)
                
                If CDate(sSentOn) > CDate(Date - 30) Then

                    If Not dictScores.Exists(sSentOn) Then
                        dictScores.Add sSentOn, CalculateMFI(GetVotingResponses(oMail))
                    Else
                        Dim MFI1 As Double
                        Dim MFI2 As Double
                        MFI1 = dictScores.Item(sSentOn)
                        MFI2 = CalculateMFI(GetVotingResponses(oMail))
                        dictScores.Item(sSentOn) = (MFI1 + MFI2) / 2
                    End If
                    
                End If
                
            End If
            
        End If
        oFile = dir
    Wend
    
    Set CalculateScores = dictScores
    Set dictScores = Nothing
    Set oFile = Nothing
    Set oMail = Nothing
    
    Debug.Print "CalculateScores took: " & sw.EndTimer & " milliseconds"

End Function

'
'
' Calculate the Meme Funniness Index
Function CalculateMFI(dictResponses As Scripting.Dictionary) As Double
    
    Dim dMFI As Double
    Dim dScoreSum As Double
    Dim dVoteSum As Double
    
    dMFI = 0
    dScoreSum = 0
    dVoteSum = 0
    
    For Each Key In dictResponses
    
        If Key <> "i don't get it" And Key <> "No Response" Then
            dScoreSum = dScoreSum + ((VotingResponse2Value(Key) - 5) * dictResponses.Item(Key))
            dVoteSum = dVoteSum + dictResponses.Item(Key)
        End If
        
    Next Key
    
    If dVoteSum <> 0 Then
        dMFI = dScoreSum / dVoteSum
    End If
    
    CalculateMFI = dMFI
    
End Function

'
'
' Get the voting responses from an email
Function GetVotingResponses(oMail As Outlook.mailItem) As Scripting.Dictionary
    
    Dim dictResponses As Scripting.Dictionary
    Set dictResponses = CreateObject("Scripting.Dictionary")
    
    Dim varVotingOptions As Variant
    Dim varVotingOption As Variant
    
    Dim varRecipients As Outlook.Recipients
    Dim varRecipient As Outlook.Recipient
    
    ' populate the dictionary with voting options as the keys
    varVotingOptions = VBA.Split(oMail.VotingOptions, ";")
    
    For Each varVotingOption In varVotingOptions
        If Not dictResponses.Exists(varVotingOption) Then
            dictResponses.Add varVotingOption, 0
        End If
    Next varVotingOption
    If Not dictResponses.Exists("No Response") Then
        dictResponses.Add "No Response", 0
    End If
    
    
    ' loop through recipients and check their response
    Set varRecipients = oMail.Recipients
    
    For Each varRecipient In varRecipients
        If varRecipient.TrackingStatus = olTrackingReplied Then
           If dictResponses.Exists(varRecipient.AutoResponse) Then
              dictResponses.Item(varRecipient.AutoResponse) = dictResponses.Item(varRecipient.AutoResponse) + 1
           Else
              dictResponses.Add varRecipient.AutoResponse, 1
           End If
        Else
           dictResponses.Item("No Response") = dictResponses.Item("No Response") + 1
        End If
    Next varRecipient
    
    Set GetVotingResponses = dictResponses
    Set dictResponses = Nothing
    Set varVotingOptions = Nothing
    Set varVotingOption = Nothing
    Set varRecipients = Nothing
    Set varRecipient = Nothing

End Function

'
'
' Backup Meme Dreams
Sub MemeDreamBackup()

    Dim sw As New StopWatch
    sw.StartTimer

    Dim Ns As Outlook.NameSpace
    Set Ns = Application.GetNamespace("MAPI")

    Dim folder As Outlook.folder
    Dim items As Outlook.items
    
    Dim objMail As Outlook.mailItem
    Dim strPath As String
    Dim strName As String
    
    strPath = oGlobals.ArchiveDir
    
    'Select a folder
    Set folder = Ns.GetDefaultFolder(olFolderSentMail).Folders(oGlobals.OutgoingFolderName)
    Set items = folder.items
    
    'Save items in the folder
    For Each olItem In items
        If TypeOf olItem Is Outlook.mailItem Then
            Set objMail = olItem
            strName = objMail.ConversationID
            
            objMail.SaveAs strPath & strName & ".msg", olMSG
            
        End If
    Next olItem
    
    Set folder = Nothing
    Set objMail = Nothing
    
    Debug.Print "MemeDreamBackup took: " & sw.EndTimer & " milliseconds"

End Sub

'
'
' Concatenate Dictionaries
Function ConcatDicts(dict1 As Scripting.Dictionary, dict2 As Scripting.Dictionary) As Scripting.Dictionary

    Dim dictConcat As Scripting.Dictionary
    Set dictConcat = CreateObject("Scripting.Dictionary")
    
    For Each Key In dict1.Keys()
        dictConcat.Item(Key) = dict1.Item(Key)
    Next Key
    
    For Each Key In dict2.Keys()
        If dictConcat.Exists(Key) Then
            dictConcat.Item(Key) = dictConcat.Item(Key) + dict2.Item(Key)
        Else
            dictConcat.Item(Key) = dict1.Item(Key)
        End If
    Next Key
    
    Set ConcatDicts = dictConcat
    Set dictConcat = Nothing

End Function

'
'
' Count number of Meme Dreams
Function GetCount() As Integer

    Dim strDir As String
    Dim strType As String
    Dim file As Variant
    Dim fileCount As Integer

    strDir = oGlobals.ArchiveDir
    strType = "*msg"
    file = dir(strDir & strType)
    fileCount = 0
    
    While (file <> "")
        fileCount = fileCount + 1
        file = dir
    Wend
    
    GetCount = fileCount + 1

End Function

'
'
' Get a recipients true name
Function GetTrueName(sName As String) As String

    Dim sFirst As String
    Dim sLast As String
    Dim sStrSplit() As String

    'name shows as email address...
    If InStr(sName, "@") <> 0 Then
    
        sStrSplit = Split(sName, "(")
        sStrSplit = Split(sStrSplit(1), ".")
        sFirst = sStrSplit(0)
        sLast = Split(sStrSplit(1), "@")(0)

    
    'name is deactivated or something...
    ElseIf InStr(sName, "*") <> 0 Then
        
        sStrSplit = Split(sName, "_")
        sStrSplit = Split(sStrSplit(1), ",")
        sLast = sStrSplit(0)
        sFirst = Split(sStrSplit(1), " ")(1)
    
    'name is normal
    Else
    
        sStrSplit = Split(sName, ",")
        sLast = sStrSplit(0)
        sFirst = Split(sStrSplit(1), " ")(1)
        
    End If
    
    GetTrueName = UCase(sLast + ", " + sFirst)

End Function


'
'
' convert a voting option to a numeric value
Function VotingResponse2Value(varVotingOption As Variant) As Integer

    If VBA.IsNumeric(CInt(VBA.Split(varVotingOption, " - ")(0))) Then
        VotingResponse2Value = CInt(VBA.Split(varVotingOption, " - ")(0))
    Else
        VotingResponse2Value = -1
    End If

End Function



'
' STATISTICS REPORT
'



'
'
'
Sub StatisticsReport()

    MemeDreamBackup

    Dim dictRecipients As Scripting.Dictionary
    Set dictRecipients = CreateObject("Scripting.Dictionary")
    
    'Get directory of archived items
    Dim strDir As String
    Dim strType As String
    Dim file As Variant
    Dim oMail As Outlook.mailItem
    
    Dim sVotingOptions As String

    strDir = oGlobals.ArchiveDir
    strType = "*msg"
    file = dir(strDir & strType)
    
    'Loop through archived items
    While (file <> "")
        Set oMail = Application.CreateItemFromTemplate(strDir & file)
        If TypeOf oMail Is Outlook.mailItem Then
        
            'DO NOT COUNT V1 or V2 VOTES!!!
            sVotingOptions = oMail.VotingOptions
            If sVotingOptions Like oGlobals.VotingOptions Then
        
                Set dictRecipients = GetStats(oMail, dictRecipients)
            
            End If
            
        End If
        file = dir
    Wend
    
    
    UpdateReport dictRecipients
        
    MsgBox "Report updated successfully!", vbOKOnly, "MDReport"

End Sub


'
'
'Gets the amount of times a recipient has voted for a certain option
Function GetStats(objMail As Outlook.mailItem, objRecipients As Scripting.Dictionary) As Scripting.Dictionary

    Dim objRecipient As Outlook.Recipient
    
    Dim objMemeDreamRecip As MemeDreamRecipient
    
    Dim varVotingCounts As Variant
    Dim varVotingOptions As Variant
    Dim varVotingOption As Variant
    
    Dim sTrueName As String
    
    'loop through each recipient
    For Each objRecipient In objMail.Recipients
    
        sTrueName = GetTrueName(objRecipient.Name)
    
        'if they are a new recipient, add them to the dict and create a voting tally dict for them
        If Not objRecipients.Exists(sTrueName) Then
            Set objMemeDreamRecip = New MemeDreamRecipient
            objMemeDreamRecip.Create objRecipient
            objRecipients.Add sTrueName, objMemeDreamRecip
        End If
        
        
        'set this recipient as the current one
        Set objMemeDreamRecip = objRecipients.Item(sTrueName)

        
        'go through the recipients voting tally dict and add possible responses to it
        varVotingOptions = Split(objMail.VotingOptions, ";")
        
        For Each varVotingOption In varVotingOptions
            If Not objMemeDreamRecip.Votes.Exists(varVotingOption) Then
                objMemeDreamRecip.Votes.Add varVotingOption, 0
            End If
        Next
        If Not objMemeDreamRecip.Votes.Exists("No Response") Then
            objMemeDreamRecip.Votes.Add "No Response", 0
        End If
        
        
        'get this recipients response and add it to their dictionary
        If objRecipient.TrackingStatus = olTrackingReplied Then
            If objMemeDreamRecip.Votes.Exists(objRecipient.AutoResponse) Then
                objMemeDreamRecip.Votes.Item(objRecipient.AutoResponse) = objMemeDreamRecip.Votes.Item(objRecipient.AutoResponse) + 1
            End If
        Else
            objMemeDreamRecip.Votes.Item("No Response") = objMemeDreamRecip.Votes.Item("No Response") + 1
        End If
        
        
        'update the recipient dictionary
        Set objRecipients.Item(sTrueName) = objMemeDreamRecip
        
    Next
    
    Set GetStats = objRecipients
    Set objRecipients = Nothing
    Set objRecipient = Nothing
    Set objMemeDreamRecip = Nothing
    Set varVotingCounts = Nothing
    Set varVotingOptions = Nothing
    Set varVotingOption = Nothing
    

End Function


'
'
'
Function UpdateReport(dictRecipients As Scripting.Dictionary)


    '
    '
    'Create Excel Application
    Dim oExcelApp As Excel.Application
    Dim oWorkbook As Excel.Workbook
    Dim oWorksheet As Excel.Worksheet
    
    Set oExcelApp = CreateObject("Excel.Application")
    Set oWorkbook = oExcelApp.Workbooks.Open(oGlobals.ReportDir & oGlobals.ReportName)
    Set oWorksheet = oWorkbook.Sheets(1)
    
    oWorkbook.Application.Visible = False
    oWorksheet.Application.ScreenUpdating = False
    'oWorkbook.Application.Visible = True
    'oWorksheet.Application.ScreenUpdating = True

    Dim stats As Excel.ListObject
    Set stats = oWorksheet.ListObjects("stats")
    
    
    'Put data into the table
    Dim mdRecipient As MemeDreamRecipient
    Dim dictVotes As Scripting.Dictionary
    Dim rowIdx As Integer
    Dim colIdx As Integer
    rowIdx = 1
    For Each Key In dictRecipients.Keys
        
        If stats.DataBodyRange Is Nothing Then
            stats.ListRows.Add
        End If
        
        stats.ListColumns("Recipient").DataBodyRange(rowIdx, 1).Value = Key
            
        Set mdRecipient = dictRecipients.Item(Key)
        Set dictVotes = mdRecipient.Votes

        colIdx = 2
        For Each Key2 In dictVotes.Keys
            
            stats.ListColumns(colIdx).DataBodyRange(rowIdx, 1).Value = dictVotes.Item(Key2)
            colIdx = colIdx + 1
        
        Next Key2
            
        rowIdx = rowIdx + 1
        
    Next Key
    
    
    'Close the workbook
    oWorkbook.Close SaveChanges:=True
    oExcelApp.Quit
    
    Set oExcelApp = Nothing
    Set oWorkbook = Nothing
    Set oWorksheet = Nothing
    Set stats = Nothing
    Set mdRecipient = Nothing
    Set dictVotes = Nothing
    
    
End Function




