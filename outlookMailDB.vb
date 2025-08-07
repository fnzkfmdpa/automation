Public olApp As Object
Public olNamespace As Object
Public olFolder As Object
Public olItem As Object
Public olFolderPath As String

Private Sub GetEmailDataFromOutlook()
    Dim xlSheet As Worksheet
    Dim i As Integer

    Dim newWorksheet As Worksheet
    Dim sheetExists As Boolean
    Dim lastSheet As Worksheet
    Dim pcusername As String
    
    Dim valueToCopy As Variant
    Dim valueToCopy2 As Variant
    Dim currentDate As String

    Call checkAccount.Users
    Call 현재_통합_문서.CopyAndRenameSheet

    Dim users As Collections
    users = checkAccount.users()

    getOngoinRequets(users)
    getCompletedRequests(users)
    compareReqeustStatus()
End Sub

'Function getOngoingRequests(ByRef users As Collections)
Function getOngoingRequests(users)
    ''create Outlook object
    Set olApp = CreateObject("Outlook.Application")
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    Dim currentTime As Date
    currentTime = Time

    Set olFolder = olNamespace.Folders(checkAccount.dataFileName).Folders("[Inbox_Name]")
    Set lastSheet = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    'SheetName = "completedSht"

    ''Check the sheet name is completedSht
    For i = 1 To ThisWorkbook.Worksheets.Count
        If ThisWorkbook.Sheets(i).Name = "completedSht" Then
            sheetExists = True
            Exit For
        End If
    Next i

    ''If Sheets2 doesn't exit, add completedSht
    If Not sheetExists Then
        Set newWorksheet = ThisWorkbook.Worksheets.Add(After:=lastSheet)
        newWorksheet.Name = "completedSht"
    End If

    ''select Excel Sheet (what you want)
    Set xlSheet = ThisWorkbook.Sheets("completedSht")

    ''check fill in the data
    If ThisWorkbook.Worksheets("completedSht").Cells(1, 1).value <> "" Then
        ThisWorkbook.Worksheets("completedSht").Cells.Clear
    End If

    i = 2
    ''set Title
    xlSheet.Cells(i - 1, 1).value = "Subject"
    xlSheet.Cells(i - 1, 2).value = "ReceivedTime"
    xlSheet.Cells(i - 1, 3).value = "SenderName"
    xlSheet.Cells(i - 1, 4).value = "SR_category"
    
    ''get email data
    For Each olItem In olFolder.Items
        If (Format(olItem.ReceivedTime, "yyyy-mm-dd") = Format(Date, "yyyy-mm-dd")) Then
            xlSheet.Cells(i, 1).value = olItem.Subject
            ''ReceivedTime Format is dd/mm/yyyy hh:mm
            xlSheet.Cells(i, 2).value = olItem.ReceivedTime
            xlSheet.Cells(i, 3).value = olItem.Sender
            xlSheet.Cells(i, 4).value = olItem.Body
            'xlSheet.Cells(i, 5).value = olItem.Content
            ''add data (what you want)
            
            Dim bodyText As String
            Dim startPos As Long, endPos As Long
            bodyText = olItem.Body
            startPos = InStr(bodyText, "[]") + Len("[]")
            endPos = InStr(startPos, bodyText, "[]") - 1
            
            If startPos > 0 And endPos > startPos Then
                xlSheet.Cells(i, 4).value = Trim(Mid(bodyText, startPos, endPos - startPos))
            Else
                xlSheet.Cells(i, 4).value = "No Category"
            End If
        End If

        i = i + 1
    Next olItem

    Dim lastRow As Long
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, "A").End(xlUp).Row
    xlSheet.Range("A1:D" & lastRow).AutoFilter Field:=2

    
    xlSheet.AutoFilter.Sort.SortFields.Clear
    xlSheet.AutoFilter.Sort.SortFields.Add Key:=Range("B1:B" & lastRow), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With xlSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, "A").End(xlUp).Row
    
    currentDate = Format(Date, "mmdd")
    
    ''Copy from completedSht to Today's sheets
    For i = 1 To lastRow - 1
        valueToCopy = ThisWorkbook.Worksheets("completedSht").Range("A" & i + 1).value
        valueToCopy2 = ThisWorkbook.Worksheets("completedSht").Range("D" & i + 1).value
        ThisWorkbook.Worksheets(currentDate).Range("C" & i + 35).value = valueToCopy
        ThisWorkbook.Worksheets(currentDate).Range("G" & i + 35).value = valueToCopy2
    Next i

    ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count - 1).Activate
    Cells.Replace What:="[] ", Replacement:="", LookAt:=xlPart, _
                     SearchOrder:=xlByRows, MatchCase:=False


    Application.DisplayAlerts = False
    'ThisWorkbook.Worksheets("completedSht").Delete
    Application.DisplayAlerts = True

    ''disable memory
    Set olApp = Nothing
    Set olNamespace = Nothing
    Set olFolder = Nothing
    Set olItem = Nothing
    Set xlSheet = Nothing
    Set newWorksheet = Nothing
    Set lastSheet = Nothing
End Function

Function getCompletedRequests()
    ''create Outlook object
    Set olApp = CreateObject("Outlook.Application")
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    Dim currentTime As Date
    currentTime = Time

    Set olFolder = olNamespace.Folders(checkAccount.dataFileName).Folders("[]")
    Set lastSheet = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    'SheetName = "completedSht"

    ''Check the sheet name is completedSht
    For i = 1 To ThisWorkbook.Worksheets.Count
        If ThisWorkbook.Sheets(i).Name = "ongoingSht" Then
            sheetExists = True
            Exit For
        End If
    Next i

    ''If Sheets2 doesn't exit, add completedSht
    If Not sheetExists Then
        Set newWorksheet = ThisWorkbook.Worksheets.Add(After:=lastSheet)
        newWorksheet.Name = "ongoingSht"
    End If

    ''select Excel Sheet (what you want)
    Set xlSheet = ThisWorkbook.Sheets("ongoingSht")

    ''check fill in the data
    If ThisWorkbook.Worksheets("ongoingSht").Cells(1, 1).value <> "" Then
        ThisWorkbook.Worksheets("ongoingSht").Cells.Clear
    End If

    i = 2
    ''set Title
    xlSheet.Cells(i - 1, 1).value = "Subject"
    xlSheet.Cells(i - 1, 2).value = "ReceivedTime"
    xlSheet.Cells(i - 1, 3).value = "SenderName"
    xlSheet.Cells(i - 1, 4).value = "SR_category"
    
    ''get email data
    For Each olItem In olFolder.Items
        If (Format(olItem.ReceivedTime, "yyyy-mm-dd") = Format(Date, "yyyy-mm-dd")) Then
            xlSheet.Cells(i, 1).value = olItem.Subject
            ''ReceivedTime Format is dd/mm/yyyy hh:mm
            xlSheet.Cells(i, 2).value = olItem.ReceivedTime
            xlSheet.Cells(i, 3).value = olItem.Sender
            xlSheet.Cells(i, 4).value = olItem.Body
            'xlSheet.Cells(i, 5).value = olItem.Content
            ''add data (what you want)
            
            Dim bodyText As String
            Dim startPos As Long, endPos As Long
            bodyText = olItem.Body
            startPos = InStr(bodyText, "[]") + Len("[]")
            endPos = InStr(startPos, bodyText, "[]") - 1
            
            If startPos > 0 And endPos > startPos Then
                xlSheet.Cells(i, 4).value = Trim(Mid(bodyText, startPos, endPos - startPos))
            Else
                xlSheet.Cells(i, 4).value = "No category"
            End If
        End If

        i = i + 1
    Next olItem

    Dim lastRow As Long
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, "A").End(xlUp).Row
    xlSheet.Range("A1:D" & lastRow).AutoFilter Field:=2

    
    xlSheet.AutoFilter.Sort.SortFields.Clear
    xlSheet.AutoFilter.Sort.SortFields.Add Key:=Range("B1:B" & lastRow), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With xlSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, "A").End(xlUp).Row
    
    currentDate = Format(Date, "mmdd")
    
    ''Copy from completedSht to Today's sheets

    Application.DisplayAlerts = False
    'ThisWorkbook.Worksheets("completedSht").Delete
    Application.DisplayAlerts = True

    ''disable memory
    Set olApp = Nothing
    Set olNamespace = Nothing
    Set olFolder = Nothing
    Set olItem = Nothing
    Set xlSheet = Nothing
    Set newWorksheet = Nothing
    Set lastSheet = Nothing
End Function

Function compareReqeustStatus()
End Function
