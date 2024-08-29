''using visual basic ver.2013
Sub GetEmailDataFromOutlook()
    Dim olApp As Object
    Dim olNamespace As Object
    Dim olFolder As Object
    Dim olItem As Object
    Dim xlSheet As Worksheet
    Dim i As Integer
    Dim olFolderPath As String
    Dim newWorksheet As Worksheet
    Dim sheetExists As Boolean
    Dim lastSheet As Worksheet
    Dim pcusername As String
    Dim dataFileName As String

    Dim valueToCopy As Variant
    Dim valueToCopy2 As Variant
    Dim currentDate As String

    Dim bodyText As String
    Dim startPos As Long, endPos As Long
    
    ''create Outlook object
    Set olApp = CreateObject("Outlook.Application")
    Set olNamespace = olApp.GetNamespace("MAPI")

    pcusername = Environ$("Username")

    Select Case pcusername
        Case "", ""
            dataFileName = ""

        Case ""
            dataFileName = ""

        Case "", ""
            dataFileName = ""

        Case "", ""
            dataFileName = ""

    End Select
    
    ''select E-mail folder (ex: Inbox)
    'Set olFolder = olNamespace.GetDefaultFolder(olFolderInbox)
    'olFolderPath = "\\" & dataFileName & "\[InboxName]"
    Set olFolder = olNamespace.Folders(dataFileName).Folders("[InboxName]")
     Set lastSheet = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    'SheetName = "Sheet2"
    


    ''Check the sheet name is Sheet2
    For i = 1 To ThisWorkbook.Worksheets.Count
        If ThisWorkbook.Sheets(i).Name = "Sheet2" Then
            sheetExists = True
            Exit For
        End If
    Next i

    ''If Sheets2 doesn't exit, add Sheet2
    If Not sheetExists Then
        Set newWorksheet = ThisWorkbook.Worksheets.Add(After:=lastSheet)
        newWorksheet.Name = "Sheet2"
    End If

    ''select Excel Sheet
    Set xlSheet = ThisWorkbook.Sheets("Sheet2")

    ''check fill in the data
    If ThisWorkbook.Worksheets("Sheet2").Cells(1, 1).Value <> "" Then
        ThisWorkbook.Worksheets("Sheet2").Cells.Clear
    End If

    i = 2
    ''set Title
    xlSheet.Cells(i - 1, 1).Value = "Subject"
    xlSheet.Cells(i - 1, 2).Value = "ReceivedTime"
    xlSheet.Cells(i - 1, 3).Value = "SenderName"
    xlSheet.Cells(i - 1, 4).Value = "category"

    'MsgBox Format(Date, "yyyy-mm-dd")

    ''get email data
    For Each olItem In olFolder.Items
        'MsgBox Format(olItem.ReceivedTime, "yyyy-mm-dd")
        If (Format(olItem.ReceivedTime, "yyyy-mm-dd") = Format(Of Date, "yyyy-mm-dd")) Then
            xlSheet.Cells(i, 1).Value = olItem.Subject
            ''ReceivedTime Format is dd/mm/yyyy hh:mm
            xlSheet.Cells(i, 2).Value = olItem.ReceivedTime
            xlSheet.Cells(i, 3).Value = olItem.Sender
            xlSheet.Cells(i, 4).Value = olItem.Body
            ''add data (what you want)

            ''1. get all of mail body to olItem.Body
            ''2. find strings location "category" and "request" using InStr function
            ''3. extract specific word between 2 strings using Mid function
            ''4. Remove the front and back spaces with the Trim function and save the extracted text to the cell.
            ''5. If dont search "category" or "request" strings, display message by "no category"


            bodyText = olItem.Body
            startPos = InStr(bodyText, "category") + Len("category")
            endPos = InStr(startPos, bodyText, "request") - 1

            If startPos > 0 And endPos > startPos Then
                xlSheet.Cells(i, 4).Value = Trim(Mid(bodyText, startPos, endPos - startPos))
            Else
                xlSheet.Cells(i, 4).Value = "no category"
            End If
        End If

        i = i + 1
    Next olItem

    Dim lastRow As Long
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, "A").End(xlUp).Row
    xlSheet.Range("A1:D" & lastRow).AutoFilter Field:=2


    xlSheet.AutoFilter.Sort.SortFields.Clear
    xlSheet.AutoFilter.Sort.SortFields.Add Key:=Range("B1:B" & lastRow),
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With xlSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


    lastRow = xlSheet.Cells(xlSheet.Rows.Count, "A").End(xlUp).Row

    currentDate = Format(Of Date, "mmdd")

    ''Copy from sheet2 to Today's sheets
    For i = 1 To lastRow - 1
        valueToCopy = ThisWorkbook.Worksheets("Sheet2").Range("A" & i + 1).Value
        valueToCopy2 = ThisWorkbook.Worksheets("Sheet2").Range("D" & i + 1).Value
        ThisWorkbook.Worksheets(currentDate).Range("C" & i + 35).Value = valueToCopy
        ThisWorkbook.Worksheets(currentDate).Range("G" & i + 35).Value = valueToCopy2
    Next i

    ''find some word and remove it
    ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count - 1).Activate
    Cells.Replace What:="[???] ", Replacement:="", LookAt:=xlPart,
                     SearchOrder:=xlByRows, MatchCase:=False


    ''Delete sheet2
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Sheet2").Delete
    Application.DisplayAlerts = True

    ''disable memory
    Set olApp = Nothing
    Set olNamespace = Nothing
    Set olFolder = Nothing
    Set olItem = Nothing
    Set xlSheet = Nothing
    Set newWorksheet = Nothing
    Set lastSheet = Nothing
End Sub






