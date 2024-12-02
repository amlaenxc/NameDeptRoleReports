Attribute VB_Name = "Module4"
' I had been hearing from many librarians that they wanted to be able to see, given
' a list of names, department and or role information, for example to know the department
' affiliations of people who swiped in to attend an event at the library. The usual
' university system doesn't give this data, only names. So, I wrote this script
' to grab that data from the Outlook directory.
' NB: This only works in classic Outlook as Microsoft hasn't added developer tools
' to new Outlook yet.
Sub GetDirectoryInfoByPartialNameFromOAB()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.NameSpace
    Dim olAddressBook As Outlook.AddressList
    Dim olEntries As Outlook.AddressEntries
    Dim olEntry As Outlook.AddressEntry
    Dim olExchangeUser As Outlook.ExchangeUser
    Dim Names As Collection
    Dim fileSystem As Object
    Dim textStream As Object
    Dim line As String
    Dim parts() As String
    Dim LastName As String
    Dim FirstName As String
    Dim found As Boolean
    
    Dim xlApp As Excel.Application
    Dim xlWorkbook As Excel.Workbook
    Dim xlWorksheet As Excel.Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim nameCol As String
    Dim jobTitleCol As String
    Dim departmentCol As String
    
    Dim excelFilePath As String
    excelFilePath = "C:\Users\alex32\OneDrive - Stanford\Documents\NameDeptRoleReports\OpenHouse2024TreasureHuntData\EastAsia2024Raffle.xlsx"
    
    Set Names = New Collection
    
    Set xlApp = New Excel.Application
    Set xlWorkbook = xlApp.Workbooks.Open(excelFilePath)
    Set xlWorksheet = xlWorkbook.Sheets(1) ' Assuming data is in the first sheet
    
    lastRow = xlWorksheet.Cells(xlWorksheet.Rows.count, "A").End(xlUp).Row
    
    nameCol = "B"
    jobTitleCol = "E"
    departmentCol = "F"
    
    For i = 2 To lastRow
        Names.Add xlWorksheet.Cells(i, nameCol).Value
    Next i
    
    Set olApp = Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    On Error Resume Next
    ' Try to access the Offline Address Book
    For Each olAddressBook In olNamespace.AddressLists
        If olAddressBook.Name = "Offline Global Address List" Then
            Set olEntries = olAddressBook.AddressEntries
            Exit For
        End If
    Next olAddressBook
    On Error GoTo 0
    
    If olEntries Is Nothing Then
        Debug.Print "Offline Global Address List not found or inaccessible."
        Exit Sub
    End If
    
    ' Insert two new columns for Job Title and Department
    xlWorksheet.Columns(jobTitleCol & ":" & departmentCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    For i = 2 To lastRow
        found = False
        vname = xlWorksheet.Cells(i, nameCol).Value
        parts = Split(vname, " ")
        If UBound(parts) >= 1 Then
            FirstName = Trim(parts(0))
            LastName = Trim(parts(1))
        Else
            FirstName = Trim(parts(0))
            LastName = ""
        End If
        For Each olEntry In olEntries
            If InStr(1, olEntry.Name, LastName, vbTextCompare) > 0 And _
               InStr(1, olEntry.Name, FirstName, vbTextCompare) > 0 Then
                found = True
                Set olExchangeUser = olEntry.GetExchangeUser
                If Not olExchangeUser Is Nothing Then
                    xlWorksheet.Cells(i, jobTitleCol).Value = olExchangeUser.JobTitle
                    xlWorksheet.Cells(i, departmentCol).Value = olExchangeUser.Department
                Else
                    Debug.Print "No ExchangeUser found for Name: " & vname
                End If
                Exit For
            End If
        Next olEntry
        If Not found Then
            Debug.Print "No AddressEntry found for Name: " & vname
        End If
        Debug.Print vname & " found!"
    Next i
    
    xlWorkbook.Save
    xlWorkbook.Close
    xlApp.Quit
    
    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
    
    Debug.Print "Done"
End Sub
