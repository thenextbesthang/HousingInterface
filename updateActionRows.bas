Attribute VB_Name = "updateActionRows"
Option Explicit

Private MedIncome80 As Long


Sub UpdateActionsRows(numStart As Long, numEnd As Long)
   
'set up word application, document
    Dim objWord As Word.Application
    Dim objDoc As Document
    Set objWord = CreateObject("Word.Application")
    Dim fileName As String
    Dim curRow As Long


    While numStart <= numEnd
        Set objDoc = objWord.Documents.Open("F:\Housing\Reports\Blank Intake Form2.docx")

        writeHousingReport objDoc, numStart

        fileName = "F:\Housing\Reports\Client Reports\" & numStart & "." & Left(Cells(numStart, 3), InStr(1, Cells(numStart, 3), " ") - 1) & "." & Left(Cells(numStart, 2), InStr(1, Cells(numStart, 2), " ") - 1) & "." & Month(Cells(numStart, 1)) & "." & Day(Cells(numStart, 1)) & "." & Year(Cells(numStart, 1)) & ".docx"

    '    objDoc.SaveAs2 fileName, wdFormatDocumentDefault
        objDoc.PrintOut
        
    '    updateHousingDatabase numStart
        numStart = numStart + 1
        objDoc.Close (False)
    Wend

    objWord.Quit
'     'Step 5

End Sub
Sub writeHousingReport(doc As Document, userInput As Long)

       'setup loop variables
        Dim tempString As String
        Dim bkMark As Bookmark
        MedIncome80 = 37000
        
        
        'i is the index that tracks the column supposed to be copied from the excel file

        'insert items into document
        On Error Resume Next
        For Each bkMark In doc.Bookmarks
        
            'date time
            If bkMark.Name = "A" Then
                    bkMark.Range.InsertAfter (Cells(userInput, 1))
            'service id, service provided
            ElseIf bkMark.Name = "AA" Then
                    bkMark.Range.InsertAfter (Cells(userInput, 2))
            'access number, client
            ElseIf bkMark.Name = "AB" Then
                    bkMark.Range.InsertAfter (Cells(userInput, 3))
            'staff id, name
            ElseIf bkMark.Name = "AC" Then
                    bkMark.Range.InsertAfter (Cells(userInput, 4))
            'mailing address
            ElseIf bkMark.Name = "AD" Then
                    bkMark.Range.InsertAfter (Cells(userInput, 5) & " " & Cells(userInput, 6) & " " & Cells(userInput, 7) & " " & Cells(userInput, 8))
            'phone numbers
            ElseIf bkMark.Name = "AE" Then
                    bkMark.Range.InsertAfter ("Phone 1 = " & Cells(userInput, 9) & ", Phone 2 = " & Cells(userInput, 10))
            'email
            ElseIf bkMark.Name = "AF" Then
                    bkMark.Range.InsertAfter (Cells(userInput, 11))
            'income
            ElseIf bkMark.Name = "AG" Then
                    bkMark.Range.InsertAfter (Cells(userInput, 12))
            'frail seniors
            ElseIf bkMark.Name = "AH" Then
                    bkMark.Range.InsertAfter (Cells(userInput, 13))
            'language
            ElseIf bkMark.Name = "AI" Then
                    bkMark.Range.InsertAfter (Cells(userInput, 14))
            'action
            ElseIf bkMark.Name = "AJ" Then
                    bkMark.Range.InsertAfter (Cells(userInput, 15))
            End If
            
        Next

End Sub

Sub updateHousingDatabase(curRow As Long)


    Dim cn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    Dim strSQL As String
    Dim strSql2 As String
    Dim strConnection As String
    
    Set cn = New ADODB.Connection
    Set Rs = New ADODB.Recordset
    
    
    
    cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                             "Data Source=F:\Housing\Access 2007 Housing Database.accdb;" & _
                             "Jet OLEDB:Engine Type=5;" & _
                             "Persist Security Info=False;"
        
     strSQL = "INSERT INTO [Services and Clients] ([First Name], [Last Name], [Last Contact], [Client ID], [Services ID], Address, City, State, Zip)" & _
               " VALUES ('" & Cells(curRow, 8) & "' , '" & Cells(curRow, 7) & "', #" & Cells(curRow, 1) & "#, " & Cells(curRow, 4) & ", " & Cells(curRow, 3) & ", '" & Cells(curRow, 10) & "', '" & Cells(curRow, 11) & "', '" & Cells(curRow, 12) & "', " & Cells(curRow, 13) & ");"

     strSql2 = "INSERT INTO Clients (FirstName, LastName, [Last Contact], ClientID, ServicesID, Address, City, State, Zip)" & _
               " VALUES ('" & Cells(curRow, 8) & "' , '" & Cells(curRow, 7) & "', #" & Cells(curRow, 1) & "#, " & Cells(curRow, 4) & ", " & Cells(curRow, 3) & ", '" & Cells(curRow, 10) & "', '" & Cells(curRow, 11) & "', '" & Cells(curRow, 12) & "', " & Cells(curRow, 13) & ");"
 
    
    Rs.Open strSQL, cn.ConnectionString
    Rs.Open strSql2, cn.ConnectionString
    
    cn.ConnectionTimeout = 10
    
    
'    rs.Close
    Set Rs = Nothing
  '  cn.Close
    Set cn = Nothing

End Sub


