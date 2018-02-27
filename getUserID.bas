Attribute VB_Name = "getUserID"
Option Explicit
Sub getUserID()

    'find max ID entered into client code worksheet
    Dim myRange As Range
    Dim maxID As Long
    Dim clientSheet As Worksheet
    
    Set clientSheet = Sheets("Client Codes")
    Set myRange = clientSheet.Range("A1:A1048576")
    maxID = WorksheetFunction.max(myRange) + 1
    
    
    'set up connections with Nina's housing database
    Dim cmd As New ADODB.Command
    Dim conn As New ADODB.Connection
    Dim Rs As New ADODB.Recordset
    Dim strConn As String
    Dim strSQL As String
    Dim FirstName As String
    Dim LastName As String
    Dim newID As Long

    'find first and last name of most recently entered client
    FirstName = clientSheet.Range("b1").End(xlDown)
    LastName = clientSheet.Range("c1").End(xlDown)

    'connection string
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=F:\Housing\Access 2007 Housing Database.accdb; Persist Security Info=False"

    'open connection database
    conn.Open strConn
    
    'sql statement
    strSQL = "SELECT * FROM Clients WHERE (((Clients.FirstName)='" & FirstName & "') AND ((Clients.LastName)='" & LastName & "'));"

    'open connection with the recordset
    Rs.Open strSQL, conn, adOpenDynamic, adLockOptimistic

    'the previous open statement should return a recordset which has met the conditions
    'if EOF is true, then no record is found
    'if false than a recordset is found
    
    If Rs.EOF Then 'If the returned RecordSet is empty

    'find the highest ID in the clients table, and add 1
       Rs.Close
       strSQL = "SELECT MAX(Clients.[Client ID]) FROM Clients;"
       Rs.Open strSQL, conn, adOpenDynamic, adLockOptimistic
       newID = Rs.Fields(0) + 1

    'is the highest ID in the clients worksheet greater than the highest ID in the clients table + 1?
       If maxID > newID Then
            'then the new ID should be drawn from the clients worksheet
            Cells(Range("a1").End(xlDown).row + 1, 1) = maxID
       Else
            Cells(Range("a1").End(xlDown).row + 1, 1) = newID
       End If
       

       
    End If



    



End Sub
