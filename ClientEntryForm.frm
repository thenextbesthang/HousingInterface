VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ClientEntryForm 
   Caption         =   "User Form Entry"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8910
   OleObjectBlob   =   "ClientEntryForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ClientEntryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim UNIQUE_ID As Long
Dim ISws As Boolean
Dim ISrs As Boolean
Dim curRow As Long

Private Sub Services_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll ClientEntryForm, Services
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    UnhookListBoxScroll
End Sub

Private Sub EnterBtn_Click()
            If IsError(Application.Match(UNIQUE_ID, Range("a1:a1048576"), False)) Then
                curRow = Range("a1").End(xlDown).row + 1
                Cells(curRow, 1) = ClientID.Value
                Cells(curRow, 2) = LastName.Value
                Cells(curRow, 3) = FirstName.Value
                Cells(curRow, 4) = AddressNumber.Value
                Cells(curRow, 5) = AddressStreet.Value
                Cells(curRow, 6) = City.Value
                Cells(curRow, 7) = State.Value
                Cells(curRow, 8) = ZIP.Value
                Cells(curRow, 9) = HomePhone.Value
                Cells(curRow, 10) = CellPhone.Value
                Cells(curRow, 11) = Email.Value
                Cells(curRow, 12) = HouseholdIncome.Value
                Cells(curRow, 13) = Language.Value
                Cells(curRow, 14) = NumberInHousehold.Value
                Cells(curRow, 15) = FrailSeniors.Value
            Else
                MsgBox ("client already entered")
            End If
End Sub

Private Sub IntoServices_Click()

            Dim servSheet As Worksheet
            Set servSheet = Sheets("Service")
            
                curRow = servSheet.Range("a1").End(xlDown).row + 1
                
                
                servSheet.Cells(curRow, 1) = Now()
                servSheet.Cells(curRow, 2) = Services.Value
                servSheet.Cells(curRow, 3) = ClientID.Value & " - " & FirstName.Value & " " & LastName.Value
                servSheet.Cells(curRow, 4) = Providers.Value
                servSheet.Cells(curRow, 5) = AddressNumber.Value & " " & AddressStreet.Value
                servSheet.Cells(curRow, 6) = City.Value
                servSheet.Cells(curRow, 7) = State.Value
                servSheet.Cells(curRow, 8) = ZIP.Value
                servSheet.Cells(curRow, 9) = HomePhone.Value
                servSheet.Cells(curRow, 10) = CellPhone.Value
                servSheet.Cells(curRow, 11) = Email.Value
                servSheet.Cells(curRow, 12) = HouseholdIncome.Value
                servSheet.Cells(curRow, 13) = FrailSeniors.Value
                servSheet.Cells(curRow, 14) = Language.Value
            
            Dim tmpArr As Variant
            tmpArr = Split(Services.Value)
            
            Dim service As Long
            service = tmpArr(0)
            
           ' MsgBox (tmpArr(0))
            
            
            
            Dim row As Long
            row = CLng(Application.Match(service, Sheets("Service Codes").Range("a1:A1048576"), False))
            
           
            
            
            Dim val As String
            val = Application.index(Sheets("Service Codes").Range("a1:c1048576"), row, 3)
            val = Replace(val, "%", FirstName.Value & " " & LastName.Value)
            servSheet.Cells(curRow, 15) = val
            
            
            
            
            

End Sub



Private Sub CloseBox_Click()

    Unload Me
End Sub

Private Sub GetID_Click()
    

    'set up connections with Nina's housing database
    Dim cmd As New ADODB.Command
    Dim conn As New ADODB.Connection
    Dim Rs As New ADODB.Recordset
    Dim strConn As String
    Dim strSQL As String

    'connection string
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=F:\Housing\Access 2007 Housing Database.accdb; Persist Security Info=False"

    'open connection database
    conn.Open strConn
    
    'sql statement
    strSQL = "SELECT * FROM Clients WHERE (((Clients.FirstName)='" & FirstName.Value & "') AND ((Clients.LastName)='" & LastName.Value & "'));"

    'open connection with the recordset
    Rs.Open strSQL, conn, adOpenDynamic, adLockOptimistic

    'check if its on worksheet or in recordset
    ISws = isOnWorksheet()
    ISrs = Not Rs.EOF
    
    'if its in the worksheet
     If ISws = True Then
        fillForm UNIQUE_ID, Rs
    'if its in the recordset
     ElseIf ISrs = True Then
        fillForm Rs.Fields("Client ID"), Rs
     Else
        'if its not in the recordset or the worksheet, find the highest ID in either the recordset or the worksheet
        'and set the new Unique ID to be one more than the highest
        
        Dim newMaxWS As Long
        Dim newMaxRS As Long
        
        newMaxWS = WorksheetFunction.max(Sheets("Client Codes").Range("A1:A1048576")) + 1
        
        Rs.Close
        strSQL = "SELECT MAX(Clients.[Client ID]) FROM Clients;"
        Rs.Open strSQL, conn, adOpenDynamic, adLockOptimistic
        newMaxRS = CLng(Rs.Fields(0)) + 1
        
        If newMaxRS > newMaxWS Then
            UNIQUE_ID = newMaxRS
            fillForm UNIQUE_ID, Rs
        ElseIf newMaxWS > newMaxRS Then
            UNIQUE_ID = newMaxWS
            fillForm UNIQUE_ID, Rs
        Else
            fillForm UNIQUE_ID, Rs
        End If
    End If
 
End Sub

Private Function fillForm(UniqueID As Long, Rs As Recordset)

    Dim i As Long
    

    If ISws = True Then
                    
            'check if a value is entered into the text fields
            'if it is, enter the value in the form
            'if it isn't, enter in a blank
            ClientID.Value = Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Client Code"",1:1,0))")
            
            If Not Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Address Street"",1:1,0))") = vbNullString Then
               AddressStreet.Value = Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Address Street"",1:1,0))")
            Else
                AddressStreet.Value = ""
            End If
            
            If Not Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Address Number"",1:1,0))") = vbNullString Then
               AddressNumber.Value = Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Address Number"",1:1,0))")
            Else
                AddressNumber.Value = ""
            End If
            
            If Not Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""City"",1:1,0))") = vbNullString Then
               City.Value = Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""City"",1:1,0))")
            Else
               City.Value = ""
            End If
            
            If Not Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""State"",1:1,0))") = vbNullString Then
               State.Value = Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""State"",1:1,0))")
            Else
               State.Value = ""
            End If

            If Not Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Zip"",1:1,0))") = vbNullString Then
               ZIP.Value = Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Zip"",1:1,0))")
            Else
               ZIP.Value = ""
            End If

            If Not Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Home"",1:1,0))") = vbNullString Then
               HomePhone.Value = Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Home"",1:1,0))")
            Else
               HomePhone.Value = ""
            End If
            
            If Not Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Home"",1:1,0))") = vbNullString Then
               HomePhone.Value = Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Home"",1:1,0))")
            Else
               HomePhone.Value = ""
            End If
            
            If Not Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Cell"",1:1,0))") = vbNullString Then
               CellPhone.Value = Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Cell"",1:1,0))")
            Else
               CellPhone.Value = ""
            End If
                       
            If Not Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Email"",1:1,0))") = vbNullString Then
               Email.Value = Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Email"",1:1,0))")
            Else
               Email.Value = ""
            End If
            
            If Not Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Household Income (Yearly, Approximate)"",1:1,0))") = vbNullString Then
               HouseholdIncome.Value = Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Household Income (Yearly, Approximate)"",1:1,0))")
            Else
               HouseholdIncome.Value = ""
            End If
            
            If Not Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Frail Seniors In Household"",1:1,0))") = vbNullString Then
                FrailSeniors.Value = Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Frail Seniors In Household"",1:1,0))")
            Else
                FrailSeniors.Value = ""
            End If
                       
            If Not Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Languages Spoken in Household"",1:1,0))") = vbNullString Then
               Language.Value = Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Languages Spoken in Household"",1:1,0))")
            Else
               Language.Value = ""
            End If

            If Not Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Number of People in Household"",1:1,0))") = vbNullString Then
              NumberInHousehold.Value = Evaluate("INDEX(1:1048576,MATCH(" & UniqueID & ",A:A,0),MATCH(""Number of People in Household"",1:1,0))")
            Else
               NumberInHousehold.Value = ""
            End If





    ElseIf ISrs = True Then
            'enter in some values into form
                    ClientID.Value = Rs.Fields("Client ID")
                    
            'check if a value is entered into the text fields
            'if it is, enter the value in the form
            'if it isn't, enter in a blank
            
                    If Not Rs.Fields("ZIP") = vbNullString Then
                        ZIP.Value = Rs.Fields("ZIP")
                    Else
                        ZIP.Value = ""
                    End If
                    
                    If Not Rs.Fields("State") = vbNullString Then
                        State.Value = Rs.Fields("State")
                    Else
                        State.Value = ""
                    End If
                    
                    If Not Rs.Fields("City") = vbNullString Then
                        City.Value = Rs.Fields("City")
                    Else
                        City.Value = ""
                    End If
                    
                    If Not Rs.Fields("Phone") = vbNullString Then
                        CellPhone.Value = Rs.Fields("Phone")
                    Else
                        CellPhone.Value = ""
                    End If

                    If Not Rs.Fields("Email") = vbNullString Then
                        Email.Value = Rs.Fields("Email")
                    Else
                        Email.Value = ""
                    End If
                    
                    If Not Rs.Fields("Language") = vbNullString Then
                        Language.Value = Rs.Fields("Language")
                    Else
                        Language.Value = ""
                    End If
                    
            'enter address, which is in database combined into one unit
                    If Not Rs.Fields("Address") = vbNullString Then
                        Dim myArr As Variant
                        myArr = Split(Rs.Fields("Address"), " ")
                        AddressNumber.Value = myArr(0)
                        For i = 1 To UBound(myArr)
                            AddressStreet.Value = AddressStreet.Value & " " & myArr(i)
                        Next i
                    Else
                        AddressNumber.Value = ""
                        AddressStreet.Value = ""
                        
                    End If

                    If Rs.Fields("Below Median") = True Then
                        HouseholdIncome.Value = "<37,000"
                    Else
                        HouseholdIncome.Value = ">=37,000"
                    End If
    Else
        ClientID.Value = UNIQUE_ID
        MsgBox (FirstName.Value & " " & LastName.Value & " needs to be entered into system...")
    End If

End Function
Private Function isOnWorksheet() As Boolean


    'to be used for matching and updates
    Dim base As Long
    base = 1
    'Dim curRow As Long
    Dim foundWB As Boolean
    
    'use the late-bound application match method to find out where the firstname and lastname values are in the worksheet, if found
    Dim first As Long
    Dim last As Long

       'First check to make sure if both values are in the worksheet
        Do While Not IsError(Application.Match(FirstName.Value, Range("c" & base & ":c1048576"), False)) And Not IsError((Application.Match(LastName.Value, Range("b" & base & ":b1048576"), False)))
           'if it is in the worksheet, find where it is
       
                    first = Application.Match(FirstName.Value, Range("c" & base & ":c1048576"), False)
                    last = Application.Match(LastName.Value, Range("b" & base & ":b1048576"), False)
         
                If first < last Then
                    curRow = last
                    base = first + 1
                  '  curRow = curRow + first
                ElseIf last < first Then
                    curRow = first
                    base = last + 1

                Else
                    foundWB = True
                    UNIQUE_ID = Cells(first, 1)
                    Exit Do
                End If
        Loop
        

    
    isOnWorksheet = foundWB

End Function
Private Sub PrintBtn_Click()
        Dim inputNumStart As Long
        Dim inputNumEnd As Long
        
        Dim i As Long
        
        On Error GoTo errhandler
        inputNumStart = CLng(Trim(FromBox.Value))
        inputNumEnd = CLng(Trim(ToBox.Value))
        
        UpdateActionsRows inputNumStart, inputNumEnd
        
        Exit Sub
errhandler:         MsgBox ("Please enter a valid number")

End Sub


Private Sub FirstName_Exit(ByVal Cancel As MSForms.ReturnBoolean)

'ByVal Cancel As MSForms.ReturnBoolean
     If Len(FirstName.Value) < 1 Or Not IsLetter(FirstName.Value) Then
     
         FirstName.BackColor = &HFF& ' change the color of the textbox
         MsgBox "Please Enter a Correct Name"
         ' setting Cancel to True means the user cannot leave this textbox
         ' until the value is in the proper date format
         
         
         Cancel = True
     Else
          FirstName.BackColor = &H80000005 ' change color of the textbox
     End If

End Sub
Private Sub LastName_Exit(ByVal Cancel As MSForms.ReturnBoolean)

'ByVal Cancel As MSForms.ReturnBoolean
     If Len(LastName.Value) < 1 Or Not IsLetter(LastName.Value) Then
     
         LastName.BackColor = &HFF& ' change the color of the textbox
         MsgBox "Please Enter a Correct Name"
         ' setting Cancel to True means the user cannot leave this textbox
         ' until the value is in the proper date format
         Cancel = True
     Else
          FirstName.BackColor = &H80000005 ' change color of the textbox
     End If

End Sub

Function IsLetter(strValue As String) As Boolean
    Dim intPos As Integer
    For intPos = 1 To Len(strValue)
        Select Case Asc(Mid(strValue, intPos, 1))
            Case 65 To 90, 97 To 122, 32
                IsLetter = True
            Case Else
                IsLetter = False
                Exit For
        End Select
    Next
End Function
Private Sub UserForm_Initialize()

   ' initialize the list of services
    Dim i As Long
    For i = 0 To WorksheetFunction.CountA(Worksheets("Service Codes").Range("A1:A1045876"))
        Services.AddItem Sheets("Service Codes").Cells(i + 1, 1) & " - " & Sheets("Service Codes").Cells(i + 1, 2)
    Next i

    For i = 0 To WorksheetFunction.CountA(Worksheets("Provider Codes").Range("A1:A1045876"))
        Providers.AddItem Sheets("Provider Codes").Cells(i + 1, 1) & " - " & Sheets("Provider Codes").Cells(i + 1, 2)
    Next i

     Providers.Value = "25 - Barry Polinsky"

    
End Sub
