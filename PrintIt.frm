VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PrintIt 
   Caption         =   "Print"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5415
   OleObjectBlob   =   "PrintIt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PrintIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ClearBtn_Click()

    FromBox.Value = ActiveSheet.Range("a1").End(xlDown).row
    ToBox = ActiveSheet.Range("a1").End(xlDown).row
    
    
End Sub

Private Sub CloseBtn_Click()
    Unload Me
    
End Sub

Private Sub FromBox_Change()

    
End Sub

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

Private Sub ToBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)

      Dim curRow As Long
      
      curRow = ActiveSheet.Range("a1").End(xlDown).row


'ByVal Cancel As MSForms.ReturnBoolean
     If Not IsNumeric(ToBox.Value) Then
         ToBox.BackColor = &HFF& ' change the color of the textbox
         MsgBox "not number"
         ' setting Cancel to True means the user cannot leave this textbox
         ' until the value is in the proper date format
         Cancel = True
     Else
          ToBox.BackColor = &H80000005 ' change color of the textbox
     End If
     

      
    If (ToBox.Value > curRow) Or (ToBox.Value < 0) Or ToBox.Value < FromBox.Value Then
    
         ToBox.BackColor = &HFF& ' change the color of the textbox
         MsgBox "number error"
         ' setting Cancel to True means the user cannot leave this textbox
         ' until the value is in the proper date format
         Cancel = True
      Else
          ToBox.BackColor = &H80000005 ' change color of the textbox
     End If
     

     
End Sub
Private Sub FromBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
      Dim curRow As Long
      
      curRow = ActiveSheet.Range("a1").End(xlDown).row

    
     If Not IsNumeric(FromBox.Value) Then
        MsgBox ("not number")
     
         FromBox.BackColor = &HFF& ' change the color of the textbox
         MsgBox "Illegal value"
         ' setting Cancel to True means the user cannot leave this textbox
         ' until the value is in the proper date format
         Cancel = True
     Else
          FromBox.BackColor = &H80000005 ' change color of the textbox
     End If
     
     If (FromBox.Value > curRow) Or (FromBox.Value < 0) Or FromBox.Value > ToBox.Value Then
        MsgBox ("number error")
        
         FromBox.BackColor = &HFF& ' change the color of the textbox
         MsgBox "Illegal value"
         ' setting Cancel to True means the user cannot leave this textbox
         ' until the value is in the proper date format
         Cancel = True
      Else
          FromBox.BackColor = &H80000005 ' change color of the textbox
     End If
     


End Sub
Private Sub ToBox_Change()
 '   ToBox.Value = ToBoxSpin.Value
    

    
End Sub

Private Sub ToBoxSpin_Change()
  '  ToBox.Value = ToBoxSpin.Value
    
End Sub


Private Sub UserForm_Initialize()

        Dim curRow As Long
        
        
        curRow = ActiveSheet.Range("a1").End(xlDown).row
        ToBox.Value = curRow
        FromBox.Value = curRow


End Sub
