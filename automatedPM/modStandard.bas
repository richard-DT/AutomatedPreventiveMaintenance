Attribute VB_Name = "modStandard"
Public userName As String
Public userWID As String

' In a standard code module
Sub ScrollText()
    ' Move the label 5 pixels to the right
    UserForm1.scrollingLabel.Left = UserForm1.scrollingLabel.Left + 50
    UserForm1.scrollingLabel2.Left = UserForm1.scrollingLabel2.Left - 50


    ' If the label has moved off the right edge of the form, move it back to the left edge
    If UserForm1.scrollingLabel.Left > UserForm1.Width Then
        UserForm1.scrollingLabel.Left = 0 - UserForm1.scrollingLabel.Width
    End If
    
    If UserForm1.scrollingLabel2.Left < 0 - UserForm1.scrollingLabel2.Width Then
    UserForm1.scrollingLabel2.Left = UserForm1.Width
    End If
    

    UserForm1.scrollingLabel.caption = Sheets("data").Range("Q17").value
    UserForm1.scrollingLabel2.caption = Sheets("data").Range("Q18").value

    ' Schedule the next call to this sub
    Application.OnTime Now + TimeValue("00:00:01"), "ScrollText"
End Sub

' In the UserForm code module
Private Sub UserForm_Activate()
    ' Start scrolling the text
    ScrollText
End Sub

