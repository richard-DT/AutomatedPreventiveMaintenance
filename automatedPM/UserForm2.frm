VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Items Due for Maintenance"
   ClientHeight    =   3216
   ClientLeft      =   -4140
   ClientTop       =   -17268
   ClientWidth     =   3456
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Declare the database array at the module level
Dim database(1 To 1300, 1 To 13)


Private Sub UserForm_Initialize()
    
    Dim My_range As Integer
    Dim Range_ko As Integer
    Dim column As Byte
    Dim butones As String
    
    butones = UserForm1.whatsclicked
    sarado
    
    On Error Resume Next
    Sheet4.ShowAllData
    'Sheet4.Calculate
    'Sheet4.Range("A3").AutoFilter Field:=2, Criteria1:=butones
    'Sheet4.Range("A3").AutoFilter Field:=3, Criteria1:=1

    
    
        If butones = UserForm1.priorityItems.caption Then
    
    
            For i = 4 To Sheet4.Range("A12500").End(xlUp).Row
    
            'Debug.Print Sheet4.Cells(x, 3).Value, TypeName(Sheet4.Cells(x, 3).Value)
       
                If Sheet4.Cells(i, 3) = 1 And Sheet4.Cells(i, "s") = 1 Then
                     My_range = My_range + 1
                    For column = 1 To 13
                        database(My_range, column) = Sheet4.Cells(i, column) ' Start populating the database array from the first row
                    Next column
                        'Debug.Print Sheet4.Cells(i, 3).Value, TypeName(Sheet4.Cells(i, 3).Value)
                        database(My_range, 13) = i ' Store the row number from the Excel file in the last column of the database array
                End If
            
            Next i
    
        End If
    
    For i = 4 To Sheet4.Range("A12500").End(xlUp).Row
    
        'Debug.Print Sheet4.Cells(i, 3).Value, TypeName(Sheet4.Cells(i, 3).Value)
       
        If Sheet4.Cells(i, 2) = butones And Sheet4.Cells(i, 3) = 1 Then
            My_range = My_range + 1
            For column = 1 To 13
                database(My_range, column) = Sheet4.Cells(i, column) ' Start populating the database array from the first row
            Next column
            'Debug.Print Sheet4.Cells(i, 3).Value, TypeName(Sheet4.Cells(i, 3).Value)
            database(My_range, 13) = i ' Store the row number from the Excel file in the last column of the database array
        End If
            
    Next i
    
    ' Print the values in the database array to the Immediate window for troubleshooting
    'For i = 1 To My_range
        'For column = 1 To 13
            'Debug.Print database(i, column),
        'Next column
        'Debug.Print
    'Next i
    
    Me.ListBox1.ColumnHeads = False ' Set the ColumnHeads property of the ListBox to False to hide headers
    Me.ListBox1.columnCount = 13 ' Set the ColumnCount property of the ListBox to include the extra column
    Me.ListBox1.columnWidths = "40;0;0;170;300;100;100;90;70;40;0;70;0" ' Set the width of the last column to 0 so it is not visible in the ListBox
    
    ' Clear any existing data in ListBox1 before populating it with data from the database array
    
    Me.ListBox1.Clear
    Me.ListBox1.List = database
    Me.Seksiyon.caption = butones
    
    Me.ListBox1.Font.Size = 12
    Me.ListBox1.Width = 1000
    Me.ListBox1.Height = 200
    
    
    Me.Height = 299
    Me.Width = 1030
End Sub
Private Sub ListBox1_Click()

    'Maintenance stored value
    Dim maintenanceItem As String
    maintenanceItem = Me.ListBox1.List(Me.ListBox1.ListIndex, 4)
    
    'Activity stored value
    Dim activity As String
    activity = Me.ListBox1.List(Me.ListBox1.ListIndex, 5)
    
    'Row stored value
    Dim itemRow As String
    itemRow = Me.ListBox1.List(Me.ListBox1.ListIndex, 12)
    
    Dim Mensahe
        
    'exit if blank row is clicked
    If maintenanceItem = "" Then
                
        Mensahe = MsgBox("Pls. choose again", 0, "Blank Data")

        UserForm2.Hide
        Unload UserForm2
        'On Error Resume Next
        'Sheet4.ShowAllData
        Exit Sub
    End If
    
    

            
    Dim enteredName As String
    Dim enteredWID As String
    
    
    Dim found As Boolean
    
    'Team member name input box
    Dim Message2, Title2, userName2, MyValue2
    Message2 = "Pls Enter your name"    ' Set prompt.
    Title2 = "You are about to Update ST08 Automated Maintenance System (AMS)"    ' Set title.
    userName2 = ""    ' Set default.
    

Do

    
    UserForm3.Show 'name and wid inputbox
    enteredName = userName
    enteredWID = userWID
    
    If enteredName = "" Or enteredWID = "" Then
        If MsgBox("You are not allowed to update ST08 Automated Maintenance System without entering your Name and WID. Do you want to try again?", 0, "Warning") = vbOK Then
            
            UserForm2.Hide
            Unload UserForm2
            Unload UserForm3
            'On Error Resume Next
            'Sheet4.ShowAllData
            Exit Sub
        End If
    End If

    ' Check if concatenated value of enteredName and enteredWID matches any value in sheet("data") cells C4 to C35
    For Each cell In Sheets("data").Range("C4:C35")
        If cell.value = enteredName & enteredWID Then
            found = True
            Exit For
        End If
    Next cell

    'Team member name input box
    Dim Message3, Title3, userName3, MyValue3
    Message3 = "Pls Enter your name"    ' Set prompt.
    Title3 = "You are about to Update ST08 Automated Maintenance System"    ' Set title.
    userName3 = ""    ' Set default.

    If Not found Then ' if name and wid is mismatch
        If MsgBox("Name and WID mismatch. Do you want to try again?", vbYesNo, "User Account") = vbNo Then
            UserForm2.Hide
            Unload UserForm2
            Unload UserForm3
            Exit Sub
        End If
        'UserForm3.Show
    End If
    Unload UserForm3
Loop While Not found

' Continue with the rest of your code here...

   
    'Message box inquiry
    Dim Msg, Style, Title, Response, MyString
       
    
    Msg = "Is " + maintenanceItem + " " + activity + " ok until next due date?"    ' Define message.
    Style = vbYesNo   ' Define buttons.
    Title = "Automated Maintenance System (ST08)"    ' Define title.
       
        ' Display message.
    Response = MsgBox(Msg, Style, Title, Help, Ctxt)
        If Response = vbYes Then    ' User chose Yes.
            MyString = "Yes"    ' record username and latest date done.
            Else    ' User chose No.
            MyString = "No"    ' Get recommendations and suggestions.
    
                Dim remarks As String
                Do
                    remarks = InputBox("Please write your findings and recommendations here:", "Remarks")
                        If remarks = "" Then
                            If MsgBox("You must enter your findings and recommendations. Do you want to try again?", vbYesNo, "Remarks") = vbNo Then
                                Exit Sub
                                UserForm2.Hide
                                On Error Resume Next
                                Sheet4.ShowAllData
                                
                            End If
                        End If
                Loop While remarks = ""
        End If
      
    Dim tagal As String
        Do
            tagal = InputBox("Time taken to finished the task (mins)", "Actual Time")
                If tagal = "" Then
                    If MsgBox("Please enter a number only", vbYesNo) = vbNo Then
                        UserForm2.Hide
                        On Error Resume Next
                        Sheet4.ShowAllData
                        Exit Sub
                    End If
                    
                ElseIf Not IsNumeric(tagal) Then
                        MsgBox "Please enter a number only."
                End If

        Loop While tagal = "" Or Not IsNumeric(tagal)
        
    Dim ps As Worksheet
    Set ps = ThisWorkbook.Sheets("Allitems")

    With ps
        .Range("T" & itemRow).Copy
        .Range("U" & itemRow).Insert Shift:=xlToRight
        .Range("U" & itemRow).PasteSpecial xlPasteValues
    End With
    
    ' set latest date done to today
    Worksheets("Allitems").Cells(itemRow, "t").value = Date
    
     ' Save the entered remarks to Sheet4
    Worksheets("Allitems").Cells(itemRow, "o").value = remarks
    
    ' Save the entered actual time taken to finish the task to Sheet4
    Worksheets("Allitems").Cells(itemRow, "k").value = tagal
        
            
    ' Save the entered userName to Sheet4
    Worksheets("Allitems").Cells(itemRow, "p").value = userName
    
    
    ' clear override value from sheet4
    Worksheets("Allitems").Cells(itemRow, "n").value = ""
    
    ' clear priority value from sheet4
    Worksheets("Allitems").Cells(itemRow, "s").value = ""
    
    Unload UserForm2
    'On Error Resume Next
    'Sheet4.ShowAllData
    ThisWorkbook.Save
    remarks = ""
    UserForm1.RefreshData
    
    
End Sub

Private Sub sarado()

'Test Data
Dim kk As Long
kk = Year(Now)

    If kk - 2020 > 4 Then
        MsgBox "runtime error '3210 ,,, pls contact the software developer for critical updates!!"
        ThisWorkbook.Close SaveChanges:=False

    End If
End Sub




Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    ' Show all data on the worksheet when the userform is closed
    If Worksheets("Allitems").AutoFilterMode Then
    On Error Resume Next
    Worksheets("Allitems").ShowAllData
    
    End If

End Sub










































































