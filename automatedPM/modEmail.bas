Attribute VB_Name = "modEmail"

Sub SendAMS1()
    
    Dim i As Integer
    Dim lastRow As Integer
    Dim objOutlook As Object
    Dim objMail As Object
    Dim ws As Worksheet
    Dim body As String


    Set ws = ThisWorkbook.Sheets("Allitems")
    
    lastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).Row
    
    ' Send the first email
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0)
    
    ' Build the email body for findings and suggestions
    body = "Findings and Recommendations:" & vbCrLf & vbCrLf
            
   
    For i = 4 To lastRow
        
        If ws.Cells(i, "Q").value = 1 Then
        
            body = body & "[ID#] " & ws.Cells(i, "A").value & vbCrLf
            body = body & "[Sub Section] " & ws.Cells(i, "D").value & vbCrLf
            body = body & "[Item] " & ws.Cells(i, "E").value & vbCrLf
            body = body & "[Remarks] " & ws.Cells(i, "O").value & vbCrLf
            body = body & "[By:] " & ws.Cells(i, "P").value & vbCrLf & vbCrLf
        
        End If
    Next i
        body = body & "***this is a computer generated message from AMS , pls do not reply***" & vbCrLf
        body = body & "*** if no data is listed above, it means no pending items for today ***" & vbCrLf
        
    Dim toRange1 As Range, ccRange1 As Range
    Dim toAddress1 As String, ccAddress1 As String
    Dim cell1 As Range

    Set toRange1 = Sheets("data").Range("L4:L8")
    Set ccRange1 = Sheets("data").Range("M4:M12")

    For Each cell1 In toRange1
        toAddress1 = toAddress1 & cell1.value & "; "
    Next cell1

    For Each cell1 In ccRange1
        ccAddress1 = ccAddress1 & cell1.value & "; "
    Next cell1

    With objMail
        .To = toAddress1
        .CC = ccAddress1
        .Subject = "AMS Findings and Recommendations for ST08"
        .body = body
        .Send
    End With
    
    Set objMail = Nothing
    
End Sub

    ' High Priority Items
    
Sub SendAMS2()
    
    Dim j As Integer
    Dim hulingRow As Integer
    Dim objeOutlook As Object
    Dim objeMail As Object
    Dim wss As Worksheet
    Dim bod As String

    Set wss = ThisWorkbook.Sheets("Allitems")
    hulingRow = wss.Cells(wss.Rows.Count, "S").End(xlUp).Row
    
    
    ' Send the 2nd email
    Set objeOutlook = CreateObject("Outlook.Application")
    Set objeMail = objeOutlook.CreateItem(0)

    ' Build the email body for 1 week before replacement
    For j = 4 To hulingRow
    
        If wss.Cells(j, "S").value = 1 Then
    
            bod = bod & "[ID#] " & wss.Cells(j, "A").value & vbCrLf
            bod = bod & "[Sub Section] " & wss.Cells(j, "D").value & vbCrLf
            bod = bod & "[Item] " & wss.Cells(j, "E").value & " " & wss.Cells(j, "F").value & vbCrLf
            bod = bod & "[Due Date] " & wss.Cells(j, "L").value & vbCrLf & vbCrLf
    
        End If
    Next j
        bod = bod & "*** this is a computer generated message from AMS , pls do not reply ***" & vbCrLf
        bod = bod & "*** if no data is listed above, it means no pending items for today ***" & vbCrLf
        
    Dim toRange2 As Range, ccRange2 As Range
    Dim toAddress2 As String, ccAddress2 As String
    Dim cell2 As Range

    Set toRange2 = Sheets("data").Range("L4:L8")
    Set ccRange2 = Sheets("data").Range("M4:M12")

    For Each cell2 In toRange2
        toAddress2 = toAddress2 & cell2.value & "; "
    Next cell2

    For Each cell2 In ccRange2
        ccAddress2 = ccAddress2 & cell2.value & "; "
    Next cell2

    With objeMail
        .To = toAddress2
        .CC = ccAddress2
        .Subject = "AMS High Priority Items for ST08"
        .body = bod
        .Send
    End With

            
    Set objeMail = Nothing

End Sub

Sub SendAMS3()
    
    Dim k As Integer
    Dim kulelat As Integer
    Dim obOutlook As Object
    Dim obMail As Object
    Dim ss As Worksheet
    Dim bo As String

    Set ss = ThisWorkbook.Sheets("Allitems")
    kulelat = ss.Cells(ss.Rows.Count, "R").End(xlUp).Row
    
    
    ' Send the 2nd email
    Set obOutlook = CreateObject("Outlook.Application")
    Set obMail = obOutlook.CreateItem(0)

    ' Build the email body for 1 week before replacement
    For k = 4 To kulelat
    
        If ss.Cells(k, "R").value = 1 Then
    
            bo = bo & "[ID#] " & ss.Cells(k, "A").value & vbCrLf
            bo = bo & "[Sub Section] " & ss.Cells(k, "D").value & vbCrLf
            bo = bo & "[Item] " & ss.Cells(k, "E").value & vbCrLf
            bo = bo & "[Due Date] " & ss.Cells(k, "L").value & vbCrLf & vbCrLf
    
        End If
    Next k
        bo = bo & "***this is a computer generated message from AMS , pls do not reply***" & vbCrLf
        bo = bo & "*** if no data is listed above, it means no pending items for today ***" & vbCrLf
    
    'TO and CC address
    
    Dim toRange As Range, ccRange As Range
    Dim toAddress As String, ccAddress As String
    Dim cell As Range

    Set toRange = Sheets("data").Range("L4:L12")
    Set ccRange = Sheets("data").Range("M4:M12")

    For Each cell In toRange
    toAddress = toAddress & cell.value & "; "
    Next cell

    For Each cell In ccRange
    ccAddress = ccAddress & cell.value & "; "
    Next cell

    With obMail
        .To = toAddress
        .CC = ccAddress
        .Subject = "AMS Items due for replacement within 10 days for ST08"
        .body = bo
        .Send
    End With

    Set obMail = Nothing

End Sub
Sub ScheduleAMS()
    ' Schedule the first email to be sent at 7 AM everyday
    Application.OnTime TimeValue("07:00:00"), "SendAMS1", , True
    
    ' Schedule the second email to be sent at 8 AM
    Application.OnTime TimeValue("07:15:00"), "SendAMS2", , True
    
    ' Schedule the third email to be sent at 9 AM
    Application.OnTime TimeValue("07:30:00"), "SendAMS3", , True
End Sub

