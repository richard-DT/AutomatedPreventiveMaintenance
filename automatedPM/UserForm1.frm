VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Automated Maintenance System - ST08 CD Pants Machine (Kimberly-Clark Singapore Tuas Mill)"
   ClientHeight    =   2988
   ClientLeft      =   -2676
   ClientTop       =   -12408
   ClientWidth     =   7356
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public whatsclicked As String
Public Sub RefreshData()
    ' Code to update controls with new data
    CheckSignal
    ScheduleAMS
End Sub


Private Sub rdtSystems_Click()
    ' Cancel the scheduled call to the ScrollText sub
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:00:01"), "ScrollText", , False
    On Error GoTo 0
End Sub

Private Sub UserForm_Activate()

    ScheduleAMS
    CheckSignal
    With Me
        .Width = Application.Width
        .Height = Application.Height
        .Left = 0
        .Top = 0
    End With
    
    
    Dim health As String
        health = ThisWorkbook.Sheets("data").Range("R15").value * 100
        
        
    Load_Chart
    
    Dim value As Double
    value = ThisWorkbook.Sheets("data").Range("R15").value

    Me.overallLabel.caption = Format(value, "0%")

        If value >= 0.85 Then
            Me.overallLabel.ForeColor = RGB(102, 255, 51) ' Green
        Else
            Me.overallLabel.ForeColor = RGB(255, 0, 0) ' Red
        End If
        
    ' Start scrolling the text
    ScrollText
        
End Sub
Private Sub Load_Chart()

    Dim sh As Worksheet
    Dim chartKo As Chart
    
    Set sh = ThisWorkbook.Sheets("health")
    Set chartKo = sh.Shapes("HealthChart").Chart
        chartKo.Parent.Width = Application.CentimetersToPoints(9)
        chartKo.Parent.Height = Application.CentimetersToPoints(9)

    
    chartKo.Export VBA.Environ("TEMP") & Application.PathSeparator & "chartKo.jpg"
    
    Me.Image1.Picture = LoadPicture(VBA.Environ("TEMP") & Application.PathSeparator & "chartKo.jpg")
End Sub



'Private Sub UserForm_Initialize()
        'AddMinMaxButtons Me
'End Sub

Private Sub UserForm_Initialize()
    ' Set the initial position of scrollingLabel to the left edge of the form
    scrollingLabel.Left = 0
    
    ' Set the initial position of scrollingLabel2 to the right edge of the form
    scrollingLabel2.Left = UserForm1.Width - scrollingLabel2.Width
End Sub


Private Sub Workbook_Open()
    Application.OnTime TimeValue("06:00:00"), "CheckDueDate"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Application.OnTime TimeValue("06:00:00"), "CheckDueDate", , False
    On Error GoTo 0
End Sub

Private Sub SEC0_Click()

    whatsclicked = "SEC0"
    UserForm2.Show
End Sub

Private Sub SEC1_Click()
    whatsclicked = "SEC1"
    UserForm2.Show
End Sub
Private Sub SEC2_Click()
    whatsclicked = "SEC2"
    UserForm2.Show
End Sub
Private Sub SEC3_Click()
    whatsclicked = "SEC3"
    UserForm2.Show
End Sub
Private Sub SEC4_Click()
    whatsclicked = "SEC4"
    UserForm2.Show
End Sub
Private Sub SEC5_Click()
    whatsclicked = "SEC5"
    UserForm2.Show
End Sub
Private Sub SEC6_Click()
    whatsclicked = "SEC6"
    UserForm2.Show
End Sub
Private Sub SEC7_Click()
    whatsclicked = "SEC7"
    UserForm2.Show
End Sub
Private Sub SEC8_Click()
    whatsclicked = "SEC8"
    UserForm2.Show
End Sub
Private Sub SEC9_Click()
    whatsclicked = "SEC9"
    UserForm2.Show
End Sub
Private Sub SEC10_Click()
    whatsclicked = "SEC10"
    UserForm2.Show
End Sub
Private Sub AuxilliaryFans_Click()
    whatsclicked = "AuxilliaryFans"
    UserForm2.Show
End Sub
Private Sub Bagger_Click()
    whatsclicked = "Bagger"
    UserForm2.Show
End Sub
Private Sub BarrierLayerUnwind_Click()
    whatsclicked = "BarrierLayerUnwind"
    UserForm2.Show
End Sub
Private Sub FlapUnwind_Click()
    whatsclicked = "FlapUnwind"
    UserForm2.Show
End Sub
Private Sub focke_Click()
    whatsclicked = "focke"
    UserForm2.Show
End Sub
Private Sub InnerPanelUnwind_Click()
    whatsclicked = "InnerPanelUnwind"
    UserForm2.Show
End Sub
Private Sub LinerUnwind_Click()
    whatsclicked = "LinerUnwind"
    UserForm2.Show
End Sub
Private Sub Moonwalk_Click()
    whatsclicked = "Moonwalk"
    UserForm2.Show
End Sub
Private Sub OCClothUnwind_Click()
    whatsclicked = "OCClothUnwind"
    UserForm2.Show
End Sub
Private Sub Pulp_Click()
    whatsclicked = "Pulp"
    UserForm2.Show
End Sub
Private Sub Stacker_Click()
    whatsclicked = "Stacker"
    UserForm2.Show
End Sub
Private Sub SurgeUnwind_Click()
    whatsclicked = "SurgeUnwind"
    UserForm2.Show
End Sub


Private Sub WrapLayerUnwind_Click()
    whatsclicked = "WrapLayerUnwind"
    UserForm2.Show
End Sub
Private Sub WSM_Click()
    whatsclicked = "WSM"
    UserForm2.Show
End Sub
Private Sub TapeUnwind_Click()
    whatsclicked = "TapeUnwind"
    UserForm2.Show
End Sub
Private Sub FEOverend_Click()
    whatsclicked = "FEOverend"
    UserForm2.Show
End Sub
Private Sub LEOverend_Click()
    whatsclicked = "LEOverend"
    UserForm2.Show
End Sub
Private Sub PanelElastic_Click()
    whatsclicked = "PanelElastic"
    UserForm2.Show
End Sub
Private Sub priorityItems_Click()
    whatsclicked = "priorityItems"
    UserForm2.Show
End Sub
Private Sub CFA_Click()
    whatsclicked = "CFA"
    UserForm2.Show
End Sub
Private Sub SAM_Click()
    whatsclicked = "SAM"
    UserForm2.Show
End Sub
Private Sub OuterPanelUnwind_Click()
    whatsclicked = "OuterPanelUnwind"
    UserForm2.Show
End Sub
Private Sub OCPolyUnwind_Click()
    whatsclicked = "OCPolyUnwind"
    UserForm2.Show
End Sub
Private Sub RELUnwind_Click()
    whatsclicked = "RELUnwind"
    UserForm2.Show
End Sub
Private Sub melter1_Click()
    whatsclicked = "melter1"
    UserForm2.Show
End Sub
Private Sub melter2_Click()
    whatsclicked = "melter2"
    UserForm2.Show
End Sub
Private Sub melter3_Click()
    whatsclicked = "melter3"
    UserForm2.Show
End Sub
Private Sub melter4_Click()
    whatsclicked = "melter4"
    UserForm2.Show
End Sub
Private Sub melter5_Click()
    whatsclicked = "melter5"
    UserForm2.Show
End Sub
Private Sub melter6_Click()
    whatsclicked = "melter6"
    UserForm2.Show
End Sub
Private Sub melter7_Click()
    whatsclicked = "melter7"
    UserForm2.Show
End Sub
Private Sub melter8_Click()
    whatsclicked = "melter8"
    UserForm2.Show
End Sub
Private Sub melter9_Click()
    whatsclicked = "melter9"
    UserForm2.Show
End Sub
Sub CheckSignal()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim btn As MSForms.CommandButton
    Dim captions As Variant
    Dim caption As Variant
 
    ' Specify the worksheet where the data is located
    Set ws = ThisWorkbook.Sheets("Allitems")
 
    ' Find the last used row in column B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
 
    ' Define an array of button captions to check
    'captions = Array("SEC0", "SEC1", "SEC2", "SEC3", "SEC4", "SEC5", "SEC6", "SEC7", "SEC8", "SEC9", "SEC10", _
                     "AuxilliaryFans", "Bagger", "BarrierLayerUnwind", "FlapUnwind", "focke", "InnerPanelUnwind", _
                     "LinerUnwind", "Moonwalk", "OCClothUnwind", "OCPolyUnwind", "OuterPanelUnwind", _
                     "Pulp", "RELUnwind", "SAM", "Stacker", "SurgeUnwind", "WrapLayerUnwind", _
                     "CFA", "WSM", "TapeUnwind", "FEOverend", "LEOverend", _
                     "PanelElastic", "priorityItems")
    captions = Application.Transpose(Worksheets("data").Range("G5:G48").value)
 
    ' Loop through each button caption and check the signal
    For Each caption In captions
        ' Reset the button color to its default value
        Set btn = UserForm1.Controls(caption)
        btn.BackColor = &HFF00&
 
        ' Loop through each row and check the signal
        For i = 4 To lastRow
        
            If ws.Cells(i, "C").value = 1 And ws.Cells(i, "s").value <> 0 Then
                priorityItems.BackColor = RGB(255, 0, 0) ' Set the button color to red
            End If

            ' Check if the value in column B matches the button caption and the value in column C is 1
            If ws.Cells(i, "B").value = caption And ws.Cells(i, "C").value = 1 Then
                btn.BackColor = RGB(255, 0, 0) ' Set the button color to red
                Exit For ' Exit the loop since a matching condition is found
            
            End If
        Next i
    Next caption
    
    'pending items and health condition
    
    Me.pendPriority.caption = "Pending Items: " & Worksheets("data").Range("J3").value

    Dim controlNames As Variant
    controlNames = Application.Transpose(Worksheets("data").Range("G5:G48").value)

    Dim z As Integer
    For z = LBound(controlNames) To UBound(controlNames)
        'Debug.Print "Processing control: " & controlNames(z)
        On Error Resume Next
        Me.Controls("pend" & controlNames(z)).caption = "Pending Items: " & Worksheets("data").Range("H" & z + 4).value
        On Error Resume Next
        Me.Controls("health" & controlNames(z)).caption = "Health: " & Format(Worksheets("data").Range("K" & z + 4).value, "0%")
    Next z

    
    
End Sub

Private Sub SpinButton1_Change()
    
    With SpinButton1
    .Min = 40
    .Max = 400
    .SmallChange = 5
End With

    Frame1.Zoom = SpinButton1.value
    


End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    ' Cancel the scheduled call to the ScrollText sub
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:00:01"), "ScrollText", , False
    On Error GoTo 0

End Sub




