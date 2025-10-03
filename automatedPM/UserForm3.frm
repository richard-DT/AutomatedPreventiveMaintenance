VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Name and WID"
   ClientHeight    =   384
   ClientLeft      =   -204
   ClientTop       =   -912
   ClientWidth     =   1068
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private exitButtonPressed As Boolean
Private Sub enterButton_Click()

    ' Get values from ComboBoxes
    userName = nameBox.value
    userWID = widBox.value
    
    ' Close UserForm3
    Me.Hide
End Sub
Private Sub UserForm_Initialize()
    ' Set default values for ComboBoxes
    nameBox.value = ""
    widBox.value = ""
    Me.Height = 106
    Me.Width = 270
    Me.StartUpPosition = 2
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        exitButtonPressed = True
        Debug.Print "exitButtonPressed set to True"
    End If
End Sub

Public Function GetExitButtonPressed() As Boolean
    GetExitButtonPressed = exitButtonPressed
End Function

Public Sub ResetExitButtonPressed()
    exitButtonPressed = False
    Debug.Print "exitButtonPressed reset to False"
End Sub

