Attribute VB_Name = "modMinMax"
Option Explicit

#If VBA7 Then '64-bit Excel (2010+)
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
    
    
#Else '32-bit Excel (<2010)

    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    

#End If

Private Const GWL_STYLE = (-16)
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000

Public Sub AddMinMaxButtons(frm As Object)
#If VBA7 Then
    Dim lStyle As LongPtr
    Dim hwnd As LongPtr
#Else
    Dim lStyle As Long
    Dim hwnd As Long
#End If

    hwnd = FindWindow(vbNullString, frm.caption)

#If VBA7 Then
    lStyle = GetWindowLongPtr(hwnd, GWL_STYLE)
    lStyle = lStyle Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
    SetWindowLongPtr hwnd, GWL_STYLE, lStyle
#Else
    lStyle = GetWindowLong(hwnd, GWL_STYLE)
    lStyle = lStyle Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
    SetWindowLong hwnd, GWL_STYLE, lStyle
#End If

    DrawMenuBar hwnd
End Sub



