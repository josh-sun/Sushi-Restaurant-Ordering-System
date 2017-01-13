Attribute VB_Name = "modPublicFunctions"
Option Explicit

Public strInputBox, strPic, strItemName  As String
Public intNumber, intCheckValue, X As Integer
Public dblTotal, dblPrice, dblItemPrice, dblGrandTotal, dblPaidPrice As Double
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public dblPriceStore() As Double
Public strItemStore() As String
Public intQuantityStore() As Integer
Public strInfoStore()  As String
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) As Long

    If Topmost = True Then 'Make the window topmost
        SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
        SetTopMostWindow = False
    End If
End Function

Public Function NumbersOnly(strInput)
    Dim intCounter, intChar As Integer
    Dim strChar As String
    intCounter = 1
    
    Do While intCounter <> Len(strInput) + 1
        If IsNumeric(Right(Left(strInput, intCounter), 1)) = False Then
            NumbersOnly = 1
        End If
        intCounter = intCounter + 1
    Loop
End Function
