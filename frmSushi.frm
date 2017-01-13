VERSION 5.00
Begin VB.Form frmSushiOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Sushi"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11160
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H8000000D&
      Caption         =   "Return To Menu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CheckBox chkMaki2 
      BackColor       =   &H8000000D&
      Caption         =   "Tekkamaki X 8 --------------------------------------------------------- $6.99"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      Top             =   1680
      Width           =   4455
   End
   Begin VB.CheckBox chkMaki3 
      BackColor       =   &H8000000D&
      Caption         =   "California Roll X 6 ----------------------------------------------------- $6.99"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   10
      Top             =   2400
      Width           =   4455
   End
   Begin VB.CheckBox chkMaki4 
      BackColor       =   &H8000000D&
      Caption         =   "Futomaki X 6 ------------------------------------------------------------ $7.99"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   9
      Top             =   3120
      Width           =   4455
   End
   Begin VB.CheckBox chkMaki5 
      BackColor       =   &H8000000D&
      Caption         =   "Tsunamayo Maki  X 6 ---------------------------------------------- $7.99"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   8
      Top             =   3840
      Width           =   4455
   End
   Begin VB.CheckBox chkSushi3 
      BackColor       =   &H8000000D&
      Caption         =   "Ikura Gukan X 4 ------------------------------------------------------- $6.99"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2400
      Width           =   4455
   End
   Begin VB.CheckBox chkSushi4 
      BackColor       =   &H8000000D&
      Caption         =   "Ebi Nigiri X 6 ------------------------------------------------------------- $5.99"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   3120
      Width           =   4455
   End
   Begin VB.CheckBox chkSushi5 
      BackColor       =   &H8000000D&
      Caption         =   "Tamagoyaki X 4 ------------------------------------------------------- $6.99"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   3840
      Width           =   4455
   End
   Begin VB.CheckBox chkMaki1 
      BackColor       =   &H8000000D&
      Caption         =   "Kappa Maki X 8 -------------------------------------------------------- $5.99"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   4
      Top             =   960
      Width           =   4455
   End
   Begin VB.Frame fraMaki 
      BackColor       =   &H8000000D&
      Caption         =   "Maki Rolls"
      BeginProperty Font 
         Name            =   "Blackoak Std"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   4335
      Left            =   5760
      TabIndex        =   1
      Top             =   360
      Width           =   4935
   End
   Begin VB.Frame fraSushi 
      BackColor       =   &H8000000D&
      Caption         =   "Sushi"
      BeginProperty Font 
         Name            =   "Blackoak Std"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4335
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4815
      Begin VB.CheckBox chkSushi2 
         BackColor       =   &H8000000D&
         Caption         =   "Sake Nigiri X 6 ---------------------------------------------------------- $7.99"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   4455
      End
      Begin VB.CheckBox chkSushi1 
         BackColor       =   &H8000000D&
         Caption         =   "Unagi  X 6 ---------------------------------------------------------------- $7.99"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   4455
      End
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   5760
      TabIndex        =   13
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Image imgDisplay 
      Height          =   2415
      Left            =   360
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   4815
   End
End
Attribute VB_Name = "frmSushiOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Price()
    If intCheckValue = 1 Then
        'imgDisplay.Picture = LoadPicture(CurDir & "/Picture/sashimi.jpg")
        strInputBox = InputBox("Enter Amount You Would Like To Order: ", "Enter Number")
        If strInputBox = "" Or NumbersOnly(strInputBox) = 1 Then
            MsgBox "Invalid Input"
            intCheckValue = 0
        ElseIf Val(strInputBox) > 99 Or Val(strInputBox) <= 0 Then
            MsgBox "You Can Only Buy 1 - 99 Pieces"
            intCheckValue = 0
        Else
            intNumber = Val(strInputBox)
            dblPrice = dblItemPrice * intNumber
            frmOrder.txtOrder.Text = frmOrder.txtOrder.Text & strItemName & " X " & intNumber & vbNewLine
            ReDim Preserve strItemStore(X), dblPriceStore(X), intQuantityStore(X)
            strItemStore(X) = strItemName
            dblPriceStore(X) = dblPrice
            intQuantityStore(X) = intNumber
            X = X + 1
            MsgBox "Your Orders Have Been Added"
        End If
    Else
        strInputBox = ""
        dblPrice = 0
        intNumber = 0
    End If
End Function

Private Sub chkMaki1_Click()
        dblItemPrice = 5.99
        strItemName = "Kappa Maki"
        intCheckValue = chkMaki1.Value
        Call Price
        chkMaki1.Value = intCheckValue
End Sub

Private Sub chkMaki1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgDisplay.Picture = LoadPicture(CurDir & "/Picture/maki1.jpg")
    lblDescription.Caption = "8 Pieces Of Cucumber Rolls"
End Sub

Private Sub chkMaki2_Click()
    dblItemPrice = 6.99
    strItemName = "Tekkamaki"
    intCheckValue = chkMaki2.Value
    Call Price
    chkMaki2.Value = intCheckValue
End Sub

Private Sub chkMaki2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgDisplay.Picture = LoadPicture(CurDir & "/Picture/maki2.jpg")
    lblDescription.Caption = "8 Pieces Of Tuna Thin Rolls"
End Sub

Private Sub chkMaki3_Click()
    dblItemPrice = 6.99
    strItemName = "California Rolls"
    intCheckValue = chkMaki3.Value
    Call Price
    chkMaki3.Value = intCheckValue
End Sub

Private Sub chkMaki3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgDisplay.Picture = LoadPicture(CurDir & "/Picture/maki3.jpg")
    lblDescription.Caption = "6 Pieces Of California Rolls"
End Sub

Private Sub chkMaki4_Click()
    dblItemPrice = 7.99
    strItemName = "Futomaki"
    intCheckValue = chkMaki4.Value
    Call Price
    chkMaki4.Value = intCheckValue
End Sub

Private Sub chkMaki4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgDisplay.Picture = LoadPicture(CurDir & "/Picture/maki4.jpg")
    lblDescription.Caption = "6 Pieces Of thick rolls containing egg, kanpyo, cucumber and mushrooms. "
End Sub

Private Sub chkMaki5_Click()
    dblItemPrice = 7.99
    strItemName = "Tsunamayo"
    intCheckValue = chkMaki5.Value
    Call Price
    chkMaki5.Value = intCheckValue
End Sub

Private Sub chkMaki5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgDisplay.Picture = LoadPicture(CurDir & "/Picture/maki5.jpg")
    lblDescription.Caption = "6 Pieces Of Tuna and Mayonnaise"
End Sub

Private Sub chkSushi1_Click()
    dblItemPrice = 7.99
    strItemName = "Unagi"
    intCheckValue = chkSushi1.Value
    Call Price
    chkSushi1.Value = intCheckValue
End Sub

Private Sub chkSushi1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgDisplay.Picture = LoadPicture(CurDir & "/Picture/sushi1.jpg")
    lblDescription.Caption = "6 Pieces Of Japanese eel Sushi"
End Sub

Private Sub chkSushi2_Click()
        dblItemPrice = 7.99
        strItemName = "Sake Nigiri"
        intCheckValue = chkSushi2.Value
    Call Price
    chkSushi2.Value = intCheckValue
End Sub

Private Sub chkSushi2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgDisplay.Picture = LoadPicture(CurDir & "/Picture/sake.jpg")
    lblDescription.Caption = "6 Pieces Of Fresh Salmon Sushi"
End Sub

Private Sub chkSushi3_Click()
    dblItemPrice = 6.99
    strItemName = "Ikura Gukan"
    intCheckValue = chkSushi3.Value
    Call Price
    chkSushi3.Value = intCheckValue
End Sub

Private Sub chkSushi3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgDisplay.Picture = LoadPicture(CurDir & "/Picture/sushi3.jpg")
    lblDescription.Caption = "4 Pieces Of Salmon Roe"
End Sub

Private Sub chkSushi4_Click()
    dblItemPrice = 5.99
    strItemName = "Ebi Nigiri"
    intCheckValue = chkSushi4.Value
    Call Price
    chkSushi4.Value = intCheckValue
End Sub

Private Sub chkSushi4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgDisplay.Picture = LoadPicture(CurDir & "/Picture/sushi4.jpg")
    lblDescription.Caption = "6 Pieces Of Cooked Shrimp Sushi"
End Sub

Private Sub chkSushi5_Click()
    dblItemPrice = 6.99
    strItemName = "Tamagoyaki"
    intCheckValue = chkSushi5.Value
    Call Price
    chkSushi5.Value = intCheckValue
End Sub

Private Sub chkSushi5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgDisplay.Picture = LoadPicture(CurDir & "/Picture/sushi5.jpg")
    lblDescription.Caption = "4 Pieces Of Fried Egg Sushi"
End Sub

Private Sub cmdReturn_Click()
    Unload frmSushiOrder
    frmMainMenu.Show
End Sub

Private Sub Form_Load()
    frmSushiOrder.Picture = LoadPicture(CurDir & "/Picture/sushiback.jpg")
End Sub
