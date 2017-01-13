VERSION 5.00
Begin VB.Form frmSashimi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Sashimi"
   ClientHeight    =   8235
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   12060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   12060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return To Main Menu"
      Height          =   735
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7080
      Width           =   2895
   End
   Begin VB.CommandButton cmdSashimi2 
      Height          =   3495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton cmdSashimi3 
      Height          =   3495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   3855
   End
   Begin VB.CommandButton cmdSashimi4 
      Height          =   3615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   3855
   End
   Begin VB.CommandButton cmdSashimi1 
      Height          =   3495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label lblPrice 
      BackStyle       =   0  'Transparent
      Caption         =   "All Sashimi Are $0.99 Per Piece"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   9120
      TabIndex        =   6
      Top             =   4080
      Width           =   2655
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
      ForeColor       =   &H00008000&
      Height          =   3495
      Left            =   9000
      TabIndex        =   4
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmSashimi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Price()
        'imgDisplay.Picture = LoadPicture(CurDir & "/Picture/sashimi.jpg")
    strInputBox = InputBox("Enter Amount You Would Like To Order: ", "Enter Number")
    If strInputBox <> "" And NumbersOnly(strInputBox) <> 1 And Val(strInputBox) > 0 And Val(strInputBox) <= 99 Then
        intNumber = Val(strInputBox)
        dblPrice = dblItemPrice * intNumber
        frmOrder.txtOrder.Text = frmOrder.txtOrder.Text & strItemName & " X " & intNumber & vbNewLine
        ReDim Preserve strItemStore(X), dblPriceStore(X), intQuantityStore(X)
        strItemStore(X) = strItemName
        dblPriceStore(X) = dblPrice
        intQuantityStore(X) = intNumber
        MsgBox "Your Orders Have Been Added"
        X = X + 1
    ElseIf Val(strInputBox) <= 0 Or Val(strInputBox) > 99 Then
        MsgBox "Invalid: You Can Only Order 1 - 99 Pieces"
    Else
        MsgBox "Invalid Input"
    End If
End Function

Private Sub cmdDone_Click()
    Unload frmSashimi
    frmMainMenu.Show
End Sub

Private Sub cmdSashimi1_Click()
    dblItemPrice = 0.99
    strItemName = "Sake (Salmon) Sashimi"
    Call Price
End Sub

Private Sub cmdSashimi1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblDescription.Caption = "Fresh Sake (Salmon) Sashimi"
End Sub

Private Sub cmdSashimi2_Click()
    dblItemPrice = 1.99
    strItemName = "Akami (Tuna) Sashimi"
    Call Price
End Sub

Private Sub cmdSashimi2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblDescription.Caption = "Fresh Cut of tuna with a dark red color. It's the lowest in fat"
End Sub

Private Sub cmdSashimi3_Click()
    dblItemPrice = 1.99
    strItemName = "Ika (Squid) Sashimi"
    Call Price
End Sub

Private Sub cmdSashimi3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblDescription.Caption = "Fresh Squid Sashimi"
End Sub

Private Sub cmdSashimi4_Click()
    dblItemPrice = 0.99
    strItemName = "Saba (Mackerel) Sashimi"
    Call Price
End Sub

Private Sub cmdSashimi4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblDescription.Caption = "Fresh Mackerel Sashimi"
End Sub

Private Sub Form_Load()
    cmdSashimi1.Picture = LoadPicture(CurDir & "/Picture/sashimi_but.jpg")
    cmdSashimi2.Picture = LoadPicture(CurDir & "/Picture/sashimi2.jpg")
    cmdSashimi3.Picture = LoadPicture(CurDir & "/Picture/sashimi3.jpg")
    cmdSashimi4.Picture = LoadPicture(CurDir & "/Picture/sashimi4.jpg")
    frmSashimi.Picture = LoadPicture(CurDir & "/Picture/sashimiback.jpg")
End Sub
