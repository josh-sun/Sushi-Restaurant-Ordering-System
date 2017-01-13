VERSION 5.00
Begin VB.Form frmCheckout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Out"
   ClientHeight    =   7620
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   11130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDone 
      Caption         =   "Return to Main Menu"
      Height          =   735
      Left            =   4560
      TabIndex        =   14
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton cmdPay 
      Caption         =   "Continue To Payment"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2160
      TabIndex        =   13
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate"
      Height          =   735
      Left            =   6600
      TabIndex        =   12
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label lblNumTotal 
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   7920
      TabIndex        =   11
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label lblNumTax 
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   8520
      TabIndex        =   10
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label lblNumSubtotal 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   8280
      TabIndex        =   9
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   6840
      TabIndex        =   8
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax (13%):"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   6840
      TabIndex        =   7
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label lblSubTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   6840
      TabIndex        =   6
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label lblPrice 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblQuan 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblItem 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblItemName 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   4455
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label lblQuantity 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   4455
      Left            =   4200
      TabIndex        =   4
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label lblItemPrice 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   4455
      Left            =   7440
      TabIndex        =   5
      Top             =   720
      Width           =   3615
   End
End
Attribute VB_Name = "frmCheckout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalc_Click()
    If dblTotal = 0 Then
        MsgBox "You Have to Order Before Checkout!"
    Else
        lblNumSubtotal.Caption = "$ " & FormatNumber(dblTotal, 2)
        lblNumTax.Caption = "$ " & FormatNumber(0.13 * dblTotal, 2)
        dblGrandTotal = dblTotal * 1.13
        lblNumTotal.Caption = "$" & FormatNumber(dblGrandTotal, 2)
        cmdPay.Enabled = True
    End If
End Sub

Private Sub cmdDone_Click()
    frmOrder.Show
    cmdPay.Enabled = False
    Unload frmCheckout
    frmMainMenu.Show
End Sub

Private Sub cmdPay_Click()
    Unload frmCheckout
    frmPay.Show
End Sub

Private Sub Form_Load()
    frmOrder.Hide
    Dim intCounter As Integer
    frmCheckout.Picture = LoadPicture(CurDir & "/Picture/menu.jpg")
    dblTotal = 0
    intCounter = 0
    Do While intCounter <> X
        lblItemName.Caption = lblItemName.Caption & strItemStore(intCounter) & vbCrLf
        lblQuantity.Caption = lblQuantity.Caption & intQuantityStore(intCounter) & vbCrLf
        lblItemPrice.Caption = lblItemPrice.Caption & dblPriceStore(intCounter) & vbCrLf
        dblTotal = dblTotal + dblPriceStore(intCounter)
        intCounter = intCounter + 1
    Loop
End Sub

