VERSION 5.00
Begin VB.Form frmCashForm 
   Caption         =   "Cash"
   ClientHeight    =   7785
   ClientLeft      =   2280
   ClientTop       =   2325
   ClientWidth     =   11310
   LinkTopic       =   "Form2"
   ScaleHeight     =   7785
   ScaleWidth      =   11310
   Begin VB.CommandButton cmdDone 
      Caption         =   "Finish Paying"
      Enabled         =   0   'False
      Height          =   735
      Left            =   6480
      TabIndex        =   9
      Top             =   6600
      Width           =   975
   End
   Begin VB.Timer tmrDetect 
      Interval        =   1
      Left            =   6600
      Top             =   1440
   End
   Begin VB.PictureBox picOneHundred 
      Height          =   1455
      Left            =   7680
      Picture         =   "frmCash.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   3195
      TabIndex        =   5
      Top             =   6120
      Width           =   3255
   End
   Begin VB.PictureBox picFive 
      Height          =   1335
      Left            =   7680
      Picture         =   "frmCash.frx":1E8D
      ScaleHeight     =   1275
      ScaleWidth      =   3195
      TabIndex        =   4
      Top             =   1800
      Width           =   3255
   End
   Begin VB.PictureBox picTen 
      Height          =   1335
      Left            =   7680
      Picture         =   "frmCash.frx":400E
      ScaleHeight     =   1275
      ScaleWidth      =   2835
      TabIndex        =   3
      Top             =   3240
      Width           =   2895
   End
   Begin VB.PictureBox picOne 
      Height          =   1215
      Left            =   7680
      OLEDropMode     =   1  'Manual
      Picture         =   "frmCash.frx":6142
      ScaleHeight     =   1155
      ScaleWidth      =   2235
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox picTwenty 
      Height          =   1335
      Left            =   7680
      Picture         =   "frmCash.frx":7B66
      ScaleHeight     =   1275
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Frame frmCash 
      Caption         =   "Select Cash:"
      Height          =   7455
      Left            =   7560
      TabIndex        =   6
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label lblTotall 
      Caption         =   "Your Total:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   11
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblTotal 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   10
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblYouPaid 
      BackStyle       =   0  'Transparent
      Caption         =   "You Paid:"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label lblCurrent 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label lblPlace 
      Caption         =   "Place Cash Here:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Shape shpCash 
      BorderWidth     =   5
      FillColor       =   &H00FF0000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   7455
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   6015
   End
   Begin VB.Menu payment 
      Caption         =   "Go Back To Payment"
   End
End
Attribute VB_Name = "frmCashForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
    frmPay.Show
    frmPay.cmdReceipt.Visible = True
    Unload frmCashForm
End Sub

Private Sub cmdDone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If dblPaidPrice < dblGrandTotal Then
        cmdDone.Enabled = False
        MsgBox "Not Enough Cash"
    End If
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move X, Y
End Sub

Private Sub Form_Load()
    lblTotal.Caption = "$ " & FormatNumber(dblGrandTotal, 2)
    dblPaidPrice = 0
End Sub

Private Sub payment_Click()
    Unload frmCashForm
    frmPay.Show
End Sub

Private Sub picOne_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picOne.Drag vbBeginDrag
End Sub

Private Sub picFive_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picFive.Drag vbBeginDrag
End Sub

Private Sub picTen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picTen.Drag vbBeginDrag
End Sub

Private Sub picTwenty_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picTwenty.Drag vbBeginDrag
End Sub

Private Sub picOneHundred_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picOneHundred.Drag vbBeginDrag
End Sub

Private Sub tmrDetect_Timer()
    If dblPaidPrice >= dblGrandTotal Then
        cmdDone.Enabled = True
    End If
    
    If picOne.Left <> 7680 Or picOne.Top <> 600 Then
        Sleep (500)
        picOne.Drag vbEndDrag
        picOne.Move 7680, 600
        dblPaidPrice = dblPaidPrice + 1
        lblCurrent.Caption = "$ " & dblPaidPrice
    End If
    
    If picFive.Left <> 7680 Or picFive.Top <> 1800 Then
        Sleep (500)
        picFive.Drag vbEndDrag
        picFive.Move 7680, 1800
        dblPaidPrice = dblPaidPrice + 5
        lblCurrent.Caption = "$ " & dblPaidPrice
    End If
    
    If picTen.Left <> 7680 Or picTen.Top <> 3240 Then
        Sleep (500)
        picTen.Drag vbEndDrag
        picTen.Move 7680, 3240
        dblPaidPrice = dblPaidPrice + 10
        lblCurrent.Caption = "$ " & dblPaidPrice
    End If
    
    If picTwenty.Left <> 7680 Or picTwenty.Top <> 4680 Then
        Sleep (500)
        picTwenty.Drag vbEndDrag
        picTwenty.Move 7680, 4680
        dblPaidPrice = dblPaidPrice + 20
        lblCurrent.Caption = "$ " & dblPaidPrice
    End If
    
    If picOneHundred.Left <> 7680 Or picOneHundred.Top <> 6120 Then
        Sleep (500)
        picOneHundred.Drag vbEndDrag
        picOneHundred.Move 7680, 6120
        dblPaidPrice = dblPaidPrice + 100
        lblCurrent.Caption = "$ " & dblPaidPrice
    End If
End Sub
