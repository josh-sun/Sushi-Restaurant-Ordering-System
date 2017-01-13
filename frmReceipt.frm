VERSION 5.00
Begin VB.Form frmReceipt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receipt"
   ClientHeight    =   8415
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   6375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReceipt 
      Caption         =   "Print Receipt"
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   7320
      Width           =   6015
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Finish Ordering"
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   6360
      Width           =   6015
   End
   Begin VB.Frame fraPayment 
      Caption         =   "Credit Card Info: (If Applicable)"
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   6015
      Begin VB.Label lblExp 
         Caption         =   "Expirary Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   5655
      End
      Begin VB.Label lblType 
         Caption         =   "Type:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   5655
      End
      Begin VB.Label lblCard 
         Caption         =   "Card Number:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   5535
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Order Information:"
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.Label lblChange 
         Caption         =   "Change:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3720
         Width           =   5415
      End
      Begin VB.Label lblPaid 
         Caption         =   "Amount Paid:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   3240
         Width           =   5535
      End
      Begin VB.Label lblStore 
         Caption         =   "Store Location:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2760
         Width           =   5655
      End
      Begin VB.Label lblPhone 
         Caption         =   "Phone # :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   5535
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   5655
      End
      Begin VB.Label lblAddress 
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   5655
      End
      Begin VB.Label lblName 
         Caption         =   "Customer Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   5655
      End
      Begin VB.Label lblDelivery 
         Caption         =   "Delivery Type:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   5535
      End
   End
End
Attribute VB_Name = "frmReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
    Unload frmReceipt
    Unload frmPay
    Unload frmOrder
    frmStartUp.Show
End Sub

Private Sub cmdReceipt_Click()
    MsgBox "Printer Is Broken"
End Sub

Private Sub Form_Load()
    If frmPay.optDelivery.Value = True Then
        lblDelivery.Caption = lblDelivery.Caption & " Delivery"
    Else
        lblDelivery.Caption = lblDelivery.Caption & " In Store Pick-Up"
    End If
    
    frmOrder.Show
    frmOrder.cmdClear.Enabled = False
    lblName.Caption = lblName.Caption & frmPay.txtFirst.Text & " " & frmPay.txtLast.Text
    lblAddress.Caption = lblAddress.Caption & frmPay.txtAddress.Text & ", " & frmPay.cboProvince.Text
    lblEmail.Caption = lblEmail.Caption & frmPay.txtEmail.Text
    lblPhone.Caption = lblPhone.Caption & frmPay.txtPhone1.Text & frmPay.txtPhone2 & frmPay.txtPhone3
    lblStore.Caption = lblStore.Caption & frmPay.cboStore.Text
    lblPaid.Caption = lblPaid.Caption & "$ " & FormatNumber(dblPaidPrice, 2)
    lblChange.Caption = "$ " & FormatNumber(dblPaidPrice - dblGrandTotal, 2)
    lblCard.Caption = lblCard.Caption & frmCreditCard.lblCardNumber.Caption
    lblType.Caption = lblType.Caption & frmCreditCard.cboType.Text
    lblExp.Caption = lblExp.Caption & frmCreditCard.lblExpMonth.Caption & "/" & frmCreditCard.lblExpDay.Caption
End Sub
