VERSION 5.00
Begin VB.Form frmOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Your Current Order"
   ClientHeight    =   6495
   ClientLeft      =   14070
   ClientTop       =   3000
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   4560
   Begin VB.Timer tmrAdd 
      Interval        =   1
      Left            =   4320
      Top             =   120
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Order"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   6000
      Width           =   2295
   End
   Begin VB.TextBox txtOrder 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5805
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmOrder.frx":0000
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    txtOrder.Text = "Your Order:" & vbCrLf
    X = 0
    ReDim strItemStore(0), dblPriceStore(0), intQuantityStore(0)
    strInputBox = ""
    dblPrice = 0
    dblItemPrice = 0
    strItemName = ""
    dblTotal = 0
End Sub

