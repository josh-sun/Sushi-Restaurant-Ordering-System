VERSION 5.00
Begin VB.Form frmCreditCard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pay by Creditcard"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   7155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSecondMove 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   360
   End
   Begin VB.Timer tmrFirstMove 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame fraPaid 
      Caption         =   "You Paid"
      Height          =   2895
      Left            =   240
      TabIndex        =   12
      Top             =   3600
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton cmdConfirm2 
         Caption         =   "Confirm Payment"
         Height          =   615
         Left            =   1200
         TabIndex        =   14
         Top             =   2040
         Width           =   3855
      End
      Begin VB.Label lblGrandTotal 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   360
         TabIndex        =   13
         Top             =   480
         Width           =   6135
      End
   End
   Begin VB.Frame fraCard 
      Height          =   3135
      Left            =   360
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Label lblSlash 
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   19
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblExpDay 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5280
         TabIndex        =   15
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label lblExpMonth 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   11
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label lblCardNumber 
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   6135
      End
      Begin VB.Image imgLogo 
         Height          =   1095
         Left            =   4440
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1815
      End
      Begin VB.Shape shpCard 
         BackColor       =   &H00FFFF80&
         BackStyle       =   1  'Opaque
         Height          =   2655
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.Frame fraInformation 
      Caption         =   "Card Information"
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtNumber4 
         Height          =   285
         Left            =   4560
         MaxLength       =   4
         TabIndex        =   18
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtNumber3 
         Height          =   285
         Left            =   3480
         MaxLength       =   4
         TabIndex        =   17
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtNumber2 
         Height          =   285
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   16
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdConfirm 
         Caption         =   "Confirm Information"
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   2280
         Width           =   1695
      End
      Begin VB.ComboBox cboDay 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1680
         Width           =   735
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtNumber1 
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         ItemData        =   "frmCreditCard.frx":0000
         Left            =   840
         List            =   "frmCreditCard.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblExp 
         Caption         =   "Expirary Date:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Card Number:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblType 
         Caption         =   "Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Image Image1 
      Height          =   6375
      Left            =   7200
      Picture         =   "frmCreditCard.frx":0020
      Stretch         =   -1  'True
      Top             =   240
      Width           =   6975
   End
   Begin VB.Menu MainMenu 
      Caption         =   "Go Back To MainMenu"
   End
End
Attribute VB_Name = "frmCreditCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConfirm_Click()
    If cboType.Text = "" Then
        MsgBox "Please Select A Card Type"
    ElseIf Len(txtNumber1.Text) <> 4 Or NumbersOnly(txtNumber1.Text) = 1 Or _
    Len(txtNumber2.Text) <> 4 Or NumbersOnly(txtNumber2.Text) = 1 Or _
    Len(txtNumber3.Text) <> 4 Or NumbersOnly(txtNumber3.Text) = 1 Or _
    Len(txtNumber4.Text) <> 4 Or NumbersOnly(txtNumber4.Text) = 1 Then
        MsgBox "Invalid Credit Card Number"
    ElseIf cboMonth.Text = "" Or cboDay.Text = "" Then
        MsgBox "Please Input Expirairy Date"
    Else
        If cboType.Text = "Visa" Then
            imgLogo.Picture = LoadPicture(CurDir & "/Picture/visa.jpg")
        Else
            imgLogo.Picture = LoadPicture(CurDir & "/Picture/mc.jpg")
        End If
        lblCardNumber.Caption = txtNumber1.Text & "    " & txtNumber2.Text & "    " & txtNumber3.Text & "    " & txtNumber4.Text
        lblExpMonth.Caption = cboMonth.Text
        lblExpDay.Caption = cboDay.Text
        fraInformation.Visible = False
        fraCard.Visible = True
        frmCreditCard.Width = 15000
        MainMenu.Enabled = False
        tmrFirstMove.Enabled = True
    End If
End Sub

Private Sub cmdConfirm2_Click()
    Unload frmCreditCard
    frmPay.Show
    frmPay.cmdReceipt.Visible = True
End Sub

Private Sub Form_Load()
    Dim intCounter As Integer
    MainMenu.Enabled = True
    intCounter = 1
    dblPaidPrice = dblGrandTotal
    lblGrandTotal.Caption = "$ " & FormatNumber(dblPaidPrice, 2)
    Do While intCounter <> 13
        cboMonth.AddItem intCounter
        cboDay.AddItem intCounter
        intCounter = intCounter + 1
    Loop
End Sub

Private Sub MainMenu_Click()
    Unload frmCreditCard
    frmMainMenu.Show
End Sub

Private Sub tmrFirstMove_Timer()
    If fraCard.Top < 2040 Then
        fraCard.Move 480, fraCard.Top + 100
    Else
        tmrFirstMove.Enabled = False
        tmrSecondMove.Enabled = True
    End If
End Sub

Private Sub tmrSecondMove_Timer()
    If fraCard.Left < 7560 Then
        fraCard.Move fraCard.Left + 100
    Else
        tmrSecondMove.Enabled = False
        fraPaid.Visible = True
    End If
End Sub
