VERSION 5.00
Begin VB.Form frmPay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment"
   ClientHeight    =   7530
   ClientLeft      =   -15
   ClientTop       =   570
   ClientWidth     =   7575
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrDisable 
      Interval        =   1
      Left            =   7080
      Top             =   0
   End
   Begin VB.Frame frmPayment 
      Caption         =   "Select Payment Method"
      Height          =   3375
      Left            =   3960
      TabIndex        =   12
      Top             =   3840
      Width           =   3255
      Begin VB.CommandButton cmdReceipt 
         Caption         =   "Confirm and Receipt"
         Height          =   1815
         Left            =   360
         TabIndex        =   24
         Top             =   840
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdCreditCard 
         Caption         =   "Credit Card"
         Height          =   735
         Left            =   360
         TabIndex        =   16
         Top             =   1920
         Width           =   2415
      End
      Begin VB.CommandButton cmdCash 
         Caption         =   "Cash"
         Height          =   735
         Left            =   360
         TabIndex        =   15
         Top             =   840
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Store"
      Height          =   3375
      Left            =   360
      TabIndex        =   11
      Top             =   3840
      Width           =   3375
      Begin VB.ComboBox cboStore 
         Height          =   315
         ItemData        =   "frmPay.frx":0000
         Left            =   240
         List            =   "frmPay.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblOrder 
         Caption         =   "Your Order Will Be Ready For Pick-Up in 20 Minutes, Else It Will Be Free Of Charge!"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2880
         Width           =   3015
      End
   End
   Begin VB.Frame frmDelivery 
      Caption         =   "Delivery Or Pick-Up"
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   7095
      Begin VB.OptionButton optPickUp 
         Caption         =   "In Store Pick-Up"
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton optDelivery 
         Caption         =   "Delivery"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame frmPersonal 
      Caption         =   "Personal Information"
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   7095
      Begin VB.TextBox txtPhone3 
         Height          =   285
         Left            =   6000
         MaxLength       =   4
         TabIndex        =   26
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtPhone2 
         Height          =   285
         Left            =   5400
         MaxLength       =   3
         TabIndex        =   25
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtPhone1 
         Height          =   285
         Left            =   4800
         MaxLength       =   3
         TabIndex        =   23
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   4800
         MaxLength       =   25
         TabIndex        =   22
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   4800
         MaxLength       =   25
         TabIndex        =   21
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtCity 
         Height          =   285
         Left            =   4800
         MaxLength       =   20
         TabIndex        =   20
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cboProvince 
         Height          =   315
         ItemData        =   "frmPay.frx":005D
         Left            =   1320
         List            =   "frmPay.frx":006D
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtLast 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   18
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtFirst 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   17
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   27
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Phone # :"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblAddress 
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblCity 
         Caption         =   "City:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblProvince 
         Caption         =   "Province:/"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblLast 
         Caption         =   "Last Name:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblFirst 
         Caption         =   "First Name:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Menu MainMenu 
      Caption         =   "Cancel Payment And Go Back To MainMenu"
      Index           =   1
      NegotiatePosition=   1  'Left
   End
End
Attribute VB_Name = "frmPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intMethod, X As Integer

Private Function StripNumber(stdText As String)
    Dim str As String, i As Integer
     'strips the number from a longer text string
    stdText = Trim(stdText)
     
    For i = 1 To Len(stdText)
        If Not IsNumeric(Mid(stdText, i, 1)) Then
            str = str & Mid(stdText, i, 1)
        End If
    Next i
    StripNumber = str ' * 1
End Function

Private Function Check()
    Dim strCheck, strEmail, strPhone1, strPhone2, strPhone3 As String
    Dim intCounter As Integer
    strPhone1 = txtPhone1.Text
    strPhone2 = txtPhone2.Text
    strPhone3 = txtPhone3.Text
    
    If optDelivery.Value = False And optPickUp.Value = False Then
        MsgBox "Please Either Choose Delivery or In-Store Pick-Up"
    ElseIf txtFirst.Text = "" Then
        MsgBox "Invalid First Name"
    ElseIf txtLast.Text = "" Then
        MsgBox "Invalid Last Name"
    ElseIf cboProvince.Text = "" Then
        MsgBox "Invalid Province"
    ElseIf txtCity.Text = "" Then
        MsgBox "Invalid City Name"
    ElseIf txtAddress.Text = "" Then
        MsgBox "Invalid Address"
    ElseIf txtEmail.Text = "" Then
        MsgBox "Invalid Email"
    ElseIf Len(strPhone1) <> 3 Or NumbersOnly(strPhone1) = 1 Then
        MsgBox "Invalid Phone Number In Field 1"
    ElseIf Len(strPhone2) <> 3 Or NumbersOnly(strPhone2) = 1 Then
        MsgBox "Invalid Phone Number In Field 2"
    ElseIf Len(strPhone3) <> 4 Or NumbersOnly(strPhone3) = 1 Then
        MsgBox "Invalid Phone Number In Field 3"
    ElseIf cboStore.Text = "" Then
        MsgBox "Please Choose A Store Location"
    Else
        strEmail = txtEmail.Text
        intCounter = 0
        Do
            intCounter = intCounter + 1
            strCheck = Left(strEmail, intCounter)
            strCheck = Right(strCheck, 1)
            If strCheck = "@" Then
                intCounter = Len(strEmail)
                If intMethod = 1 Then
                    frmPay.Hide
                    frmCashForm.Show
                ElseIf intMethod = 2 Then
                    frmPay.Hide
                    frmCreditCard.Show
                Else
                    frmPay.Hide
                    frmReceipt.Show
                    Unload frmMainMenu
                End If
            ElseIf intCounter = Len(strEmail) And strCheck <> "@" Then
                MsgBox "Invalid Email"
            End If
        Loop While Len(strEmail) <> intCounter
    End If
End Function

Private Sub cmdCash_Click()
    intMethod = 1
    Call Check
End Sub

Private Sub cmdCreditCard_Click()
    intMethod = 2
    Call Check
End Sub

Private Sub cmdReceipt_Click()
    intMethod = 0
    Call Check
End Sub

Private Sub Form_Load()
    optDelivery.Value = False
    optPickUp.Value = False
End Sub

Private Sub MainMenu_Click(Index As Integer)
    Unload frmPay
    Unload frmCashForm
    Unload frmCreditCard
    frmMainMenu.Show
    frmOrder.Show
End Sub

Private Sub txtCity_Change()
    txtCity.Text = StripNumber(txtCity.Text)
End Sub

Private Sub txtFirst_Change()
    txtFirst.Text = StripNumber(txtFirst.Text)
End Sub

Private Sub txtLast_Change()
    txtLast.Text = StripNumber(txtLast.Text)
End Sub

