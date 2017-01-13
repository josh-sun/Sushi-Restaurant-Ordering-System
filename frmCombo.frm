VERSION 5.00
Begin VB.Form frmCombo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Combos"
   ClientHeight    =   8970
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   10410
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   10410
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrEnable 
      Interval        =   1
      Left            =   9960
      Top             =   0
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Return To Main Menu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      TabIndex        =   14
      Top             =   7200
      Width           =   3255
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add To Order"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      TabIndex        =   13
      Top             =   6120
      Width           =   3255
   End
   Begin VB.Frame frmChoices 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Combo And Specials"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9855
      Begin VB.ComboBox cboSashimiSize 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmCombo.frx":0000
         Left            =   6360
         List            =   "frmCombo.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2640
         Width           =   3135
      End
      Begin VB.ComboBox cboMakiSize 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmCombo.frx":0042
         Left            =   6360
         List            =   "frmCombo.frx":004F
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1560
         Width           =   3135
      End
      Begin VB.ComboBox cboSushiSize 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmCombo.frx":0084
         Left            =   6360
         List            =   "frmCombo.frx":0091
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   3135
      End
      Begin VB.CheckBox chkSashimi 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Assorted Sashimi Boat"
         Height          =   735
         Left            =   240
         TabIndex        =   3
         Top             =   2400
         Width           =   3735
      End
      Begin VB.CheckBox chkMaki 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Assorted Maki Platter"
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   3495
      End
      Begin VB.CheckBox chkSushi 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Assorted Sushi Platter"
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Select Size:"
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
         Left            =   5040
         TabIndex        =   11
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Select Size:"
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
         Left            =   5040
         TabIndex        =   9
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Select Size:"
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
         Left            =   5040
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Salmon Sashimi, Tuna Sashimi, Ika Sashimi, Mackerel Sashimi"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3120
         Width           =   5415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Kappa Maki, Tekkamaki, California Rolls, Futomaki, Tsunamayo Maki"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   5535
      End
      Begin VB.Label lblSushi 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Unagi, Sake Nigiri, Ikura Gukan, Ebi Nigiri, Tamagoyaki"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   5415
      End
   End
   Begin VB.Image imgComboPic 
      Height          =   4335
      Left            =   240
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   6015
   End
End
Attribute VB_Name = "frmCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub PrintOnReceipt()
    frmOrder.txtOrder.Text = frmOrder.txtOrder.Text & strItemName & " X " & 1 & vbNewLine
    ReDim Preserve strItemStore(X), dblPriceStore(X), intQuantityStore(X)
    strItemStore(X) = strItemName
    dblPriceStore(X) = dblItemPrice
    intQuantityStore(X) = 1
    X = X + 1
    MsgBox "Your Order Has Been Added"
End Sub

Private Sub chkMaki_Click()
    If chkMaki.Value = 1 Then
        cboMakiSize.Enabled = True
    Else
        cboMakiSize.Enabled = False
    End If
End Sub

Private Sub chkMaki_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgComboPic.Picture = LoadPicture(CurDir & "/Picture/makicombo.jpg")
End Sub

Private Sub chkSashimi_Click()
    If chkSashimi.Value = 1 Then
        cboSashimiSize.Enabled = True
    Else
        cboSashimiSize.Enabled = False
    End If
End Sub

Private Sub chkSashimi_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgComboPic.Picture = LoadPicture(CurDir & "/Picture/sashimicombo.jpg")
End Sub

Private Sub chkSushi_Click()
    If chkSushi.Value = 1 Then
        cboSushiSize.Enabled = True
    Else
        cboSushiSize.Enabled = False
    End If
End Sub

Private Sub chkSushi_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgComboPic.Picture = LoadPicture(CurDir & "/Picture/sushicombo.jpg")
End Sub

Private Sub cmdAdd_Click()
    
    If chkSushi.Value = 1 Then
        If cboSushiSize.Text = cboSushiSize.List(2) Then
            dblItemPrice = 69.99
            strItemName = "Sushi Combo (Large)"
            Call PrintOnReceipt
        ElseIf cboSushiSize.Text = cboSushiSize.List(1) Then
            dblItemPrice = 49.99
            strItemName = "Sushi Combo (Medium)"
            Call PrintOnReceipt
        ElseIf cboSushiSize.Text = cboSushiSize.List(0) Then
            dblItemPrice = 29.99
            strItemName = "Sushi Combo (Small)"
            Call PrintOnReceipt
        Else
            chkSushi.Value = 0
            cboSushiSize.Enabled = False
        End If
    End If
    
    If chkMaki.Value = 1 Then
        If cboMakiSize.Text = cboMakiSize.List(2) Then
            dblItemPrice = 69.99
            strItemName = "Maki Combo (Large)"
            Call PrintOnReceipt
        ElseIf cboMakiSize.Text = cboMakiSize.List(1) Then
            dblItemPrice = 49.99
            strItemName = "Maki Combo (Medium)"
            Call PrintOnReceipt
        ElseIf cboMakiSize.Text = cboMakiSize.List(0) Then
            dblItemPrice = 29.99
            strItemName = "Maki Combo (Small)"
            Call PrintOnReceipt
        Else
            chkMaki.Value = 0
            cboMakiSize.Enabled = False
        End If
    End If

    If chkSashimi.Value = 1 Then
        If cboSashimiSize.Text = cboSashimiSize.List(2) Then
            dblItemPrice = 69.99
            strItemName = "Sashimi Boat (Large)"
            Call PrintOnReceipt
        ElseIf cboSashimiSize.Text = cboSashimiSize.List(1) Then
            dblItemPrice = 49.99
            strItemName = "Sashimi Boat (Medium)"
            Call PrintOnReceipt
        ElseIf cboSashimiSize.Text = cboSashimiSize.List(0) Then
            dblItemPrice = 29.99
            strItemName = "Sashimi Boat (Small)"
            Call PrintOnReceipt
        Else
            chkSashimi.Value = 0
            cboSashimiSize.Enabled = False
        End If
    End If

    If chkSushi.Value = 0 And chkSashimi.Value = 0 And chkMaki.Value = 0 Then
        MsgBox "Please Place Your Order Before Proceeding"
    ElseIf (cboSashimiSize.Text = "" And chkSashimi.Value = 1) Or (cboMakiSize.Text = "" And chkMaki.Value = 1) Or (cboSushiSize.Text = "" And chkSushi.Value = 1) Then
        MsgBox "Please Select A Size"
    Else
        Unload frmCombo
        frmMainMenu.Show
    End If
End Sub

Private Sub cmdDone_Click()
    Unload frmCombo
    frmMainMenu.Show
End Sub

Private Sub Form_Load()
    frmCombo.Picture = LoadPicture(CurDir & "/Picture/comboback.jpg")
End Sub

Private Sub tmrEnable_Timer()
    If chkSushi.Enabled = True Or chkMaki.Enabled = True Or chkSashimi.Enabled = True Then
        cmdAdd.Enabled = True
    Else
        cmdAdd.Enabled = False
    End If
End Sub
