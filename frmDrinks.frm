VERSION 5.00
Begin VB.Form frmDrinks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Drinks"
   ClientHeight    =   8940
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   6630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0000C000&
      Caption         =   "Cancel Order"
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
      Height          =   1575
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FF0000&
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
      Height          =   1575
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H000000FF&
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
      Height          =   1575
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Frame fraSelection 
      Caption         =   "Select Your Drink:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      Begin VB.CommandButton cmdPepsi 
         Height          =   1335
         Left            =   4080
         Picture         =   "frmDrinks.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmd7Up 
         Height          =   1335
         Left            =   2160
         Picture         =   "frmDrinks.frx":1562
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdJapAlh 
         Height          =   1335
         Left            =   240
         Picture         =   "frmDrinks.frx":255D
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label lblPrice 
      BackStyle       =   0  'Transparent
      Caption         =   "$1.99 For All Drinks"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   6600
      Width           =   4695
   End
   Begin VB.Shape shpDrink 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape shpCupOutline 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      FillColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   720
      Top             =   3960
      Width           =   1215
   End
End
Attribute VB_Name = "frmDrinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intFlag As Integer
Private Const intHeight As Integer = 855
Private Const intWidth As Integer = 135

Private Sub DrinkAnimation()
    cmdDone.Enabled = False
    cmdJapAlh.Enabled = False
    cmdPepsi.Enabled = False
    cmd7Up.Enabled = False
    Do
        shpDrink.Move shpDrink.Left - 35, shpDrink.Top + 100
        shpDrink.Height = shpDrink.Height + 90
        shpDrink.Width = shpDrink.Width + 60
        Sleep (100)
    Loop While shpDrink.Top < 3800
    cmdAdd.Enabled = True
    cmdCancel.Enabled = True
    cmdDone.Enabled = True
End Sub

Private Sub cmd7Up_Click()
    shpDrink.Move 3380
    shpCupOutline.Move 2760
    shpDrink.FillColor = vbGreen
    shpDrink.Visible = True
    strItemName = "Seven Up"
    Call DrinkAnimation
End Sub

Private Sub cmdCancel_Click()
    cmdAdd.Enabled = False
    cmdCancel.Enabled = False
    shpDrink.Visible = False
    shpDrink.Height = intHeight
    shpDrink.Width = intWidth
    shpDrink.Move 1320, 2160
    cmdJapAlh.Enabled = True
    cmdPepsi.Enabled = True
    cmd7Up.Enabled = True
End Sub

Private Sub cmdJapAlh_Click()
    shpDrink.Move 1340
    shpCupOutline.Move 720
    shpDrink.FillColor = vbCyan
    shpDrink.Visible = True
    strItemName = "Japanese Shochu"
    Call DrinkAnimation
End Sub

Private Sub cmdAdd_Click()
    shpDrink.Visible = False
    shpDrink.Height = intHeight
    shpDrink.Width = intWidth
    shpDrink.Move 1320, 2160
    cmdJapAlh.Enabled = True
    cmdPepsi.Enabled = True
    cmd7Up.Enabled = True
    ReDim Preserve strItemStore(X), dblPriceStore(X), intQuantityStore(X)
    frmOrder.txtOrder.Text = frmOrder.txtOrder.Text & strItemName & " X " & 1 & vbNewLine
    strItemStore(X) = strItemName
    dblPriceStore(X) = 2.99
    intQuantityStore(X) = 1
    MsgBox "Your Orders Have Been Added"
    X = X + 1
    cmdCancel.Enabled = False
    cmdAdd.Enabled = False
End Sub

Private Sub cmdPepsi_Click()
    shpDrink.Move 5290
    shpCupOutline.Move 4665
    shpDrink.FillColor = vbBlack
    shpDrink.Visible = True
    strItemName = "Pepsi"
    Call DrinkAnimation
End Sub

Private Sub cmdDone_Click()
    Unload frmDrinks
    frmMainMenu.Show
End Sub

Private Sub Form_Load()
    frmDrinks.Picture = LoadPicture(CurDir & "/Picture/drinkback.jpg")
End Sub
