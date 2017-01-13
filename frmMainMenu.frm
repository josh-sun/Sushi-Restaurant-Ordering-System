VERSION 5.00
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   9930
   ClientLeft      =   4920
   ClientTop       =   2565
   ClientWidth     =   16605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmMainMenu.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   9930
   ScaleMode       =   0  'User
   ScaleWidth      =   16605
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdDrinks 
      BackColor       =   &H80000009&
      Height          =   3135
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   4815
   End
   Begin VB.CommandButton cmdCheckOut 
      BackColor       =   &H8000000B&
      Caption         =   "Checkout"
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   14280
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.Timer tmrTextColor 
      Interval        =   250
      Left            =   15000
      Top             =   120
   End
   Begin VB.CommandButton cmdSashimi 
      BackColor       =   &H80000009&
      Height          =   3135
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   4815
   End
   Begin VB.CommandButton cmdCombo 
      BackColor       =   &H80000009&
      Height          =   3135
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   4815
   End
   Begin VB.CommandButton cmdSushi 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   4815
   End
   Begin VB.Label lblMenuTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   99.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheckOut_Click()
    frmMainMenu.Hide
    frmCheckout.Show
End Sub

Private Sub cmdCombo_Click()
    frmCombo.Show
    frmMainMenu.Hide
End Sub

Private Sub cmdCombo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdCombo.Picture = LoadPicture(CurDir & "/Picture/combos.jpg")
End Sub

Private Sub cmdDrinks_Click()
    frmDrinks.Show
    frmMainMenu.Hide
End Sub

Private Sub cmdDrinks_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdDrinks.Picture = LoadPicture(CurDir & "/Picture/drinkstitle.jpg")
End Sub

Private Sub cmdSashimi_Click()
    frmSashimi.Show
    frmMainMenu.Hide
End Sub

Private Sub cmdSashimi_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdSashimi.Picture = LoadPicture(CurDir & "/Picture/sashimititle.jpg")
End Sub

Private Sub cmdSushi_Click()
    frmSushiOrder.Show
    frmMainMenu.Hide
End Sub

Private Sub cmdSushi_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdSushi.Picture = LoadPicture(CurDir & "/Picture/sushititle.jpg")
End Sub

Private Sub Form_Load()
    cmdSushi.Picture = LoadPicture(CurDir & "/Picture/sushi.jpg")
    cmdSashimi.Picture = LoadPicture(CurDir & "/Picture/sashimi.jpg")
    cmdCombo.Picture = LoadPicture(CurDir & "/Picture/combo.jpg")
    cmdDrinks.Picture = LoadPicture(CurDir & "/Picture/drinks.jpg")
    frmMainMenu.Picture = LoadPicture(CurDir & "/Picture/menu.jpg")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdSushi.Picture = LoadPicture(CurDir & "/Picture/sushi.jpg")
    cmdSashimi.Picture = LoadPicture(CurDir & "/Picture/sashimi.jpg")
    cmdCombo.Picture = LoadPicture(CurDir & "/Picture/combo.jpg")
    cmdDrinks.Picture = LoadPicture(CurDir & "/Picture/drinks.jpg")
End Sub

Private Sub tmrTextColor_Timer()
    If lblMenuTitle.ForeColor = vbWhite Then
        lblMenuTitle.ForeColor = vbRed
    ElseIf lblMenuTitle.ForeColor = vbRed Then
        lblMenuTitle.ForeColor = vbBlue
    ElseIf lblMenuTitle.ForeColor = vbBlue Then
        lblMenuTitle.ForeColor = vbGreen
    ElseIf lblMenuTitle.ForeColor = vbGreen Then
        lblMenuTitle.ForeColor = vbBlack
    Else
        lblMenuTitle.ForeColor = vbWhite
    End If
End Sub
