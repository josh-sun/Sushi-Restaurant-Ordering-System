VERSION 5.00
Begin VB.Form frmStartUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome"
   ClientHeight    =   9960
   ClientLeft      =   5115
   ClientTop       =   2760
   ClientWidth     =   16605
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmStartUp.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   9960
   ScaleWidth      =   16605
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7920
      Width           =   2655
   End
   Begin VB.Timer tmrTime 
      Interval        =   1
      Left            =   14040
      Top             =   120
   End
   Begin VB.CommandButton cmdJap 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7920
      Width           =   3015
   End
   Begin VB.CommandButton cmdEng 
      BackColor       =   &H00FFFFFF&
      Caption         =   "English"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7920
      Width           =   3015
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.Image imgShop3 
      Height          =   960
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   4140
   End
   Begin VB.Image imgShop2 
      Height          =   1005
      Left            =   9600
      Top             =   240
      Width           =   2955
   End
   Begin VB.Image imgShop1 
      Height          =   1020
      Left            =   3600
      Top             =   240
      Width           =   6060
   End
   Begin VB.Image imgBack 
      Height          =   8775
      Left            =   -480
      Stretch         =   -1  'True
      Top             =   720
      Width           =   16575
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEng_Click()
    frmMainMenu.Show
    Unload frmStartUp
End Sub

Private Sub cmdJap_Click()
    MsgBox "Japanese Is Currently Unavailable!"
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub Form_Load()
    Dim lR As Long
    lR = SetTopMostWindow(frmOrder.hwnd, True)
    frmStartUp.Picture = LoadPicture(CurDir & "/Picture/Background.jpg")
    imgBack.Picture = LoadPicture(CurDir & "/Picture/sushilogo.jpg")
    cmdJap.Picture = LoadPicture(CurDir & "/Picture/japan.jpg")
    imgShop1.Picture = LoadPicture(CurDir & "/Picture/Shop1.jpeg")
    imgShop2.Picture = LoadPicture(CurDir & "/Picture/shop2.jpg")
    imgShop3.Picture = LoadPicture(CurDir & "/Picture/Abe.jpg")
    frmOrder.Show
End Sub


Private Sub tmrTime_Timer()
    lblTime.Caption = Now
End Sub
