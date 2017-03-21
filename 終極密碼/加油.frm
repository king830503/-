VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   5265
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4560
      Top             =   3720
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   3
      Left            =   2400
      Picture         =   "加油.frx":0000
      Stretch         =   -1  'True
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1245
      Index           =   2
      Left            =   240
      Picture         =   "加油.frx":AD23
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   1
      Left            =   2400
      Picture         =   "加油.frx":1BD46
      Stretch         =   -1  'True
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1560
      Index           =   0
      Left            =   0
      Picture         =   "加油.frx":29F12
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   1905
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As Integer

Private Sub Timer1_Timer()

Image1(t Mod 4).Visible = False

Image1((t + 1) Mod 4).Visible = True

t = t + 1
End Sub

