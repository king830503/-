VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   4395
   ClientTop       =   5910
   ClientWidth     =   5970
   FillColor       =   &H00FFC0C0&
   ForeColor       =   &H00FFC0C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   5970
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "結束"
      BeginProperty Font 
         Name            =   "華康POP1體W5"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "再玩一次"
      BeginProperty Font 
         Name            =   "華康POP1體W5"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '單線固定
      Height          =   615
      Left            =   1680
      TabIndex        =   7
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label7 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   2280
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "One Stroke Script LET"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "你輸入的數"
      BeginProperty Font 
         Name            =   "華康粗圓體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "所猜次數"
      BeginProperty Font 
         Name            =   "華康粗圓體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Orange LET"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   735
      Left            =   4560
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "終極密碼"
      BeginProperty Font 
         Name            =   "華康粗圓體"
         Size            =   24
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Orange LET"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Randomize Timer
Dim pw, k As Integer
Dim a As Integer
pw = Int(Rnd * 98) + 1

Do
   Do
      a = InputBox("請輸入通關密碼", "終極密碼")
    If a >= Label1.Caption And a <= Label3.Caption Then
   Exit Do
    End If
   Loop
  k = k + 1
   Label7.Caption = a
   Label6.Caption = k
If pw = a Then
   Label8.Caption = "恭喜!猜中了~"
   Exit Do
Else
   If pw > a Then
   Label1.Caption = a
   Label8.Caption = "數字再大一點"
Else
   Label3.Caption = a
   Label8.Caption = "數字再小一點"
   End If
End If
Loop

Select Case k
   Case 1
      MsgBox "你太神了~~~", 0, "結束畫面"
    
   Case 2 To 4
      MsgBox "幸運會降臨在你身上", 0, "結束畫面"
   Case 5 To 7
      MsgBox "手氣普通o___o", 0, "結束畫面"
   Case Else
      MsgBox "太遜摟!!在加油哦~", 0, "結束畫面"
End Select

End Sub

Private Sub Command2_Click()
Label8.Caption = ""
Label7.Caption = ""
Label1.Caption = 1
Label3.Caption = 100
Label1.Caption = "0"
End Sub

Private Sub Command3_Click()
End
End Sub

