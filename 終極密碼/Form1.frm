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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�رdPOP1��W5"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   10
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "�A���@��"
      BeginProperty Font 
         Name            =   "�رdPOP1��W5"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      Style           =   1  '�Ϥ��~�[
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
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '��u�T�w
      Height          =   615
      Left            =   1680
      TabIndex        =   7
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label7 
      Alignment       =   2  '�m�����
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '��u�T�w
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
      Alignment       =   2  '�m�����
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '��u�T�w
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
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "�A��J����"
      BeginProperty Font 
         Name            =   "�رd�ʶ���"
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
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "�Ҳq����"
      BeginProperty Font 
         Name            =   "�رd�ʶ���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '��u�T�w
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
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "�׷��K�X"
      BeginProperty Font 
         Name            =   "�رd�ʶ���"
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
      Alignment       =   2  '�m�����
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '��u�T�w
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
      a = InputBox("�п�J�q���K�X", "�׷��K�X")
    If a >= Label1.Caption And a <= Label3.Caption Then
   Exit Do
    End If
   Loop
  k = k + 1
   Label7.Caption = a
   Label6.Caption = k
If pw = a Then
   Label8.Caption = "����!�q���F~"
   Exit Do
Else
   If pw > a Then
   Label1.Caption = a
   Label8.Caption = "�Ʀr�A�j�@�I"
Else
   Label3.Caption = a
   Label8.Caption = "�Ʀr�A�p�@�I"
   End If
End If
Loop

Select Case k
   Case 1
      MsgBox "�A�ӯ��F~~~", 0, "�����e��"
    
   Case 2 To 4
      MsgBox "���B�|���{�b�A���W", 0, "�����e��"
   Case 5 To 7
      MsgBox "��𴶳qo___o", 0, "�����e��"
   Case Else
      MsgBox "�ӻ��O!!�b�[�o�@~", 0, "�����e��"
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

