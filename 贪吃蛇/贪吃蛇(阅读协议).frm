VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000009&
   Caption         =   "进入游戏前需要阅读协议"
   ClientHeight    =   12060
   ClientLeft      =   10440
   ClientTop       =   1485
   ClientWidth     =   7095
   LinkTopic       =   "Form2"
   ScaleHeight     =   12060
   ScaleWidth      =   7095
   Begin VB.CommandButton Command1 
      Caption         =   "开始游戏"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   9
      Top             =   9960
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000013&
      Caption         =   "我已详细阅读并确认以上协议"
      ForeColor       =   &H80000011&
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   10800
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   6120
      Top             =   2640
   End
   Begin VB.Label Label8 
      Caption         =   "最大长度：100"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   8400
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "初始长度为：0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "游戏设置："
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "我们的目标是：要做永远的NO1."
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   4
      Top             =   4680
      Width           =   5415
   End
   Begin VB.Label Label4 
      Caption         =   "史上最强，计应19-2班"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   3
      Top             =   3120
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "故城职教中心"
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "注意"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   4215
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Command1.Enabled = Not Command1.Enabled
End Sub

Private Sub Command1_Click()
Form1.Show
Unload Form2

End Sub

Private Sub Form_Load()
Label1.BackStyle = 0: Label2.BackStyle = 0: Label3.BackStyle = 0: Label4.BackStyle = 0: Label5.BackStyle = 0
Label1.Caption = "职教人," & Chr(10) & "职教魂," & Chr(10) & "职教都是人上人"
Label4.Caption = "史上最强，" & Chr(10) & "计应19-2班"
Label1.FontSize = 25
Label1.WordWrap = True
Label1.AutoSize = True
Command1.Enabled = False
End Sub



Private Sub Label7_Click()
Dim initial As Integer
Do
initial = Val(InputBox("初始长度应低于100", "自定义设置初始长度", "0"))
Loop Until initial < 100 And initial >= 0
initialleng = initial
Label7.Caption = "初始长度为：" & initial
End Sub

Private Sub Label8_Click()
Dim max As Long
Do
max = Val(InputBox("最大长度应不低于100", "自定义设置最大长度", "100"))
Loop Until max >= 100
maxleng = max
Label8.Caption = "最大长度：" & max
End Sub

Private Sub Timer1_Timer()
Randomize
Label1.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Label2.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Label3.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Label4.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Label5.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub
