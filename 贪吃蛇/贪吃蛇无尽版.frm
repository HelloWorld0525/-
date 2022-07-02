VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11910
   ClientLeft      =   3060
   ClientTop       =   2025
   ClientWidth     =   21270
   LinkTopic       =   "Form1"
   ScaleHeight     =   11910
   ScaleWidth      =   21270
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   2160
      Top             =   4560
   End
   Begin VB.CommandButton Command3 
      Caption         =   "开始"
      Height          =   615
      Left            =   14880
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Enabled         =   0   'False
      Height          =   495
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   20
      X1              =   0
      X2              =   1920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   19
      X1              =   0
      X2              =   1920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   18
      X1              =   0
      X2              =   1920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   17
      X1              =   0
      X2              =   1920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   16
      X1              =   0
      X2              =   1920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   15
      X1              =   0
      X2              =   1920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   14
      X1              =   0
      X2              =   1920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   13
      X1              =   0
      X2              =   1920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   12
      X1              =   0
      X2              =   1920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   11
      X1              =   0
      X2              =   1920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   10
      X1              =   0
      X2              =   1920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   9
      X1              =   0
      X2              =   1920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   8
      X1              =   0
      X2              =   1920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   7
      X1              =   0
      X2              =   1920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   6
      X1              =   0
      X2              =   1920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   5
      X1              =   0
      X2              =   600
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   4
      X1              =   0
      X2              =   600
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   3
      X1              =   0
      X2              =   600
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   2
      X1              =   0
      X2              =   600
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   600
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   3240
      X2              =   5160
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "我的得分："
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   2010
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As Boolean
Private s As Boolean
Private w As Boolean
Private d As Boolean

Private Sub Command1_GotFocus(Index As Integer)
Command3.SetFocus
End Sub

Private Sub Command3_Click()
Timer1.Enabled = Not Timer1
Command3.Caption = IIf(Timer1, "暂停", "继续")
End Sub

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 87 And Not s Then
 w = True
 a = False
 d = False
ElseIf KeyCode = 83 And Not w Then
 s = True
 a = False
 d = False
ElseIf KeyCode = 65 And Not d Then
 a = True
 s = False
 w = False
ElseIf KeyCode = 68 And Not a Then
 d = True
 s = False
 w = False
End If
End Sub

Private Sub Form_Activate()
Command2.Left = Int(Rnd * Form1.ScaleWidth)
Command2.Top = Int(Rnd * Form1.ScaleHeight)

Command1(0).Left = Form1.ScaleWidth \ 2
Command1(0).Top = Form1.ScaleHeight \ 2
For i = 1 To Command1.UBound
 Command1(i).BackColor = RGB(i / Command1.UBound * 200 + 55, i / Command1.UBound * 200 + 55, i / Command1.UBound * 200 + 55)
 Command1(i).Left = Command1(i - 1).Left - Command1(i).Width
  Command1(i).Top = Command1(i - 1).Top
Next
End Sub

Private Sub Form_Load()
Randomize

For i = 0 To Line1.UBound
 Line1(i).Visible = False
Next i

d = True
Rem 等级初始化
dj = -1
Do
dj = Val(InputBox("0:我是新手" & Chr(10) & "1:我一定经验" & Chr(10) & "2:我经验丰富" & Chr(10) & "3:我是大神", "根据熟练度选择等级", "1"))
Loop Until dj >= 0 And dj <= 3
Select Case dj
 Case 0
  Timer1.Interval = 100
 Case 1
  Timer1.Interval = 75
 Case 2
  Timer1.Interval = 50
 Case 3
  Timer1.Interval = 25
End Select
Form1.WindowState = 2

End Sub

Private Sub Timer1_Timer()
Static defen As Long
Static r As Integer, g As Integer, b As Integer
Randomize

Rem 移动
For i = Command1.UBound To 1 Step -1
Command1(i).Left = Command1(i - 1).Left
Command1(i).Top = Command1(i - 1).Top
Next i
If w Then Command1(0).Top = Command1(0).Top - Command1(0).Height
If s Then Command1(0).Top = Command1(0).Top + Command1(0).Height
If a Then Command1(0).Left = Command1(0).Left - Command1(0).Width
If d Then Command1(0).Left = Command1(0).Left + Command1(0).Width
Rem 死忙
If Command1(0).Top <= 0 Or Command1(0).Top + Command1(0).Height >= Form1.ScaleHeight Or Command1(0).Left <= 0 Or Command1(0).Left + Command1(0).Width >= Form1.ScaleWidth Then

 For i = 0 To Line1.UBound
  Line1(i).Visible = True
  Line1(i).X1 = Command1(0).Left: Line1(i).Y1 = Command1(0).Top
  Line1(i).X2 = Line1(i).X1 + Rnd * 2001 - 1000: Line1(i).Y2 = Line1(i).Y1 + Rnd * 2001 - 1000
 Next i
 MsgBox "你碰壁自杀了", , "一首凉凉送给你": End
End If

For i = 1 To Command1.UBound
 If Command1(0).Top = Command1(i).Top And Command1(0).Left = Command1(i).Left Then
  For j = 0 To Line1.UBound
   Line1(j).Visible = True
   Line1(j).X1 = Command1(0).Left: Line1(j).Y1 = Command1(0).Top
   Line1(j).X2 = Line1(j).X1 + Rnd * (2001 - 1000): Line1(j).Y2 = Line1(j).Y1 + Rnd * (2000 - 1000)
  Next j
 MsgBox "你把自己咬死了", , "一首凉凉送给你": End
 End If
Next i
Rem 成长
If ((Command1(0).Left >= Command2.Left And Command1(0).Left <= Command2.Left + Command2.Width) Or (Command1(0).Left + Command1(0).Width >= Command2.Left And Command1(0).Left + Command1(0).Width <= Command2.Left + Command2.Width)) And ((Command1(0).Top >= Command2.Top And Command1(0).Top <= Command2.Top + Command2.Height) Or (Command1(0).Top + Command1(0).Height >= Command2.Top And Command1(0).Top + Command1(0).Height <= Command2.Top + Command2.Height)) Then
 Do
  Command2.Top = Rnd * Form1.ScaleHeight
  Command2.Left = Rnd * Form1.ScaleWidth
 Loop Until Command2.Top + Command2.Height < Form1.ScaleHeight And Command2.Left + Command2.Width < Form1.ScaleWidth
  Command2.BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
  Load Command1(Command1.UBound + 1)
  Command1(Command1.UBound).Visible = True
  Command1(Command1.UBound).Left = Command1(Command1.UBound - 1).Left
  Command1(Command1.UBound).Top = Command1(Command1.UBound - 1).Top
  For i = 1 To Command1.UBound
   Command1(i).BackColor = RGB(i / Command1.UBound * 200 + 55, i / Command1.UBound * 200 + 55, i / Command1.UBound * 200 + 55)
  Next i
  defen = defen + 10
  Label1.Caption = "我的得分：" & defen
End If

End Sub
